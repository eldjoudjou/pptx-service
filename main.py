"""
PPTX Service — Micro-service de manipulation PowerPoint
Reproduit le workflow Claude pour la manipulation de fichiers PPTX.

Architecture :
  1. Reçoit une demande utilisateur + fichier PPTX (ou demande de création)
  2. Inspecte le fichier (structure, contenu)
  3. Appelle un LLM pour générer du code Python
  4. Exécute le code (édition XML directe ou python-pptx)
  5. Si erreur → renvoie le traceback au LLM → retry
  6. Sauvegarde sur S3 et retourne un lien pré-signé
""

import asyncio
import io
import json
import os
import re
import shutil
import tempfile
import traceback
import uuid
import zipfile
from pathlib import Path

import httpx
from fastapi import FastAPI, UploadFile, File, Form, HTTPException, Request
from fastapi.responses import JSONResponse, StreamingResponse
from pptx import Presentation
from pptx.util import Inches, Pt, Emu, Cm
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR, MSO_AUTO_SIZE
from pptx.enum.shapes import MSO_SHAPE
from lxml import etree
import defusedxml.minidom

# ============================================================
# Configuration
# ============================================================

app = FastAPI(title="PPTX Service", version="1.0.0")

# CORS — permettre les appels depuis Langflow/SiaGPT
from fastapi.middleware.cors import CORSMiddleware
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

# LLM API — SiaGPT /plain_llm endpoint
# Format : POST /chat/plain_llm { systemPrompt, query, llm, temperature }
# Modèles disponibles : claude-4.5-sonnet, claude-4.5-haiku, claude-4-sonnet,
#   gpt-5, gpt-5-mini, gpt-4o, gpt-4.1, o3, o4-mini,
#   gemini-2.5-pro, gemini-3-pro, mistral-large-2, etc.
LLM_API_URL = os.environ.get("LLM_API_URL", "https://backend.siagpt.ai/chat/plain_llm")
LLM_API_KEY = os.environ.get("LLM_API_KEY", "")
LLM_MODEL = os.environ.get("LLM_MODEL", "claude-4.5-sonnet")

# SiaGPT Medias API — stockage des fichiers dans la collection
SIAGPT_MEDIAS_URL = os.environ.get("SIAGPT_MEDIAS_URL", "https://backend.siagpt.ai/medias")
SIAGPT_COLLECTION_ID = os.environ.get("SIAGPT_COLLECTION_ID", "")  # UUID de la collection cible

SYSTEM_PROMPT_PATH = os.environ.get("SYSTEM_PROMPT_PATH", "/app/system_prompt.md")

# Retry
MAX_RETRIES = int(os.environ.get("MAX_RETRIES", "4"))

# ============================================================
# Initialisation
# ============================================================

# Dossier temporaire de travail
Path("/tmp/pptx-work").mkdir(parents=True, exist_ok=True)


# ============================================================
# Fonctions utilitaires — Stockage SiaGPT Medias
# ============================================================

async def save_to_siagpt_medias(data: bytes, filename: str, auth_token: str) -> dict:
    """
    Upload un fichier dans la collection SiaGPT via POST /medias/.
    Retourne les infos du media créé (uuid, name, versions...).
    """
    import json as _json

    media_metadata = _json.dumps({"collectionId": SIAGPT_COLLECTION_ID})

    async with httpx.AsyncClient(timeout=60.0) as client:
        response = await client.post(
            f"{SIAGPT_MEDIAS_URL}/",
            files={"file": (filename, data, "application/vnd.openxmlformats-officedocument.presentationml.presentation")},
            data={"media_metadata": media_metadata},
            headers={"Authorization": f"Bearer {auth_token}"},
        )
        response.raise_for_status()
        return response.json()


async def download_from_siagpt_medias(file_uuid: str, auth_token: str) -> tuple[bytes, str]:
    """
    Télécharge un fichier depuis la collection SiaGPT via GET /medias/{uuid}/download.
    Retourne (bytes, filename).
    """
    async with httpx.AsyncClient(timeout=60.0, follow_redirects=True) as client:
        # D'abord récupérer les métadonnées pour le nom du fichier
        meta_response = await client.get(
            f"{SIAGPT_MEDIAS_URL}/{file_uuid}",
            headers={"Authorization": f"Bearer {auth_token}"},
        )
        meta_response.raise_for_status()
        meta = meta_response.json()
        filename = meta.get("name", f"{file_uuid}.pptx")

        # Télécharger le fichier
        dl_response = await client.get(
            f"{SIAGPT_MEDIAS_URL}/{file_uuid}/download",
            headers={"Authorization": f"Bearer {auth_token}"},
        )
        dl_response.raise_for_status()
        return dl_response.content, filename
def load_system_prompt() -> str:
    try:
        return Path(SYSTEM_PROMPT_PATH).read_text(encoding="utf-8")
    except FileNotFoundError:
        return "Tu es un expert en manipulation PowerPoint. Retourne uniquement du code Python."

SYSTEM_PROMPT = load_system_prompt()




# ============================================================
# Inspection PPTX
# ============================================================

def inspect_pptx_structure(pptx_bytes: bytes) -> str:
    """Inspecte la structure complète d'un PPTX, retourne du JSON."""
    prs = Presentation(io.BytesIO(pptx_bytes))

    structure = {
        "slide_width_emu": str(prs.slide_width),
        "slide_height_emu": str(prs.slide_height),
        "slide_count": len(prs.slides),
        "slide_layouts": [],
        "slides": [],
    }

    # Layouts disponibles
    for i, layout in enumerate(prs.slide_layouts):
        structure["slide_layouts"].append({"index": i, "name": layout.name})

    # Contenu de chaque slide
    for i, slide in enumerate(prs.slides):
        slide_info = {
            "index": i,
            "layout": slide.slide_layout.name,
            "shapes": [],
        }
        for shape in slide.shapes:
            shape_info = {
                "name": shape.name,
                "shape_type": str(shape.shape_type),
                "left_emu": str(shape.left),
                "top_emu": str(shape.top),
                "width_emu": str(shape.width),
                "height_emu": str(shape.height),
            }
            if shape.has_text_frame:
                shape_info["text"] = shape.text_frame.text[:500]  # Tronquer si long
                shape_info["paragraphs"] = []
                for p in shape.text_frame.paragraphs:
                    para_info = {"text": p.text, "level": p.level}
                    if p.runs:
                        run = p.runs[0]
                        para_info["font_size"] = str(run.font.size) if run.font.size else None
                        para_info["bold"] = run.font.bold
                    shape_info["paragraphs"].append(para_info)
            if shape.has_table:
                table = shape.table
                shape_info["table"] = {
                    "rows": len(table.rows),
                    "cols": len(table.columns),
                    "cells_preview": [
                        [table.cell(r, c).text[:50] for c in range(min(len(table.columns), 5))]
                        for r in range(min(len(table.rows), 5))
                    ],
                }
            slide_info["shapes"].append(shape_info)
        structure["slides"].append(slide_info)

    return json.dumps(structure, ensure_ascii=False, indent=2)


def inspect_slide_xml(pptx_bytes: bytes, slide_index: int) -> str:
    """Retourne le XML brut d'un slide."""
    prs = Presentation(io.BytesIO(pptx_bytes))
    if slide_index >= len(prs.slides):
        return f"Erreur : slide {slide_index} n'existe pas (max: {len(prs.slides) - 1})"
    slide = prs.slides[slide_index]
    return etree.tostring(slide._element, pretty_print=True).decode()


# ============================================================
# Unpack / Repack PPTX (workflow d'édition XML)
# ============================================================

def unpack_pptx(pptx_bytes: bytes, dest_dir: str) -> str:
    """Décompresse un PPTX dans un dossier, retourne le chemin."""
    unpacked_dir = Path(dest_dir) / "unpacked"
    unpacked_dir.mkdir(parents=True, exist_ok=True)

    with zipfile.ZipFile(io.BytesIO(pptx_bytes), "r") as zf:
        zf.extractall(unpacked_dir)

    # Pretty-print les XML pour faciliter l'édition
    for xml_file in list(unpacked_dir.rglob("*.xml")) + list(unpacked_dir.rglob("*.rels")):
        try:
            content = xml_file.read_text(encoding="utf-8")
            dom = defusedxml.minidom.parseString(content)
            xml_file.write_bytes(dom.toprettyxml(indent="  ", encoding="utf-8"))
        except Exception:
            pass

    return str(unpacked_dir)


def repack_pptx(unpacked_dir: str, original_bytes: bytes = None) -> bytes:
    """Recompresse un dossier en PPTX, retourne les bytes."""
    unpacked_path = Path(unpacked_dir)

    # Condenser le XML (retirer les whitespaces ajoutés par pretty-print)
    for pattern in ["*.xml", "*.rels"]:
        for xml_file in unpacked_path.rglob(pattern):
            try:
                with open(xml_file, encoding="utf-8") as f:
                    dom = defusedxml.minidom.parse(f)

                # Retirer les text nodes vides (sauf dans les <a:t> tags)
                for element in dom.getElementsByTagName("*"):
                    if element.tagName.endswith(":t"):
                        continue
                    for child in list(element.childNodes):
                        if (
                            child.nodeType == child.TEXT_NODE
                            and child.nodeValue
                            and child.nodeValue.strip() == ""
                        ) or child.nodeType == child.COMMENT_NODE:
                            element.removeChild(child)

                xml_file.write_bytes(dom.toxml(encoding="UTF-8"))
            except Exception as e:
                print(f"Warning: Failed to condense {xml_file.name}: {e}")

    # Créer le ZIP
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for f in unpacked_path.rglob("*"):
            if f.is_file():
                zf.write(f, f.relative_to(unpacked_path))
    buf.seek(0)
    return buf.read()


# ============================================================
# Appel LLM
# ============================================================

async def call_llm(system_prompt: str, query: str) -> str:
    """
    Appelle SiaGPT /plain_llm endpoint.
    Format : { systemPrompt, query, llm, temperature } → string
    """
    async with httpx.AsyncClient(timeout=120.0) as client:
        response = await client.post(
            LLM_API_URL,
            json={
                "systemPrompt": system_prompt,
                "query": query,
                "llm": LLM_MODEL,
                "temperature": 0.1,
            },
            headers={
                "Authorization": f"Bearer {LLM_API_KEY}",
                "Content-Type": "application/json",
            },
        )
        response.raise_for_status()

        # /plain_llm retourne directement un string
        data = response.json()
        if isinstance(data, str):
            return data
        # Au cas où c'est wrappé dans un objet
        if isinstance(data, dict):
            return data.get("content", data.get("text", str(data)))
        return str(data)


def extract_code(llm_response: str) -> str:
    """Extrait le code Python de la réponse LLM (enlève les ```python si présents)."""
    code = llm_response.strip()

    # Enlever les blocs markdown si le LLM en a mis malgré les instructions
    if code.startswith("```python"):
        code = code[len("```python") :].strip()
    if code.startswith("```"):
        code = code[3:].strip()
    if code.endswith("```"):
        code = code[:-3].strip()

    return code


# ============================================================
# Exécution du code — MODE ÉDITION (XML direct)
# ============================================================

def execute_edit_code(code: str, unpacked_dir: str) -> dict:
    """Exécute du code d'édition XML dans le dossier décompressé."""
    exec_globals = {
        "__builtins__": __builtins__,
        "unpacked_dir": unpacked_dir,
        "os": os,
        "re": re,
        "shutil": shutil,
        "Path": Path,
        "defusedxml": defusedxml,
        "etree": etree,
        "json": json,
    }

    try:
        exec(code, exec_globals)
        return {"success": True}
    except Exception as e:
        return {
            "success": False,
            "error": str(e),
            "traceback": traceback.format_exc(),
        }


# ============================================================
# Exécution du code — MODE CRÉATION (python-pptx)
# ============================================================

def execute_create_code(code: str, prs: Presentation) -> dict:
    """Exécute du code de création python-pptx."""
    exec_globals = {
        "__builtins__": __builtins__,
        "prs": prs,
        "Presentation": Presentation,
        "Inches": Inches,
        "Pt": Pt,
        "Emu": Emu,
        "Cm": Cm,
        "RGBColor": RGBColor,
        "PP_ALIGN": PP_ALIGN,
        "MSO_ANCHOR": MSO_ANCHOR,
        "MSO_AUTO_SIZE": MSO_AUTO_SIZE,
        "MSO_SHAPE": MSO_SHAPE,
        "etree": etree,
        "os": os,
        "re": re,
        "Path": Path,
        "json": json,
    }

    try:
        exec(code, exec_globals)
        return {"success": True}
    except Exception as e:
        return {
            "success": False,
            "error": str(e),
            "traceback": traceback.format_exc(),
        }


# ============================================================
# Endpoint principal — Modification de PPTX
# ============================================================

@app.post("/api/edit")
async def edit_pptx(
    request: Request,
    prompt: str = Form(...),
    file: UploadFile = File(...),
    output_filename: str = Form(None),
):
    """
    Modifie un PPTX existant selon la demande utilisateur.
    Utilise le workflow unpack → edit XML → repack.
    """
    # Extraire le token d'auth pour le forwarding vers l'API medias
    auth_token = (request.headers.get("authorization", "").removeprefix("Bearer ").strip()) or LLM_API_KEY

    # 1. Lire le fichier uploadé
    pptx_bytes = await file.read()
    if not output_filename:
        output_filename = f"modified_{uuid.uuid4().hex[:8]}.pptx"

    # 2. Inspecter la structure
    structure = inspect_pptx_structure(pptx_bytes)

    # 3. Créer un dossier temporaire et décompresser
    with tempfile.TemporaryDirectory() as tmp_dir:
        unpacked_dir = unpack_pptx(pptx_bytes, tmp_dir)

        # 4. Construire la requête initiale pour le LLM
        base_query = (
            f"MODE : ÉDITION (modifier un fichier existant)\n\n"
            f"Structure du fichier PPTX :\n{structure}\n\n"
            f"Demande de l'utilisateur : {prompt}\n\n"
            f"Écris du code Python qui modifie les fichiers XML dans `unpacked_dir`.\n"
            f'unpacked_dir = "{unpacked_dir}"'
        )

        # 5. Boucle d'exécution avec retry
        query = base_query
        for attempt in range(MAX_RETRIES):
            llm_response = await call_llm(SYSTEM_PROMPT, query)
            code = extract_code(llm_response)

            result = execute_edit_code(code, unpacked_dir)

            if result["success"]:
                # Repack et sauvegarder dans la collection SiaGPT
                output_bytes = repack_pptx(unpacked_dir, pptx_bytes)
                media_info = await save_to_siagpt_medias(output_bytes, output_filename, auth_token)
                return {
                    "status": "ok",
                    "attempts": attempt + 1,
                    "media_uuid": media_info.get("uuid"),
                    "media_name": media_info.get("name"),
                }

            # Échec → concaténer l'erreur dans la query et retenter
            query = (
                f"{base_query}\n\n"
                f"--- TENTATIVE PRÉCÉDENTE (échouée) ---\n"
                f"Code généré :\n```python\n{code}\n```\n\n"
                f"Erreur (tentative {attempt + 1}/{MAX_RETRIES}) :\n"
                f"{result['traceback']}\n\n"
                f"Corrige le code. Retourne UNIQUEMENT le code Python corrigé."
            )

    raise HTTPException(status_code=500, detail=f"Échec après {MAX_RETRIES} tentatives")


# ============================================================
# Endpoint — Création de PPTX
# ============================================================

@app.post("/api/create")
async def create_pptx(
    request: Request,
    prompt: str = Form(...),
    template: UploadFile = File(None),
    output_filename: str = Form(None),
):
    """
    Crée un PPTX from scratch (ou depuis un template) selon la demande.
    Utilise python-pptx.
    """
    auth_token = (request.headers.get("authorization", "").removeprefix("Bearer ").strip()) or LLM_API_KEY
    if not output_filename:
        output_filename = f"new_{uuid.uuid4().hex[:8]}.pptx"

    # Charger le template si fourni
    if template:
        template_bytes = await template.read()
        prs = Presentation(io.BytesIO(template_bytes))
        structure = inspect_pptx_structure(template_bytes)
        mode_info = f"MODE : CRÉATION depuis un template\n\nStructure du template :\n{structure}"
    else:
        prs = Presentation()
        mode_info = "MODE : CRÉATION from scratch (présentation vide)"

    # Construire la requête
    base_query = (
        f"{mode_info}\n\n"
        f"Demande de l'utilisateur : {prompt}\n\n"
        f"Écris du code Python qui modifie l'objet `prs` (Presentation)."
    )

    # Boucle d'exécution avec retry
    query = base_query
    for attempt in range(MAX_RETRIES):
        llm_response = await call_llm(SYSTEM_PROMPT, query)
        code = extract_code(llm_response)

        result = execute_create_code(code, prs)

        if result["success"]:
            buf = io.BytesIO()
            prs.save(buf)
            output_bytes = buf.getvalue()
            media_info = await save_to_siagpt_medias(output_bytes, output_filename, auth_token)
            return {
                "status": "ok",
                "attempts": attempt + 1,
                "media_uuid": media_info.get("uuid"),
                "media_name": media_info.get("name"),
            }

        # Échec → concaténer l'erreur et retenter
        query = (
            f"{base_query}\n\n"
            f"--- TENTATIVE PRÉCÉDENTE (échouée) ---\n"
            f"Code généré :\n```python\n{code}\n```\n\n"
            f"Erreur (tentative {attempt + 1}/{MAX_RETRIES}) :\n"
            f"{result['traceback']}\n\n"
            f"Corrige le code. Retourne UNIQUEMENT le code Python corrigé."
        )

    raise HTTPException(status_code=500, detail=f"Échec après {MAX_RETRIES} tentatives")


# ============================================================
# Endpoint — Inspection
# ============================================================

@app.post("/api/inspect")
async def inspect_pptx(file: UploadFile = File(...)):
    """Retourne la structure d'un PPTX en JSON."""
    pptx_bytes = await file.read()
    structure = inspect_pptx_structure(pptx_bytes)
    return JSONResponse(content=json.loads(structure))


@app.post("/api/inspect/xml")
async def inspect_xml(file: UploadFile = File(...), slide_index: int = Form(0)):
    """Retourne le XML brut d'un slide."""
    pptx_bytes = await file.read()
    xml = inspect_slide_xml(pptx_bytes, slide_index)
    return {"slide_index": slide_index, "xml": xml}


# ============================================================
# MCP Server — SSE Transport
# ============================================================

# Sessions MCP actives : session_id → asyncio.Queue
mcp_sessions: dict[str, asyncio.Queue] = {}


def mcp_jsonrpc_response(req_id, result):
    """Construit une réponse JSON-RPC 2.0."""
    return {"jsonrpc": "2.0", "id": req_id, "result": result}


def mcp_jsonrpc_error(req_id, code, message):
    """Construit une erreur JSON-RPC 2.0."""
    return {"jsonrpc": "2.0", "id": req_id, "error": {"code": code, "message": message}}


@app.get("/mcp/sse")
async def mcp_sse_get(request: Request):
    """
    Endpoint SSE pour le protocole MCP (ancien transport).
    """
    session_id = uuid.uuid4().hex
    queue: asyncio.Queue = asyncio.Queue()
    mcp_sessions[session_id] = queue

    async def event_stream():
        scheme = request.headers.get("x-forwarded-proto", "https")
        host = request.headers.get("host", request.base_url.hostname)
        endpoint_url = f"{scheme}://{host}/mcp/messages?session_id={session_id}"
        yield f"event: endpoint\ndata: {endpoint_url}\n\n"

        try:
            while True:
                if await request.is_disconnected():
                    break
                try:
                    message = await asyncio.wait_for(queue.get(), timeout=30.0)
                    yield f"event: message\ndata: {json.dumps(message)}\n\n"
                except asyncio.TimeoutError:
                    yield ": keepalive\n\n"
        finally:
            mcp_sessions.pop(session_id, None)

    return StreamingResponse(
        event_stream(),
        media_type="text/event-stream",
        headers={
            "Cache-Control": "no-cache",
            "Connection": "keep-alive",
            "X-Accel-Buffering": "no",
        },
    )


async def handle_mcp_request(body: dict, session_id: str = "") -> tuple[dict, str]:
    """
    Traite une requête JSON-RPC MCP et retourne (réponse, session_id).
    """
    method = body.get("method", "")
    req_id = body.get("id")
    params = body.get("params", {})

    # --- initialize ---
    if method == "initialize":
        if not session_id:
            session_id = uuid.uuid4().hex
        return mcp_jsonrpc_response(req_id, {
            "protocolVersion": "2024-11-05",
            "capabilities": {"tools": {"listChanged": False}},
            "serverInfo": {"name": "pptx-service", "version": "1.0.0"},
        }), session_id

    # --- notifications/initialized ---
    if method == "notifications/initialized":
        return None, session_id

    # --- tools/list ---
    if method == "tools/list":
        return mcp_jsonrpc_response(req_id, {
            "tools": [
                {
                    "name": "generate_pptx",
                    "description": "Génère une présentation PowerPoint à partir d'une description textuelle. Le fichier est sauvegardé dans la collection SiaGPT.",
                    "inputSchema": {
                        "type": "object",
                        "properties": {
                            "prompt": {
                                "type": "string",
                                "description": "Description de la présentation à créer (contenu, nombre de slides, style...)",
                            }
                        },
                        "required": ["prompt"],
                    },
                },
                {
                    "name": "edit_pptx",
                    "description": "Modifie une présentation PowerPoint existante dans la collection SiaGPT. Récupère le fichier par son UUID, applique les modifications demandées, et uploade la version modifiée.",
                    "inputSchema": {
                        "type": "object",
                        "properties": {
                            "prompt": {
                                "type": "string",
                                "description": "Description des modifications à apporter (ex: changer les couleurs, ajouter une slide, modifier le texte...)",
                            },
                            "source_file_id": {
                                "type": "string",
                                "description": "UUID du fichier PPTX dans la collection SiaGPT à modifier",
                            }
                        },
                        "required": ["prompt", "source_file_id"],
                    },
                }
            ]
        }), session_id

    # --- tools/call ---
    if method == "tools/call":
        tool_name = params.get("name", "")
        tool_args = params.get("arguments", {})

        if tool_name == "generate_pptx":
            prompt = tool_args.get("prompt", "")
            if not prompt:
                return mcp_jsonrpc_error(req_id, -32602, "Le paramètre 'prompt' est requis"), session_id

            try:
                auth_token = LLM_API_KEY
                output_filename = f"new_{uuid.uuid4().hex[:8]}.pptx"

                prs = Presentation()
                mode_info = "MODE : CRÉATION from scratch (présentation vide)"
                base_query = (
                    f"{mode_info}\n\n"
                    f"Demande de l'utilisateur : {prompt}\n\n"
                    f"Écris du code Python qui modifie l'objet `prs` (Presentation)."
                )

                query = base_query
                for attempt in range(MAX_RETRIES):
                    llm_response = await call_llm(SYSTEM_PROMPT, query)
                    code = extract_code(llm_response)
                    result = execute_create_code(code, prs)

                    if result["success"]:
                        buf = io.BytesIO()
                        prs.save(buf)
                        output_bytes = buf.getvalue()
                        media_info = await save_to_siagpt_medias(output_bytes, output_filename, auth_token)
                        return mcp_jsonrpc_response(req_id, {
                            "content": [
                                {
                                    "type": "text",
                                    "text": f"Présentation créée avec succès !\n- Fichier : {media_info.get('name', output_filename)}\n- UUID : {media_info.get('uuid', 'N/A')}\n- Tentatives : {attempt + 1}",
                                }
                            ]
                        }), session_id

                    query = (
                        f"{base_query}\n\n"
                        f"--- TENTATIVE PRÉCÉDENTE (échouée) ---\n"
                        f"Code généré :\n```python\n{code}\n```\n\n"
                        f"Erreur (tentative {attempt + 1}/{MAX_RETRIES}) :\n"
                        f"{result['traceback']}\n\n"
                        f"Corrige le code. Retourne UNIQUEMENT le code Python corrigé."
                    )

                return mcp_jsonrpc_error(req_id, -32000, f"Échec après {MAX_RETRIES} tentatives"), session_id

            except Exception as e:
                return mcp_jsonrpc_error(req_id, -32000, str(e)), session_id

        if tool_name == "edit_pptx":
            prompt = tool_args.get("prompt", "")
            source_file_id = tool_args.get("source_file_id", "")
            if not prompt:
                return mcp_jsonrpc_error(req_id, -32602, "Le paramètre 'prompt' est requis"), session_id
            if not source_file_id:
                return mcp_jsonrpc_error(req_id, -32602, "Le paramètre 'source_file_id' est requis"), session_id

            try:
                auth_token = LLM_API_KEY

                # 1. Télécharger le fichier depuis la collection SiaGPT
                pptx_bytes, original_filename = await download_from_siagpt_medias(source_file_id, auth_token)
                output_filename = f"modified_{uuid.uuid4().hex[:8]}.pptx"

                # 2. Inspecter la structure
                structure = inspect_pptx_structure(pptx_bytes)

                # 3. Décompresser et éditer
                with tempfile.TemporaryDirectory() as tmp_dir:
                    unpacked_dir = unpack_pptx(pptx_bytes, tmp_dir)

                    base_query = (
                        f"MODE : ÉDITION (modifier un fichier existant)\n\n"
                        f"Fichier source : {original_filename}\n"
                        f"Structure du fichier PPTX :\n{structure}\n\n"
                        f"Demande de l'utilisateur : {prompt}\n\n"
                        f"Écris du code Python qui modifie les fichiers XML dans `unpacked_dir`.\n"
                        f'unpacked_dir = "{unpacked_dir}"'
                    )

                    query = base_query
                    for attempt in range(MAX_RETRIES):
                        llm_response = await call_llm(SYSTEM_PROMPT, query)
                        code = extract_code(llm_response)
                        result = execute_edit_code(code, unpacked_dir)

                        if result["success"]:
                            output_bytes = repack_pptx(unpacked_dir, pptx_bytes)
                            media_info = await save_to_siagpt_medias(output_bytes, output_filename, auth_token)
                            return mcp_jsonrpc_response(req_id, {
                                "content": [
                                    {
                                        "type": "text",
                                        "text": f"Présentation modifiée avec succès !\n- Source : {original_filename} ({source_file_id})\n- Nouveau fichier : {media_info.get('name', output_filename)}\n- UUID : {media_info.get('uuid', 'N/A')}\n- Tentatives : {attempt + 1}",
                                    }
                                ]
                            }), session_id

                        query = (
                            f"{base_query}\n\n"
                            f"--- TENTATIVE PRÉCÉDENTE (échouée) ---\n"
                            f"Code généré :\n```python\n{code}\n```\n\n"
                            f"Erreur (tentative {attempt + 1}/{MAX_RETRIES}) :\n"
                            f"{result['traceback']}\n\n"
                            f"Corrige le code. Retourne UNIQUEMENT le code Python corrigé."
                        )

                return mcp_jsonrpc_error(req_id, -32000, f"Échec après {MAX_RETRIES} tentatives"), session_id

            except httpx.HTTPStatusError as e:
                return mcp_jsonrpc_error(req_id, -32000, f"Impossible de récupérer le fichier {source_file_id} : {e.response.status_code}"), session_id
            except Exception as e:
                return mcp_jsonrpc_error(req_id, -32000, str(e)), session_id

        return mcp_jsonrpc_error(req_id, -32601, f"Tool inconnu : {tool_name}"), session_id

    return mcp_jsonrpc_error(req_id, -32601, f"Méthode inconnue : {method}"), session_id


@app.post("/mcp/sse")
async def mcp_sse_post(request: Request):
    """
    Endpoint Streamable HTTP — POST direct avec réponse JSON-RPC.
    """
    body = await request.json()
    session_id = request.headers.get("mcp-session-id", "")
    response, session_id = await handle_mcp_request(body, session_id)
    
    if response is None:
        # Notification — pas de réponse body
        return JSONResponse(
            content={},
            status_code=202,
            headers={"mcp-session-id": session_id},
        )
    
    return JSONResponse(
        content=response,
        headers={"mcp-session-id": session_id},
    )


@app.delete("/mcp/sse")
async def mcp_sse_delete(request: Request):
    """Fermeture de session MCP."""
    return JSONResponse({"status": "ok"})


@app.get("/mcp/messages")
async def mcp_messages_get(request: Request, session_id: str = ""):
    """
    GET sur /mcp/messages — Langflow vérifie l'endpoint ou ouvre un stream SSE.
    """
    return JSONResponse({"status": "ok", "message": "Use POST to send MCP messages"})


@app.post("/mcp/messages")
async def mcp_messages(request: Request, session_id: str = ""):
    """
    Endpoint messages pour le transport SSE classique.
    """
    if session_id not in mcp_sessions:
        body = await request.json()
        response, _ = await handle_mcp_request(body, session_id)
        if response is None:
            return JSONResponse({}, status_code=202)
        return JSONResponse(response)

    queue = mcp_sessions[session_id]
    body = await request.json()
    response, _ = await handle_mcp_request(body, session_id)
    if response is not None:
        await queue.put(response)
    return JSONResponse({"status": "ok"})


# ============================================================
# Endpoint unifié — Génère ou modifie un PPTX automatiquement
# ============================================================

@app.post("/api/generate")
async def generate_pptx(
    request: Request,
):
    """
    Endpoint unifié — accepte JSON ou form-data :
    - JSON : {"prompt": "..."}
    - Form-data : prompt + fichier optionnel
    """
    content_type = request.headers.get("content-type", "")

    if "application/json" in content_type:
        # Mode JSON (Langflow, API calls)
        body = await request.json()
        prompt = body.get("prompt", "")
        output_filename = body.get("output_filename", None)
        if not prompt:
            raise HTTPException(status_code=400, detail="Le champ 'prompt' est requis")
        return await create_pptx(request, prompt, None, output_filename)
    else:
        # Mode form-data (curl, upload de fichier)
        form = await request.form()
        prompt = form.get("prompt", "")
        file = form.get("file", None)
        output_filename = form.get("output_filename", None)
        if not prompt:
            raise HTTPException(status_code=400, detail="Le champ 'prompt' est requis")
        if file and hasattr(file, 'filename') and file.filename:
            return await edit_pptx(request, prompt, file, output_filename)
        else:
            return await create_pptx(request, prompt, None, output_filename)


# ============================================================
# Health check
# ============================================================

@app.get("/health")
async def health():
    return {
        "status": "ok",
        "llm_configured": bool(LLM_API_KEY),
        "collection_configured": bool(SIAGPT_COLLECTION_ID),
        "model": LLM_MODEL,
    }


@app.api_route("/", methods=["GET", "POST", "DELETE"])
async def root(request: Request):
    """Racine — healthcheck et fallback MCP."""
    if request.method == "POST":
        try:
            body = await request.json()
            if "jsonrpc" in body:
                # C'est une requête MCP
                session_id = request.headers.get("mcp-session-id", "")
                response, session_id = await handle_mcp_request(body, session_id)
                if response is None:
                    return JSONResponse({}, status_code=202, headers={"mcp-session-id": session_id})
                return JSONResponse(content=response, headers={"mcp-session-id": session_id})
        except Exception:
            pass
    return JSONResponse({"status": "ok", "service": "pptx-service"})
