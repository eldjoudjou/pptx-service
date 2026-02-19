"""
pptx_tools.py — Manipulation de fichiers PPTX.

Fonctions extraites et adaptées du skill PPTX d'Anthropic :
- unpack : décompresse + pretty-print XML + escape smart quotes
- pack : condense XML + repackage
- clean : supprime slides orphelines, fichiers non-référencés
- duplicate_slide : duplique une slide avec ses relations

La validation est dans pptx_validate.py (module séparé).
"""

import io
import re
import shutil
import tempfile
import zipfile
from pathlib import Path

import defusedxml.minidom


# ============================================================
# Smart Quotes — escape/unescape
# ============================================================

SMART_QUOTE_REPLACEMENTS = {
    "\u201c": "&#x201C;",  # "
    "\u201d": "&#x201D;",  # "
    "\u2018": "&#x2018;",  # '
    "\u2019": "&#x2019;",  # '
}

SMART_QUOTE_RESTORE = {v: k for k, v in SMART_QUOTE_REPLACEMENTS.items()}


# ============================================================
# UNPACK — Décompresse un PPTX avec pretty-print XML
# ============================================================

def unpack(pptx_bytes: bytes, output_dir: str) -> str:
    """
    Décompresse un PPTX en mémoire vers output_dir.
    - Pretty-print les fichiers XML pour lisibilité
    - Escape les smart quotes pour éviter les problèmes d'encodage

    Retourne le chemin du dossier décompressé.
    """
    output_path = Path(output_dir)
    output_path.mkdir(parents=True, exist_ok=True)

    with zipfile.ZipFile(io.BytesIO(pptx_bytes), "r") as zf:
        zf.extractall(output_path)

    # Pretty-print tous les XML
    xml_files = list(output_path.rglob("*.xml")) + list(output_path.rglob("*.rels"))
    for xml_file in xml_files:
        _pretty_print_xml(xml_file)

    # Escape smart quotes
    for xml_file in xml_files:
        _escape_smart_quotes(xml_file)

    return str(output_path)


def _pretty_print_xml(xml_file: Path) -> None:
    """Pretty-print un fichier XML avec indentation."""
    try:
        content = xml_file.read_text(encoding="utf-8")
        dom = defusedxml.minidom.parseString(content)
        xml_file.write_bytes(dom.toprettyxml(indent="  ", encoding="utf-8"))
    except Exception:
        pass  # Fichiers non-XML (binaires) ignorés silencieusement


def _escape_smart_quotes(xml_file: Path) -> None:
    """Remplace les smart quotes par des entités XML."""
    try:
        content = xml_file.read_text(encoding="utf-8")
        for char, entity in SMART_QUOTE_REPLACEMENTS.items():
            content = content.replace(char, entity)
        xml_file.write_text(content, encoding="utf-8")
    except Exception:
        pass


# ============================================================
# PACK — Repackage un dossier en PPTX
# ============================================================

def pack(unpacked_dir: str, original_bytes: bytes = None) -> bytes:
    """
    Repackage un dossier décompressé en PPTX.
    - Restore les smart quotes en vrais caractères unicode
    - Condense le XML (supprime whitespace inutile sauf dans les <a:t>)
    - Retourne les bytes du fichier PPTX.
    """
    input_dir = Path(unpacked_dir)

    with tempfile.TemporaryDirectory() as temp_dir:
        temp_content_dir = Path(temp_dir) / "content"
        shutil.copytree(input_dir, temp_content_dir)

        # Restaurer les smart quotes puis condenser le XML
        for pattern in ["*.xml", "*.rels"]:
            for xml_file in temp_content_dir.rglob(pattern):
                _restore_smart_quotes(xml_file)
                _condense_xml(xml_file)

        # Créer le ZIP
        buf = io.BytesIO()
        with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
            for f in sorted(temp_content_dir.rglob("*")):
                if f.is_file():
                    zf.write(f, f.relative_to(temp_content_dir))

        return buf.getvalue()


def _restore_smart_quotes(xml_file: Path) -> None:
    """Restaure les entités smart quotes en vrais caractères unicode."""
    try:
        content = xml_file.read_text(encoding="utf-8")
        for entity, char in SMART_QUOTE_RESTORE.items():
            content = content.replace(entity, char)
        xml_file.write_text(content, encoding="utf-8")
    except Exception:
        pass


def _condense_xml(xml_file: Path) -> None:
    """
    Condense un fichier XML en supprimant le whitespace inutile.
    Préserve le contenu des tags <a:t> (texte visible dans les slides).
    """
    try:
        with open(xml_file, encoding="utf-8") as f:
            dom = defusedxml.minidom.parse(f)

        for element in dom.getElementsByTagName("*"):
            # Ne pas toucher aux éléments de texte
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
    except Exception:
        pass  # Fichiers qui ne parsent pas → laisser tels quels


# ============================================================
# CLEAN — Supprime fichiers orphelins
# ============================================================

def clean(unpacked_dir: str) -> list[str]:
    """
    Supprime les fichiers orphelins d'un PPTX décompressé.
    - Slides non référencées dans presentation.xml
    - Fichiers media/embeddings/charts non référencés
    - Met à jour [Content_Types].xml

    Retourne la liste des fichiers supprimés.
    """
    path = Path(unpacked_dir)
    all_removed = []

    # 1. Slides orphelines
    slides_removed = _remove_orphaned_slides(path)
    all_removed.extend(slides_removed)

    # 2. Dossier [trash]
    trash_removed = _remove_trash_directory(path)
    all_removed.extend(trash_removed)

    # 3. Fichiers non référencés (boucle jusqu'à stabilisation)
    while True:
        removed_rels = _remove_orphaned_rels_files(path)
        referenced = _get_referenced_files(path)
        removed_files = _remove_orphaned_files(path, referenced)

        total_removed = removed_rels + removed_files
        if not total_removed:
            break
        all_removed.extend(total_removed)

    # 4. Mettre à jour Content_Types
    if all_removed:
        _update_content_types(path, all_removed)

    return all_removed


def _get_slides_in_sldidlst(unpacked_dir: Path) -> set[str]:
    """Retourne les noms de fichiers slides référencés dans presentation.xml."""
    pres_path = unpacked_dir / "ppt" / "presentation.xml"
    pres_rels_path = unpacked_dir / "ppt" / "_rels" / "presentation.xml.rels"

    if not pres_path.exists() or not pres_rels_path.exists():
        return set()

    rels_dom = defusedxml.minidom.parse(str(pres_rels_path))
    rid_to_slide = {}
    for rel in rels_dom.getElementsByTagName("Relationship"):
        rid = rel.getAttribute("Id")
        target = rel.getAttribute("Target")
        rel_type = rel.getAttribute("Type")
        if "slide" in rel_type and target.startswith("slides/"):
            rid_to_slide[rid] = target.replace("slides/", "")

    pres_content = pres_path.read_text(encoding="utf-8")
    referenced_rids = set(re.findall(r'<p:sldId[^>]*r:id="([^"]+)"', pres_content))

    return {rid_to_slide[rid] for rid in referenced_rids if rid in rid_to_slide}


def _remove_orphaned_slides(unpacked_dir: Path) -> list[str]:
    """Supprime les slides non référencées dans <p:sldIdLst>."""
    slides_dir = unpacked_dir / "ppt" / "slides"
    slides_rels_dir = slides_dir / "_rels"
    pres_rels_path = unpacked_dir / "ppt" / "_rels" / "presentation.xml.rels"

    if not slides_dir.exists():
        return []

    referenced_slides = _get_slides_in_sldidlst(unpacked_dir)
    removed = []

    for slide_file in slides_dir.glob("slide*.xml"):
        if slide_file.name not in referenced_slides:
            rel_path = slide_file.relative_to(unpacked_dir)
            slide_file.unlink()
            removed.append(str(rel_path))

            rels_file = slides_rels_dir / f"{slide_file.name}.rels"
            if rels_file.exists():
                rels_file.unlink()
                removed.append(str(rels_file.relative_to(unpacked_dir)))

    # Nettoyer presentation.xml.rels
    if removed and pres_rels_path.exists():
        rels_dom = defusedxml.minidom.parse(str(pres_rels_path))
        changed = False

        for rel in list(rels_dom.getElementsByTagName("Relationship")):
            target = rel.getAttribute("Target")
            if target.startswith("slides/"):
                slide_name = target.replace("slides/", "")
                if slide_name not in referenced_slides:
                    if rel.parentNode:
                        rel.parentNode.removeChild(rel)
                        changed = True

        if changed:
            with open(pres_rels_path, "wb") as f:
                f.write(rels_dom.toxml(encoding="utf-8"))

    return removed


def _remove_trash_directory(unpacked_dir: Path) -> list[str]:
    """Supprime le dossier [trash] s'il existe."""
    trash_dir = unpacked_dir / "[trash]"
    removed = []

    if trash_dir.exists() and trash_dir.is_dir():
        for file_path in trash_dir.iterdir():
            if file_path.is_file():
                removed.append(str(file_path.relative_to(unpacked_dir)))
                file_path.unlink()
        trash_dir.rmdir()

    return removed


def _get_slide_referenced_files(unpacked_dir: Path) -> set:
    """Retourne l'ensemble des fichiers référencés par les slides."""
    referenced = set()
    slides_rels_dir = unpacked_dir / "ppt" / "slides" / "_rels"

    if not slides_rels_dir.exists():
        return referenced

    for rels_file in slides_rels_dir.glob("*.rels"):
        dom = defusedxml.minidom.parse(str(rels_file))
        for rel in dom.getElementsByTagName("Relationship"):
            target = rel.getAttribute("Target")
            if not target:
                continue
            target_path = (rels_file.parent.parent / target).resolve()
            try:
                referenced.add(target_path.relative_to(unpacked_dir.resolve()))
            except ValueError:
                pass

    return referenced


def _remove_orphaned_rels_files(unpacked_dir: Path) -> list[str]:
    """Supprime les fichiers .rels orphelins."""
    resource_dirs = ["charts", "diagrams", "drawings"]
    removed = []
    slide_referenced = _get_slide_referenced_files(unpacked_dir)

    for dir_name in resource_dirs:
        rels_dir = unpacked_dir / "ppt" / dir_name / "_rels"
        if not rels_dir.exists():
            continue

        for rels_file in rels_dir.glob("*.rels"):
            resource_file = rels_dir.parent / rels_file.name.replace(".rels", "")
            try:
                resource_rel_path = resource_file.resolve().relative_to(unpacked_dir.resolve())
            except ValueError:
                continue

            if not resource_file.exists() or resource_rel_path not in slide_referenced:
                rels_file.unlink()
                removed.append(str(rels_file.relative_to(unpacked_dir)))

    return removed


def _get_referenced_files(unpacked_dir: Path) -> set:
    """Retourne l'ensemble de tous les fichiers référencés dans les .rels."""
    referenced = set()

    for rels_file in unpacked_dir.rglob("*.rels"):
        dom = defusedxml.minidom.parse(str(rels_file))
        for rel in dom.getElementsByTagName("Relationship"):
            target = rel.getAttribute("Target")
            if not target:
                continue
            target_path = (rels_file.parent.parent / target).resolve()
            try:
                referenced.add(target_path.relative_to(unpacked_dir.resolve()))
            except ValueError:
                pass

    return referenced


def _remove_orphaned_files(unpacked_dir: Path, referenced: set) -> list[str]:
    """Supprime les fichiers media/embeddings/etc non référencés."""
    resource_dirs = ["media", "embeddings", "charts", "diagrams", "tags", "drawings", "ink"]
    removed = []

    for dir_name in resource_dirs:
        dir_path = unpacked_dir / "ppt" / dir_name
        if not dir_path.exists():
            continue

        for file_path in dir_path.glob("*"):
            if not file_path.is_file():
                continue
            rel_path = file_path.relative_to(unpacked_dir)
            if rel_path not in referenced:
                file_path.unlink()
                removed.append(str(rel_path))

    # Themes orphelins
    theme_dir = unpacked_dir / "ppt" / "theme"
    if theme_dir.exists():
        for file_path in theme_dir.glob("theme*.xml"):
            rel_path = file_path.relative_to(unpacked_dir)
            if rel_path not in referenced:
                file_path.unlink()
                removed.append(str(rel_path))
                theme_rels = theme_dir / "_rels" / f"{file_path.name}.rels"
                if theme_rels.exists():
                    theme_rels.unlink()
                    removed.append(str(theme_rels.relative_to(unpacked_dir)))

    # Notes slides orphelines
    notes_dir = unpacked_dir / "ppt" / "notesSlides"
    if notes_dir.exists():
        for file_path in notes_dir.glob("*.xml"):
            if not file_path.is_file():
                continue
            rel_path = file_path.relative_to(unpacked_dir)
            if rel_path not in referenced:
                file_path.unlink()
                removed.append(str(rel_path))

        notes_rels_dir = notes_dir / "_rels"
        if notes_rels_dir.exists():
            for file_path in notes_rels_dir.glob("*.rels"):
                notes_file = notes_dir / file_path.name.replace(".rels", "")
                if not notes_file.exists():
                    file_path.unlink()
                    removed.append(str(file_path.relative_to(unpacked_dir)))

    return removed


def _update_content_types(unpacked_dir: Path, removed_files: list[str]) -> None:
    """Met à jour [Content_Types].xml après suppression de fichiers."""
    ct_path = unpacked_dir / "[Content_Types].xml"
    if not ct_path.exists():
        return

    dom = defusedxml.minidom.parse(str(ct_path))
    changed = False

    for override in list(dom.getElementsByTagName("Override")):
        part_name = override.getAttribute("PartName").lstrip("/")
        if part_name in removed_files:
            if override.parentNode:
                override.parentNode.removeChild(override)
                changed = True

    if changed:
        with open(ct_path, "wb") as f:
            f.write(dom.toxml(encoding="utf-8"))


# ============================================================
# DUPLICATE SLIDE — Duplique une slide avec ses relations
# ============================================================

def duplicate_slide(unpacked_dir: str, source_filename: str) -> dict:
    """
    Duplique une slide existante.
    - Copie le XML de la slide
    - Copie les .rels (sans les notesSlide pour éviter doublons)
    - Met à jour Content_Types, presentation.xml.rels

    Retourne un dict avec :
    - new_filename: nom du nouveau fichier (ex: "slide6.xml")
    - new_sld_id: id pour <p:sldId>
    - new_r_id: rId pour la relation
    """
    path = Path(unpacked_dir)
    slides_dir = path / "ppt" / "slides"
    rels_dir = slides_dir / "_rels"

    source_slide = slides_dir / source_filename
    if not source_slide.exists():
        raise FileNotFoundError(f"Slide source {source_filename} introuvable")

    # Trouver le prochain numéro
    existing = [int(m.group(1)) for f in slides_dir.glob("slide*.xml")
                if (m := re.match(r"slide(\d+)\.xml", f.name))]
    next_num = max(existing) + 1 if existing else 1
    dest = f"slide{next_num}.xml"
    dest_slide = slides_dir / dest

    # Copier la slide
    shutil.copy2(source_slide, dest_slide)

    # Copier les rels (sans notesSlide)
    source_rels = rels_dir / f"{source_filename}.rels"
    dest_rels = rels_dir / f"{dest}.rels"

    if source_rels.exists():
        shutil.copy2(source_rels, dest_rels)
        # Retirer les références aux notesSlide pour éviter les doublons
        rels_content = dest_rels.read_text(encoding="utf-8")
        rels_content = re.sub(
            r'\s*<Relationship[^>]*Type="[^"]*notesSlide"[^>]*/>\s*',
            "\n",
            rels_content,
        )
        dest_rels.write_text(rels_content, encoding="utf-8")

    # Mettre à jour [Content_Types].xml
    ct_path = path / "[Content_Types].xml"
    ct_content = ct_path.read_text(encoding="utf-8")
    if f"/ppt/slides/{dest}" not in ct_content:
        new_override = f'<Override PartName="/ppt/slides/{dest}" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>'
        ct_content = ct_content.replace("</Types>", f"  {new_override}\n</Types>")
        ct_path.write_text(ct_content, encoding="utf-8")

    # Ajouter dans presentation.xml.rels
    pres_rels_path = path / "ppt" / "_rels" / "presentation.xml.rels"
    pres_rels = pres_rels_path.read_text(encoding="utf-8")
    rids = [int(m) for m in re.findall(r'Id="rId(\d+)"', pres_rels)]
    next_rid = max(rids) + 1 if rids else 1
    rid = f"rId{next_rid}"

    if f"slides/{dest}" not in pres_rels:
        new_rel = f'<Relationship Id="{rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/{dest}"/>'
        pres_rels = pres_rels.replace("</Relationships>", f"  {new_rel}\n</Relationships>")
        pres_rels_path.write_text(pres_rels, encoding="utf-8")

    # Trouver le prochain slide ID
    pres_path = path / "ppt" / "presentation.xml"
    pres_content = pres_path.read_text(encoding="utf-8")
    slide_ids = [int(m) for m in re.findall(r'<p:sldId[^>]*id="(\d+)"', pres_content)]
    new_sld_id = max(slide_ids) + 1 if slide_ids else 256

    return {
        "new_filename": dest,
        "new_sld_id": new_sld_id,
        "new_r_id": rid,
    }


def add_slide_to_presentation(unpacked_dir: str, sld_id: int, r_id: str, position: int = None) -> None:
    """
    Ajoute un <p:sldId> dans <p:sldIdLst> de presentation.xml.
    Si position est spécifié (1-based), insère à cette position.
    Sinon, ajoute à la fin.
    """
    path = Path(unpacked_dir)
    pres_path = path / "ppt" / "presentation.xml"
    pres_content = pres_path.read_text(encoding="utf-8")

    new_entry = f'<p:sldId id="{sld_id}" r:id="{r_id}"/>'

    if position is not None:
        # Extraire les sldId existants, insérer à la bonne position
        sld_id_pattern = r'(<p:sldId[^/]*/\s*>)'
        existing_entries = re.findall(sld_id_pattern, pres_content)

        # Insérer à la position demandée (1-based, clampé)
        idx = max(0, min(position - 1, len(existing_entries)))
        existing_entries.insert(idx, new_entry)

        # Reconstruire le bloc <p:sldIdLst>
        new_block = "<p:sldIdLst>\n    " + "\n    ".join(existing_entries) + "\n  </p:sldIdLst>"
        pres_content = re.sub(
            r'<p:sldIdLst>.*?</p:sldIdLst>',
            new_block,
            pres_content,
            flags=re.DOTALL,
        )
    else:
        # Ajouter à la fin
        pres_content = pres_content.replace(
            "</p:sldIdLst>",
            f"  {new_entry}\n  </p:sldIdLst>",
        )

    pres_path.write_text(pres_content, encoding="utf-8")

