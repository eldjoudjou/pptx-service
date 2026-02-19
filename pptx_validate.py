"""
pptx_validate.py — Validation complète de fichiers PPTX décompressés.

Deux niveaux de validation :
1. Structurelle : XML bien formé, références .rels, IDs uniques, Content_Types, etc.
2. XSD : chaque fichier XML vérifié contre les schemas officiels Office Open XML.

Logique extraite et adaptée du skill PPTX d'Anthropic
(skill/scripts/office/validators/base.py + pptx.py).

Usage depuis main.py :
    from pptx_validate import validate_pptx
    result = validate_pptx("/tmp/unpacked", original_bytes=pptx_bytes)
    # result = {"valid": True, "repairs": 0, "errors": [], "xsd_errors": []}
"""

import re
import tempfile
import zipfile
from pathlib import Path

import logging

import defusedxml.minidom
import lxml.etree

logger = logging.getLogger(__name__)


# ============================================================
# Constantes
# ============================================================

# Namespaces Office XML
PKG_RELS_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
OFFICE_RELS_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
CONTENT_TYPES_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
PML_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
MC_NS = "http://schemas.openxmlformats.org/markup-compatibility/2006"

# Namespaces standards OOXML — tout ce qui n'est PAS dans cette liste
# est considéré comme une extension Microsoft propriétaire et ignoré
# lors de la validation XSD (car les schemas ISO ne les connaissent pas).
OOXML_NAMESPACES = {
    "http://schemas.openxmlformats.org/officeDocument/2006/math",
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "http://schemas.openxmlformats.org/schemaLibrary/2006/main",
    "http://schemas.openxmlformats.org/drawingml/2006/main",
    "http://schemas.openxmlformats.org/drawingml/2006/chart",
    "http://schemas.openxmlformats.org/drawingml/2006/chartDrawing",
    "http://schemas.openxmlformats.org/drawingml/2006/diagram",
    "http://schemas.openxmlformats.org/drawingml/2006/picture",
    "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
    "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "http://schemas.openxmlformats.org/presentationml/2006/main",
    "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "http://schemas.openxmlformats.org/officeDocument/2006/sharedTypes",
    "http://www.w3.org/XML/1998/namespace",
}

# Correspondance fichier XML → schema XSD.
# "ppt" = tout fichier directement sous ppt/ (slides, layouts, masters)
SCHEMA_MAPPINGS = {
    "ppt": "ISO-IEC29500-4_2016/pml.xsd",
    "[Content_Types].xml": "ecma/fouth-edition/opc-contentTypes.xsd",
    "app.xml": "ISO-IEC29500-4_2016/shared-documentPropertiesExtended.xsd",
    "core.xml": "ecma/fouth-edition/opc-coreProperties.xsd",
    "custom.xml": "ISO-IEC29500-4_2016/shared-documentPropertiesCustom.xsd",
    ".rels": "ecma/fouth-edition/opc-relationships.xsd",
    "chart": "ISO-IEC29500-4_2016/dml-chart.xsd",
    "theme": "ISO-IEC29500-4_2016/dml-main.xsd",
    "drawing": "ISO-IEC29500-4_2016/dml-main.xsd",
}

# Erreurs XSD connues et bénignes qu'on ignore.
# - hyphenationZone : extension Word, pas dans les schemas ISO
# - purl.org/dc/terms : namespace Dublin Core, utilisé dans core.xml
# - xml:space : attribut XML standard, toujours valide mais pas déclaré
#   dans les schemas Office (ajouté par notre auto-repair whitespace)
IGNORED_XSD_ERRORS = [
    "hyphenationZone",
    "purl.org/dc/terms",
    "{http://www.w3.org/XML/1998/namespace}space",
]

# Chemin vers les schemas XSD — relatif à ce fichier (dev) ou /app (Docker)
def _find_schemas_dir() -> Path:
    """Trouve le dossier schemas/ contenant les .xsd Office."""
    candidates = [
        Path(__file__).parent / "schemas",   # En dev local
        Path("/app/schemas"),                 # En Docker
    ]
    for c in candidates:
        if c.is_dir() and any(c.rglob("*.xsd")):
            return c
    raise FileNotFoundError(
        "Dossier schemas/ introuvable. "
        "Vérifier que schemas/ existe à côté de pptx_validate.py ou dans /app/."
    )


# ============================================================
# Validation rapide d'un slide XML (pour le retry loop)
# ============================================================

def validate_slide_xml_string(xml_string: str) -> tuple[bool, str]:
    """
    Valide un XML de slide contre le schema PresentationML (pml.xsd).

    Utilisé dans le retry loop de modify_slide_xml() pour détecter
    les erreurs XSD AVANT d'écrire le fichier sur disque.

    Vérifie :
    1. XML bien formé (parsing)
    2. Conformité XSD contre pml.xsd (grammaire PowerPoint)

    Returns:
        (True, "")           si valide
        (False, "message")   si invalide (parsing ou XSD)
    """
    # 1. Vérifier que le XML se parse
    try:
        xml_doc = lxml.etree.ElementTree(
            lxml.etree.fromstring(xml_string.encode("utf-8"))
        )
    except lxml.etree.XMLSyntaxError as e:
        return False, f"XML mal formé : {e}"

    # 2. Charger le schema pml.xsd
    try:
        schemas_dir = _find_schemas_dir()
        schema_path = schemas_dir / SCHEMA_MAPPINGS["ppt"]
        with open(schema_path, "rb") as xsd_fh:
            parser = lxml.etree.XMLParser()
            xsd_doc = lxml.etree.parse(xsd_fh, parser=parser, base_url=str(schema_path))
            schema = lxml.etree.XMLSchema(xsd_doc)
    except Exception:
        # Schema indisponible → fallback sur validation parsing seule
        return True, ""

    # 3. Pré-traiter (même logique que _validate_one_file_xsd)
    xml_doc = _strip_template_tags(xml_doc)
    xml_doc = _strip_mc_ignorable(xml_doc)
    xml_doc = _strip_non_ooxml(xml_doc)

    # 4. Valider contre le schema
    if schema.validate(xml_doc):
        return True, ""

    # Filtrer les erreurs bénignes connues
    real_errors = []
    for error in schema.error_log:
        if not any(ignored in error.message for ignored in IGNORED_XSD_ERRORS):
            real_errors.append(error.message)

    if not real_errors:
        return True, ""

    return False, " | ".join(real_errors[:3])


# ============================================================
# Point d'entrée principal
# ============================================================

def validate_pptx(unpacked_dir: str, original_bytes: bytes = None) -> dict:
    """
    Validation complète d'un PPTX décompressé.

    Exécute dans l'ordre :
    1. Auto-repair (xml:space="preserve" manquant)
    2. Checks structurels (XML, namespaces, IDs, refs, Content_Types, layouts)
    3. Validation XSD (chaque XML contre son schema Office)

    Args:
        unpacked_dir: chemin du PPTX décompressé (dossier avec ppt/, [Content_Types].xml, etc.)
        original_bytes: bytes du PPTX original (optionnel). Si fourni, les erreurs XSD
                        déjà présentes dans l'original sont ignorées — on ne remonte que
                        les NOUVELLES erreurs introduites par nos modifications.

    Returns:
        dict avec :
        - valid (bool) : True si toutes les validations passent
        - repairs (int) : nombre de réparations auto-appliquées
        - errors (list[str]) : erreurs structurelles
        - xsd_errors (list[str]) : erreurs XSD (nouvelles uniquement si original fourni)
    """
    path = Path(unpacked_dir)
    xml_files = list(path.rglob("*.xml")) + list(path.rglob("*.rels"))

    # --- Auto-repair ---
    repairs = _repair_whitespace(xml_files)

    # --- Checks structurels ---
    errors = []

    # 1. XML bien formé — si ça échoue, les autres checks vont planter
    xml_errors = _check_wellformed_xml(xml_files, path)
    if xml_errors:
        return {"valid": False, "repairs": repairs, "errors": xml_errors, "xsd_errors": []}

    # 2-8. Checks structurels
    errors += _check_namespaces(xml_files, path)
    errors += _check_unique_ids(xml_files, path)
    errors += _check_file_references(xml_files, path)
    errors += _check_content_types(path)
    errors += _check_slide_layout_ids(path)
    errors += _check_no_duplicate_layouts(path)
    errors += _check_notes_slides(path)

    # --- Validation XSD ---
    xsd_errors = _check_xsd(xml_files, path, original_bytes)

    return {
        "valid": len(errors) == 0 and len(xsd_errors) == 0,
        "repairs": repairs,
        "errors": errors,
        "xsd_errors": xsd_errors,
    }


# ============================================================
# Auto-repair
# ============================================================

def _repair_whitespace(xml_files: list[Path]) -> int:
    """
    Ajoute xml:space="preserve" sur les <a:t> dont le texte commence
    ou finit par un espace. Sans ça, PowerPoint supprime silencieusement
    ces espaces à l'ouverture.
    """
    repairs = 0
    for xml_file in xml_files:
        try:
            content = xml_file.read_text(encoding="utf-8")
            dom = defusedxml.minidom.parseString(content)
            modified = False

            for elem in dom.getElementsByTagName("*"):
                if elem.tagName.endswith(":t") and elem.firstChild:
                    text = elem.firstChild.nodeValue
                    if text and (text.startswith((" ", "\t")) or text.endswith((" ", "\t"))):
                        if elem.getAttribute("xml:space") != "preserve":
                            elem.setAttribute("xml:space", "preserve")
                            repairs += 1
                            modified = True

            if modified:
                xml_file.write_bytes(dom.toxml(encoding="UTF-8"))
        except Exception:
            logger.debug("Skipping whitespace repair for: %s", xml_file.name)
    return repairs


# ============================================================
# Checks structurels
# ============================================================

def _check_wellformed_xml(xml_files: list[Path], base: Path) -> list[str]:
    """Vérifie que chaque fichier XML est parsable."""
    errors = []
    for f in xml_files:
        try:
            lxml.etree.parse(str(f))
        except lxml.etree.XMLSyntaxError as e:
            errors.append(f"XML invalide — {f.relative_to(base)}: ligne {e.lineno}: {e.msg}")
    return errors


def _check_namespaces(xml_files: list[Path], base: Path) -> list[str]:
    """Vérifie que mc:Ignorable ne référence pas de préfixes non déclarés."""
    errors = []
    for f in xml_files:
        try:
            root = lxml.etree.parse(str(f)).getroot()
            declared = set(root.nsmap.keys()) - {None}
            for attr_val in [v for k, v in root.attrib.items() if k.endswith("Ignorable")]:
                for ns in set(attr_val.split()) - declared:
                    errors.append(
                        f"Namespace non déclaré — {f.relative_to(base)}: "
                        f"'{ns}' dans Ignorable mais pas déclaré"
                    )
        except lxml.etree.XMLSyntaxError:
            continue
    return errors


def _check_unique_ids(xml_files: list[Path], base: Path) -> list[str]:
    """
    Vérifie l'unicité des IDs critiques :
    - sp, pic, cxnSp, grpSp (shape IDs) : uniques par fichier
    - sldMasterId, sldLayoutId : uniques globalement
    - sldId : unique par fichier
    """
    FILE_SCOPE = {"sldid", "sp", "pic", "cxnsp", "grpsp"}
    GLOBAL_SCOPE = {"sldmasterid", "sldlayoutid"}

    errors = []
    global_ids = {}  # id_value → (fichier, tag)

    for f in xml_files:
        try:
            root = lxml.etree.parse(str(f)).getroot()
            file_ids = {}  # id_value → tag

            for elem in root.iter():
                tag = elem.tag.split("}")[-1].lower() if "}" in elem.tag else elem.tag.lower()
                if tag not in FILE_SCOPE and tag not in GLOBAL_SCOPE:
                    continue

                # Chercher l'attribut "id"
                id_value = None
                for attr, value in elem.attrib.items():
                    attr_local = attr.split("}")[-1].lower() if "}" in attr else attr.lower()
                    if attr_local == "id":
                        id_value = value
                        break
                if id_value is None:
                    continue

                rel = f.relative_to(base)

                if tag in GLOBAL_SCOPE:
                    if id_value in global_ids:
                        prev_file, prev_tag = global_ids[id_value]
                        errors.append(
                            f"ID dupliqué (global) — {rel}: <{tag}> id='{id_value}' "
                            f"déjà utilisé dans {prev_file} (<{prev_tag}>)"
                        )
                    else:
                        global_ids[id_value] = (rel, tag)
                else:
                    if id_value in file_ids:
                        errors.append(
                            f"ID dupliqué (fichier) — {rel}: <{tag}> id='{id_value}' "
                            f"déjà utilisé par <{file_ids[id_value]}>"
                        )
                    else:
                        file_ids[id_value] = tag
        except Exception:
            logger.debug("Skipping ID check for: %s", f.name)
            continue
    return errors


def _check_file_references(xml_files: list[Path], base: Path) -> list[str]:
    """Vérifie que chaque Target dans les .rels pointe vers un fichier existant."""
    errors = []
    for rels_file in [f for f in xml_files if f.suffix == ".rels"]:
        try:
            root = lxml.etree.parse(str(rels_file)).getroot()
            for rel in root.findall(f".//{{{PKG_RELS_NS}}}Relationship"):
                target = rel.get("Target", "")
                if not target or target.startswith(("http", "mailto:")):
                    continue

                if target.startswith("/"):
                    target_path = base / target.lstrip("/")
                elif rels_file.name == ".rels":
                    target_path = base / target
                else:
                    target_path = rels_file.parent.parent / target

                try:
                    if not target_path.resolve().exists():
                        errors.append(
                            f"Référence cassée — {rels_file.relative_to(base)}: "
                            f"'{target}' n'existe pas"
                        )
                except (OSError, ValueError):
                    errors.append(
                        f"Référence invalide — {rels_file.relative_to(base)}: '{target}'"
                    )
        except Exception as e:
            errors.append(f"Erreur parsing — {rels_file.relative_to(base)}: {e}")
    return errors


def _check_content_types(base: Path) -> list[str]:
    """Vérifie que les fichiers importants sont déclarés dans [Content_Types].xml."""
    errors = []
    ct_path = base / "[Content_Types].xml"
    if not ct_path.exists():
        return ["[Content_Types].xml introuvable"]

    try:
        root = lxml.etree.parse(str(ct_path)).getroot()
        declared = set()
        for override in root.findall(f".//{{{CONTENT_TYPES_NS}}}Override"):
            part = override.get("PartName", "").lstrip("/")
            if part:
                declared.add(part)

        important_roots = {"sld", "sldLayout", "sldMaster", "presentation", "theme"}
        for xml_file in base.rglob("*.xml"):
            if "_rels" in xml_file.parts or xml_file.name == "[Content_Types].xml":
                continue
            if "docProps" in xml_file.parts:
                continue
            try:
                file_root = lxml.etree.parse(str(xml_file)).getroot()
                root_name = file_root.tag.split("}")[-1] if "}" in file_root.tag else file_root.tag
                rel_path = str(xml_file.relative_to(base)).replace("\\", "/")
                if root_name in important_roots and rel_path not in declared:
                    errors.append(
                        f"Content_Types manquant — {rel_path} (root: <{root_name}>) "
                        f"pas déclaré dans [Content_Types].xml"
                    )
            except Exception:
                logger.debug("Skipping content type check for: %s", xml_file.name)
                continue
    except Exception as e:
        errors.append(f"Erreur parsing [Content_Types].xml: {e}")
    return errors


def _check_slide_layout_ids(base: Path) -> list[str]:
    """Vérifie que les sldLayoutId dans les slideMasters référencent des relations existantes."""
    errors = []
    for master in base.glob("ppt/slideMasters/*.xml"):
        try:
            root = lxml.etree.parse(str(master)).getroot()
            rels_file = master.parent / "_rels" / f"{master.name}.rels"
            if not rels_file.exists():
                errors.append(f"Fichier .rels manquant pour {master.relative_to(base)}")
                continue

            rels_root = lxml.etree.parse(str(rels_file)).getroot()
            valid_rids = set()
            for rel in rels_root.findall(f".//{{{PKG_RELS_NS}}}Relationship"):
                if "slideLayout" in rel.get("Type", ""):
                    valid_rids.add(rel.get("Id"))

            for layout_id_elem in root.findall(f".//{{{PML_NS}}}sldLayoutId"):
                rid = layout_id_elem.get(f"{{{OFFICE_RELS_NS}}}id")
                if rid and rid not in valid_rids:
                    errors.append(
                        f"Layout ID invalide — {master.relative_to(base)}: "
                        f"r:id='{rid}' introuvable dans les relations"
                    )
        except Exception:
            logger.debug("Skipping layout ID check for: %s", master.name)
            continue
    return errors


def _check_no_duplicate_layouts(base: Path) -> list[str]:
    """Vérifie que chaque slide a exactement un slideLayout dans ses .rels."""
    errors = []
    for rels_file in base.glob("ppt/slides/_rels/*.xml.rels"):
        try:
            root = lxml.etree.parse(str(rels_file)).getroot()
            layout_count = sum(
                1 for rel in root.findall(f".//{{{PKG_RELS_NS}}}Relationship")
                if "slideLayout" in rel.get("Type", "")
            )
            if layout_count > 1:
                errors.append(
                    f"Layouts dupliqués — {rels_file.relative_to(base)}: "
                    f"{layout_count} slideLayout (attendu: 1)"
                )
        except Exception:
            logger.debug("Skipping duplicate layout check for: %s", rels_file.name)
            continue
    return errors


def _check_notes_slides(base: Path) -> list[str]:
    """Vérifie que chaque notesSlide n'est référencée que par une seule slide."""
    errors = []
    notes_refs = {}  # target → [slide1, slide2, ...]

    for rels_file in base.glob("ppt/slides/_rels/*.xml.rels"):
        try:
            root = lxml.etree.parse(str(rels_file)).getroot()
            slide_name = rels_file.stem.replace(".xml", "")
            for rel in root.findall(f".//{{{PKG_RELS_NS}}}Relationship"):
                if "notesSlide" in rel.get("Type", ""):
                    target = rel.get("Target", "").replace("../", "")
                    notes_refs.setdefault(target, []).append(slide_name)
        except Exception:
            logger.debug("Skipping notes check for: %s", rels_file.name)
            continue

    for target, slides in notes_refs.items():
        if len(slides) > 1:
            errors.append(
                f"Note partagée — {target} référencée par {len(slides)} slides: "
                f"{', '.join(slides)}"
            )
    return errors


# ============================================================
# Validation XSD
# ============================================================
#
# Chaque fichier XML d'un PPTX a un schema XSD correspondant
# défini par la norme Office Open XML (ISO/IEC 29500).
# Par exemple :
#   - ppt/slides/slide1.xml  → pml.xsd (PresentationML)
#   - ppt/theme/theme1.xml   → dml-main.xsd (DrawingML)
#   - [Content_Types].xml    → opc-contentTypes.xsd
#   - *.rels                 → opc-relationships.xsd
#
# Le processus :
# 1. Trouver le schema correspondant au fichier
# 2. Nettoyer le XML (retirer extensions Microsoft non-standard)
# 3. Valider contre le schema
# 4. Si on a le fichier original, comparer les erreurs :
#    on ne remonte que les NOUVELLES erreurs (pas celles pré-existantes)
# ============================================================

def _check_xsd(xml_files: list[Path], base: Path, original_bytes: bytes = None) -> list[str]:
    """
    Valide chaque fichier XML contre son schema XSD Office.

    Si original_bytes est fourni, on décompresse le PPTX original dans un dossier
    temporaire et on compare : seules les erreurs NOUVELLES sont remontées.
    Ça évite de remonter des erreurs qui existaient déjà dans le template d'origine.
    """
    try:
        schemas_dir = _find_schemas_dir()
    except FileNotFoundError:
        return ["XSD: dossier schemas/ introuvable — validation XSD skippée"]

    errors = []

    # Si on a l'original, on le décompresse pour pouvoir comparer
    original_dir = None
    temp_obj = None
    if original_bytes:
        temp_obj = tempfile.TemporaryDirectory()
        original_dir = Path(temp_obj.name)
        try:
            import io
            with zipfile.ZipFile(io.BytesIO(original_bytes), "r") as zf:
                zf.extractall(original_dir)
        except Exception:
            original_dir = None

    try:
        for xml_file in xml_files:
            relative = xml_file.relative_to(base)

            # Trouver le schema
            schema_path = _get_schema_path(xml_file, base, schemas_dir)
            if schema_path is None:
                continue  # Pas de schema pour ce type de fichier → on skip

            # Valider le fichier modifié
            current_errors = _validate_one_file_xsd(xml_file, base, schema_path)
            if current_errors is None:
                continue  # Erreur de parsing du schema → skip
            if not current_errors:
                continue  # Pas d'erreurs → OK

            # Si on a l'original, calculer les erreurs pré-existantes
            if original_dir:
                original_xml = original_dir / relative
                if original_xml.exists():
                    original_errors = _validate_one_file_xsd(original_xml, original_dir, schema_path)
                    original_errors = original_errors or set()
                    # Garder uniquement les NOUVELLES erreurs
                    current_errors = current_errors - original_errors

            # Filtrer les erreurs bénignes connues
            current_errors = {
                e for e in current_errors
                if not any(pattern in e for pattern in IGNORED_XSD_ERRORS)
            }

            if current_errors:
                errors.append(f"XSD — {relative}: {len(current_errors)} erreur(s)")
                for err in list(current_errors)[:3]:
                    truncated = err[:200] + "..." if len(err) > 200 else err
                    errors.append(f"  → {truncated}")

    finally:
        if temp_obj:
            temp_obj.cleanup()

    return errors


def _get_schema_path(xml_file: Path, base: Path, schemas_dir: Path) -> Path | None:
    """
    Trouve le schema XSD correspondant à un fichier XML.

    Logique :
    - Fichier nommé explicitement (app.xml, core.xml, etc.) → mapping direct
    - Fichier .rels → schema OPC relationships
    - Fichier dans ppt/charts/ → schema chart
    - Fichier dans ppt/theme/ → schema DrawingML
    - Fichier dans ppt/ (slides, layouts, masters) → schema PresentationML
    - Sinon → None (pas de validation XSD pour ce fichier)
    """
    # Mapping par nom de fichier
    if xml_file.name in SCHEMA_MAPPINGS:
        return schemas_dir / SCHEMA_MAPPINGS[xml_file.name]

    # Tous les .rels
    if xml_file.suffix == ".rels":
        return schemas_dir / SCHEMA_MAPPINGS[".rels"]

    # Charts
    if "charts/" in str(xml_file) and xml_file.name.startswith("chart"):
        return schemas_dir / SCHEMA_MAPPINGS["chart"]

    # Themes
    if "theme/" in str(xml_file) and xml_file.name.startswith("theme"):
        return schemas_dir / SCHEMA_MAPPINGS["theme"]

    # Fichiers dans ppt/ (slides, slideLayouts, slideMasters, etc.)
    try:
        relative = xml_file.relative_to(base)
        if relative.parts and relative.parts[0] == "ppt":
            return schemas_dir / SCHEMA_MAPPINGS["ppt"]
    except ValueError:
        pass

    return None


def _validate_one_file_xsd(xml_file: Path, base: Path, schema_path: Path) -> set[str] | None:
    """
    Valide un fichier XML contre un schema XSD.

    Avant validation, nettoie le XML :
    - Retire mc:Ignorable (Mark Compatibility, extensions Microsoft)
    - Retire les attributs/éléments de namespaces non-OOXML
    - Retire les tags {{template}} des nœuds texte

    Returns:
        set d'erreurs (vide si valide), ou None si le schema ne peut pas être chargé.
    """
    try:
        # Charger le schema XSD
        with open(schema_path, "rb") as xsd_fh:
            parser = lxml.etree.XMLParser()
            xsd_doc = lxml.etree.parse(xsd_fh, parser=parser, base_url=str(schema_path))
            schema = lxml.etree.XMLSchema(xsd_doc)
    except Exception:
        return None  # Schema invalide ou manquant → skip

    try:
        # Charger et pré-traiter le XML
        xml_doc = lxml.etree.parse(str(xml_file))
        xml_doc = _strip_template_tags(xml_doc)
        xml_doc = _strip_mc_ignorable(xml_doc)

        # Pour les fichiers dans ppt/, nettoyer les namespaces non-OOXML
        try:
            relative = xml_file.relative_to(base)
            if relative.parts and relative.parts[0] == "ppt":
                xml_doc = _strip_non_ooxml(xml_doc)
        except ValueError:
            pass

        # Valider
        if schema.validate(xml_doc):
            return set()
        else:
            return {error.message for error in schema.error_log}

    except Exception as e:
        return {str(e)}


def _strip_mc_ignorable(xml_doc: lxml.etree._ElementTree) -> lxml.etree._ElementTree:
    """Retire l'attribut mc:Ignorable du root element (Mark Compatibility)."""
    root = xml_doc.getroot()
    mc_attr = f"{{{MC_NS}}}Ignorable"
    if mc_attr in root.attrib:
        del root.attrib[mc_attr]
    return xml_doc


def _strip_non_ooxml(xml_doc: lxml.etree._ElementTree) -> lxml.etree._ElementTree:
    """
    Retire les attributs et éléments de namespaces non-OOXML standards.

    Microsoft ajoute des extensions propriétaires (a14:, a16:, etc.) que les
    schemas ISO ne connaissent pas. On les retire avant validation pour éviter
    des faux positifs.
    """
    xml_string = lxml.etree.tostring(xml_doc, encoding="unicode")
    root = lxml.etree.fromstring(xml_string)

    # Retirer les attributs non-OOXML
    for elem in root.iter():
        attrs_to_remove = [
            attr for attr in elem.attrib
            if "{" in attr and attr.split("}")[0][1:] not in OOXML_NAMESPACES
        ]
        for attr in attrs_to_remove:
            del elem.attrib[attr]

    # Retirer les éléments non-OOXML (récursif)
    _remove_non_ooxml_elements(root)

    return lxml.etree.ElementTree(root)


def _remove_non_ooxml_elements(parent):
    """Supprime récursivement les éléments enfants dont le namespace n'est pas OOXML."""
    to_remove = []
    for elem in list(parent):
        if not hasattr(elem, "tag") or callable(elem.tag):
            continue
        tag_str = str(elem.tag)
        if tag_str.startswith("{"):
            ns = tag_str.split("}")[0][1:]
            if ns not in OOXML_NAMESPACES:
                to_remove.append(elem)
                continue
        _remove_non_ooxml_elements(elem)

    for elem in to_remove:
        parent.remove(elem)


def _strip_template_tags(xml_doc: lxml.etree._ElementTree) -> lxml.etree._ElementTree:
    """
    Retire les {{tags}} de type template des attributs et textes
    (sauf des nœuds <a:t> / <w:t> qui sont du texte visible).
    Ces tags sont utilisés pour les templates dynamiques mais ne sont
    pas valides selon les schemas XSD.
    """
    template_re = re.compile(r"\{\{[^}]*\}\}")
    xml_string = lxml.etree.tostring(xml_doc, encoding="unicode")
    root = lxml.etree.fromstring(xml_string)

    for elem in root.iter():
        if not hasattr(elem, "tag") or callable(elem.tag):
            continue
        tag_str = str(elem.tag)
        # Ne pas toucher aux nœuds de texte visible
        if tag_str.endswith("}t") or tag_str == "t":
            continue
        if elem.text and template_re.search(elem.text):
            elem.text = template_re.sub("", elem.text)
        if elem.tail and template_re.search(elem.tail):
            elem.tail = template_re.sub("", elem.tail)

    return lxml.etree.ElementTree(root)
