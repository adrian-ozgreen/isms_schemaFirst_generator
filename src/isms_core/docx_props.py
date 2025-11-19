
from pathlib import Path
from zipfile import ZipFile
import zipfile
import time
import os
import xml.etree.ElementTree as ET
from datetime import datetime


# Namespaces
NS = {
    "cp": "http://schemas.openxmlformats.org/package/2006/metadata/core-properties",
    "dc": "http://purl.org/dc/elements/1.1/",
    "dcterms": "http://purl.org/dc/terms/",
    "vt": "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes",
    "cust": "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties",
    "xsi": "http://www.w3.org/2001/XMLSchema-instance"
}

ET.register_namespace('cp', NS['cp'])
ET.register_namespace('dc', NS['dc'])
ET.register_namespace('dcterms', NS['dcterms'])
ET.register_namespace('vt', NS['vt'])
ET.register_namespace('cust', NS['cust'])
ET.register_namespace('xsi', NS['xsi'])

CORE_PATH = "docProps/core.xml"
CUSTOM_PATH = "docProps/custom.xml"

CUSTOM_NS = "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties"
VT_NS = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"

ET.register_namespace("", CUSTOM_NS)
ET.register_namespace("vt", VT_NS)

ET.register_namespace("", CUSTOM_NS)
ET.register_namespace("vt", VT_NS)



def _read_xml_from_zip(z: ZipFile, path: str) -> ET.ElementTree | None:
    try:
        with z.open(path) as f:
            return ET.parse(f)
    except KeyError:
        return None



def _atomic_replace(src_path: Path, dst_path: Path, attempts: int = 8, delay: float = 0.25):
    """
    Robust replace on Windows: retry, attempt to remove destination if needed,
    and fall back to writing a side-by-side output if still locked.
    """
    src_path = Path(src_path)
    dst_path = Path(dst_path)
    last_err = None
    for _ in range(attempts):
        try:
            os.replace(src_path, dst_path)
            return dst_path
        except PermissionError as e:
            last_err = e
            try:
                if dst_path.exists():
                    os.remove(dst_path)
            except Exception:
                pass
            time.sleep(delay)
    # Fallback: keep a side-by-side file
    fallback = dst_path.with_name(dst_path.stem + "_props" + dst_path.suffix)
    try:
        if fallback.exists():
            os.remove(fallback)
        os.replace(src_path, fallback)
        return fallback
    except Exception as e2:
        # As a last resort, try copying bytes
        import shutil
        shutil.copyfile(src_path, fallback)
        try:
            os.remove(src_path)
        except Exception:
            pass
        return fallback


def _write_xml_to_zip(zdst_path: Path, xml_path: str, tree: ET.ElementTree):
    # Rebuild the docx with updated XML because in-place editing isn't supported
    tmp = zdst_path.with_suffix(".tmp")
    # Read all existing entries first, then write a new zip
    with ZipFile(zdst_path, "r") as zsrc:
        entries = [(i, zsrc.read(i.filename)) for i in zsrc.infolist()]
    with ZipFile(tmp, "w") as zout:
        for item, data in entries:
            if item.filename == xml_path:
                xml_bytes = ET.tostring(tree.getroot(), encoding="utf-8", xml_declaration=True)
                zout.writestr(item, xml_bytes)
            else:
                zout.writestr(item, data)
    # Replace original
    _atomic_replace(tmp, zdst_path)

def _ensure_custom_root(tree: ET.ElementTree | None) -> ET.ElementTree:
    if tree is None:
        root = ET.Element(ET.QName(NS['cust'], "Properties"))
        return ET.ElementTree(root)
    return tree

def _find_custom_prop(root: ET.Element, name: str):
    for p in root.findall(f"cust:property", NS):
        if p.get('name') == name:
            return p
    return None

def _set_custom_prop(root: ET.Element, pid: int, name: str, value: str, vtype: str = "lpwstr"):
    # vtype can be one of vt types; we use lpwstr for strings, filetime for dates
    node = _find_custom_prop(root, name)
    if node is None:
        node = ET.SubElement(root, ET.QName(NS['cust'], "property"), {
            "fmtid": "{D5CDD502-2E9C-101B-9397-08002B2CF9AE}",
            "pid": str(pid),
            "name": name
        })
    # Remove existing child
    for c in list(node):
        node.remove(c)
    if vtype == "filetime":
        v = ET.SubElement(node, ET.QName(NS['vt'], "filetime"))
        v.text = value  # must be ISO 8601 like 2025-11-12T00:00:00Z
    else:
        v = ET.SubElement(node, ET.QName(NS['vt'], "lpwstr"))
        v.text = value

def _set_core_props(core_tree: ET.ElementTree, title: str | None, subject: str | None = None):
    root = core_tree.getroot()
    # dc:title
    if title is not None:
        t = root.find("dc:title", NS)
        if t is None:
            t = ET.SubElement(root, ET.QName(NS['dc'], "title"))
        t.text = title
    # dc:subject (optional)
    if subject is not None:
        s = root.find("dc:subject", NS)
        if s is None:
            s = ET.SubElement(root, ET.QName(NS['dc'], "subject"))
        s.text = subject
    # cp:lastModifiedBy timestamp
    lm = root.find("cp:lastModifiedBy", NS)
    if lm is None:
        lm = ET.SubElement(root, ET.QName(NS['cp'], "lastModifiedBy"))
    lm.text = "ISMS Hybrid Generator"
    mod = root.find("dcterms:modified", NS)
    if mod is None:
        mod = ET.SubElement(root, ET.QName(NS['dcterms'], 'modified'))
    # set xsi:type="dcterms:W3CDTF"
    mod.set(ET.QName(NS['xsi'], 'type'), 'dcterms:W3CDTF')
    mod.text = datetime.utcnow().strftime("%Y-%m-%dT%H:%M:%SZ")

def set_doc_properties(docx_path: Path, meta: dict) -> None:
    """
    Update Word custom properties in docProps/custom.xml for key ISMS fields.
    This function overwrites existing properties with the same name and
    creates them if missing.
    """
    docx_path = Path(docx_path)

    # Read existing docx
    with zipfile.ZipFile(docx_path, "r") as zin:
        namelist = zin.namelist()
        custom_bytes = zin.read(CUSTOM_PATH) if CUSTOM_PATH in namelist else None
        other_files = {name: zin.read(name) for name in namelist if name != CUSTOM_PATH}

    # Build or parse custom.xml
    if custom_bytes is None:
        root = ET.Element(f"{{{CUSTOM_NS}}}Properties")
    else:
        root = ET.fromstring(custom_bytes)

    def max_pid() -> int:
        m = 1
        for p in root.findall(f"{{{CUSTOM_NS}}}property"):
            try:
                pid = int(p.get("pid", "1"))
                m = max(m, pid)
            except ValueError:
                continue
        return m

    def upsert(name: str, value: str | None):
        if value is None:
            return
        # find existing
        prop = None
        for p in root.findall(f"{{{CUSTOM_NS}}}property"):
            if p.get("name") == name:
                prop = p
                break
        if prop is None:
            pid = max_pid() + 1
            prop = ET.SubElement(
                root,
                f"{{{CUSTOM_NS}}}property",
                {
                    "fmtid": "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}",
                    "pid": str(pid),
                    "name": name,
                },
            )
        # replace value child
        for child in list(prop):
            prop.remove(child)
        v = ET.SubElement(prop, f"{{{VT_NS}}}lpwstr")
        v.text = str(value)

    # ---- map metadata to your existing property names ----
    upsert("DocID",            meta.get("doc_id"))
    upsert("Version",          meta.get("version"))
    upsert("Owner",            meta.get("owner"))
    upsert("ApprovedBy",       meta.get("approved_by"))
    upsert("Status",           meta.get("status"))
    upsert("DocumentType",     meta.get("document_type"))

    # dates as simple text (avoid weird filetime conversions)
    dc = meta.get("date_completed")
    nr = meta.get("next_review_date")
    upsert("DateCompleted",    dc)
    upsert("Date completed",   dc)
    upsert("NextReviewDate",   nr)

    # confidentiality – if you want from classification:
    conf = meta.get("confidentiality")
    upsert("Confidentiality",  conf)

    new_custom = ET.tostring(root, encoding="utf-8", xml_declaration=True)

    # Write new docx
    tmp = docx_path.with_suffix(".tmpdocx")
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for name, data in other_files.items():
            zout.writestr(name, data)
        zout.writestr(CUSTOM_PATH, new_custom)

    os.replace(tmp, docx_path)