
from typing import List, Dict, Any
from docx.document import Document as _Document
from docx.oxml.ns import qn

def _iter_body_elements(document: _Document):
    """Yield paragraphs and tables in document order."""
    body = document.element.body
    # Build paragraph and table index for quick identity mapping
    para_map = {p._p: p for p in document.paragraphs}
    table_map = {t._tbl: t for t in document.tables}
    for child in body.iterchildren():
        if child.tag == qn('w:p'):
            p = para_map.get(child)
            if p is not None:
                yield p
        elif child.tag == qn('w:tbl'):
            t = table_map.get(child)
            if t is not None:
                yield t

def _find_table_after_heading(document: _Document, heading_text: str):
    target = (heading_text or "").strip().lower()
    seen_heading = False
    for el in _iter_body_elements(document):
        if hasattr(el, "text"):  # paragraph
            if (el.text or "").strip().lower() == target:
                seen_heading = True
        else:
            if seen_heading:
                return el
    return None

def _ensure_columns(table, min_cols: int):
    for row in table.rows:
        while len(row.cells) < min_cols:
            row.add_cell()

def _populate_kv_table(table, data: Dict[str, Any]):
    """Populate a 2-col key/value table with provided dict."""
    _ensure_columns(table, 2)
    # Build existing index
    existing = {}
    for i, row in enumerate(table.rows):
        key = (row.cells[0].text or "").strip()
        if key:
            existing[key] = i
    # Write values
    for k, v in data.items():
        if k in existing:
            idx = existing[k]
        else:
            row = table.add_row()
            row.cells[0].text = k
            idx = len(table.rows) - 1
        table.rows[idx].cells[0].text = k
        table.rows[idx].cells[1].text = str(v) if v is not None else ""

def populate_distribution_list(doc: _Document, rows: List[Dict[str, str]]) -> bool:
    """
    Fill the table after 'Distribution List' with columns: Recipient, Role/Dept, Method, Notes
    Creates rows for each entry in rows.
    """
    if not rows:
        return False
    t = _find_table_after_heading(doc, "Distribution List")
    if t is None:
        return False
    _ensure_columns(t, 4)
    # optional: detect header row if present; otherwise add one
    header = ["Recipient", "Role/Dept", "Method", "Notes"]
    if len(t.rows) == 0 or "recipient" not in (t.rows[0].cells[0].text or "").lower():
        hr = t.add_row()
        for i, h in enumerate(header):
            hr.cells[i].text = h
    # add data rows
    for r in rows:
        row = t.add_row()
        row.cells[0].text = str(r.get("recipient",""))
        row.cells[1].text = str(r.get("role_or_dept",""))
        row.cells[2].text = str(r.get("method",""))
        row.cells[3].text = str(r.get("notes",""))
    return True

def populate_approval_signatures(doc: _Document, approvals: List[Dict[str, str]]) -> bool:
    """
    Fill the table after 'Approval Signatures' with columns: Name, Role, Signature, Date
    """
    if not approvals:
        return False
    t = _find_table_after_heading(doc, "Approval Signatures")
    if t is None:
        return False
    _ensure_columns(t, 4)
    header = ["Name", "Role", "Signature", "Date"]
    if len(t.rows) == 0 or "name" not in (t.rows[0].cells[0].text or "").lower():
        hr = t.add_row()
        for i, h in enumerate(header):
            hr.cells[i].text = h
    for a in approvals:
        row = t.add_row()
        row.cells[0].text = str(a.get("name",""))
        row.cells[1].text = str(a.get("role",""))
        row.cells[2].text = str(a.get("signature",""))
        row.cells[3].text = str(a.get("date",""))
    return True

def populate_document_classification(doc: _Document, data: Dict[str, Any]) -> bool:
    """
    Fill KV table after 'Document Classification' (2 columns).
    Expected keys may include: 'Classification', 'Handling', 'Owner Dept', etc.
    """
    if not data:
        return False
    t = _find_table_after_heading(doc, "Document Classification")
    if t is None:
        return False
    _populate_kv_table(t, data)
    return True

def populate_handling_requirements(doc: _Document, data: Dict[str, Any]) -> bool:
    """
    Fill KV table after 'Handling Requirements' (2 columns).
    """
    if not data:
        return False
    t = _find_table_after_heading(doc, "Handling Requirements")
    if t is None:
        return False
    _populate_kv_table(t, data)
    return True

def populate_retention_period(doc: _Document, data: Dict[str, Any]) -> bool:
    """
    Fill KV table after 'Retention Period' (2 columns).
    Example keys: 'Minimum Retention', 'System of Record', 'Disposition Owner'
    """
    if not data:
        return False
    t = _find_table_after_heading(doc, "Retention Period")
    if t is None:
        return False
    _populate_kv_table(t, data)
    return True
