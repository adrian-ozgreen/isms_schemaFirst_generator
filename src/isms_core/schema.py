# src/isms_core/schema.py
from __future__ import annotations
from pathlib import Path
from typing import Any, Dict, List, Tuple

StrDict = Dict[str, Any]

# ---------- helpers

def _is_nonempty_str(v: Any) -> bool:
    return isinstance(v, str) and v.strip() != ""

def _is_block_ref(v: Any) -> bool:
    return isinstance(v, dict) and set(v.keys()) == {"use_block"} and _is_nonempty_str(v.get("use_block"))

def _is_section_value(v: Any) -> bool:
    # Allowed shapes for any section: str, {"use_block": "<name>"}, list[str|dict]
    if _is_nonempty_str(v) or _is_block_ref(v):
        return True
    if isinstance(v, list):
        return all(isinstance(x, (str, dict)) for x in v)
    return False

def _kv_dict(v: Any) -> bool:
    return isinstance(v, dict) and all(isinstance(k, str) and isinstance(val, (str, int, float, bool, type(None))) for k, val in v.items())

def _list_of_dicts(v: Any, required_keys: List[str] | None = None) -> bool:
    if not isinstance(v, list):
        return False
    for item in v:
        if not isinstance(item, dict):
            return False
        if required_keys:
            if any(k not in item for k in required_keys):
                return False
    return True

def _list_of_rows(v: Any) -> bool:
    # rows in dynamic_tables: list[list|dict|str]
    return isinstance(v, list) and all(isinstance(x, (list, dict, str)) for x in v)

# ---------- main validation

def validate_input_payload(data: StrDict, profiles: StrDict, *, strict: bool = True) -> Tuple[List[str], List[str]]:
    """
    Returns (errors, warnings). If strict=True, caller may abort on errors.
    """
    errors: List[str] = []
    warns:  List[str] = []

    # --- metadata
    meta = data.get("metadata", {})
    if not isinstance(meta, dict):
        errors.append("metadata: must be an object")
        return errors, warns

    must = ["doc_id", "title", "version", "document_type"]
    for k in must:
        if not _is_nonempty_str(meta.get(k)):
            errors.append(f"metadata.{k}: required non-empty string")

    doc_type = (meta.get("document_type") or "").strip()
    if doc_type not in ("Policy", "Procedure", "Record"):
        errors.append("metadata.document_type: must be one of 'Policy' | 'Procedure' | 'Record'")

    # --- sections
    sections = data.get("sections", {})
    if not isinstance(sections, dict):
        errors.append("sections: must be an object")
        sections = {}

    # generic shape check
    for key, val in sections.items():
        if not _is_section_value(val):
            warns.append(f"sections.{key}: unrecognised shape (expected string, {{'use_block': name}}, or list)")

    # profile-based presence check
    if doc_type in profiles:
        prof = profiles[doc_type]
        req = [s.strip() for s in prof.get("required_sections", [])]
        opt = [s.strip() for s in prof.get("optional_sections", [])]
        # map canonical headings to JSON keys you use
        keymap = {
            "Purpose": "purpose",
            "Scope": "scope",
            "Policy Statements": "policy_statements",
            "Roles and Responsibilities": "roles_and_responsibilities",
            "Related Documents": "related_documents",
            "Definitions and Acronyms": "definitions_and_acronyms",
            "Normative References": "normative_references",
            "Compliance and Enforcement": "compliance_and_enforcement",
            "Data Retention": "data_retention",
            "Procedure Steps": "procedure_steps",
            "Record Content": "record_content",
        }
        present = {h for h, k in keymap.items() if k in sections and sections.get(k)}

        for h in req:
            if h not in present:
                errors.append(f"Missing required section: {h}")
        # warn on unknown-but-present sections
        for k in sections.keys():
            if k not in keymap.values():
                warns.append(f"sections.{k}: not referenced in profile mapping for {doc_type} (will still render if your renderer supports it)")

    # --- front-matter (optional but shape-checked)
    fm = {
        "document_classification": _kv_dict,
        "handling_requirements": _kv_dict,
        "retention_period": _kv_dict,
    }
    for k, pred in fm.items():
        val = data.get(k)
        if val is not None and not pred(val):
            errors.append(f"{k}: must be a simple key/value object")

    if (dl := data.get("distribution_list")) is not None:
        if not _list_of_dicts(dl, None):
            errors.append("distribution_list: must be a list of objects")
    if (ap := data.get("approvals")) is not None:
        if not _list_of_dicts(ap, None):
            errors.append("approvals: must be a list of objects")

    # --- dynamic_tables (optional)
    dyn = data.get("dynamic_tables")
    if dyn is not None:
        if not isinstance(dyn, list):
            errors.append("dynamic_tables: must be a list")
        else:
            for i, spec in enumerate(dyn):
                if not isinstance(spec, dict):
                    errors.append(f"dynamic_tables[{i}]: must be an object")
                    continue
                cols = spec.get("columns")
                rows = spec.get("rows")
                if cols is not None and not (isinstance(cols, list) and all(isinstance(c, str) for c in cols)):
                    errors.append(f"dynamic_tables[{i}].columns: must be a list of strings")
                if rows is not None and not _list_of_rows(rows):
                    errors.append(f"dynamic_tables[{i}].rows: must be a list of rows (list|dict|string)")
                tgt = spec.get("target")
                if tgt is not None and not (isinstance(tgt, dict) and ("after_heading" in tgt)):
                    warns.append(f"dynamic_tables[{i}].target: recommend {{'after_heading': 'Heading'}}")
                if "create_if_missing" in spec and not isinstance(spec["create_if_missing"], bool):
                    errors.append(f"dynamic_tables[{i}].create_if_missing: must be boolean")

    return errors, warns
