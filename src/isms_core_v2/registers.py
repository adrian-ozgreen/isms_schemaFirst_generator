"""
registers.py

Helpers for maintaining CSV-based:
- Document Control Register
- Master Reference Register

These are designed to be called from the CLI after a successful
DocumentModel validation + Word generation.
"""

from __future__ import annotations

from pathlib import Path
from typing import Iterable, List, Dict
import csv

from .models import (
    DocumentModel,
    DocumentControlRegisterRow,
    ReferenceRegisterEntry,
)


# ---------------------------------------------------------------------------
# CSV helpers
# ---------------------------------------------------------------------------

def _ensure_parent_dir(path: Path) -> None:
    if not path.parent.exists():
        path.parent.mkdir(parents=True, exist_ok=True)


def _read_csv(path: Path) -> List[Dict[str, str]]:
    if not path.is_file():
        return []
    with path.open("r", encoding="utf-8", newline="") as f:
        reader = csv.DictReader(f)
        return list(reader)


def _write_csv(path: Path, fieldnames: Iterable[str], rows: List[Dict[str, str]]) -> None:
    _ensure_parent_dir(path)
    with path.open("w", encoding="utf-8", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        for row in rows:
            writer.writerow(row)


# ---------------------------------------------------------------------------
# Document Control Register
# ---------------------------------------------------------------------------

DCR_FIELDNAMES = [
    "doc_id",
    "title",
    "doc_type",
    "version",
    "status",
    "owner",
    "approver",
    "confidentiality",
    "date_completed",
    "next_review_date",
    "file_path",
    "notes",
]


def _dcr_row_from_model(model: DocumentModel, output_path: Path) -> DocumentControlRegisterRow:
    m = model.metadata
    return DocumentControlRegisterRow(
        doc_id=m.doc_id,
        title=m.title,
        doc_type=m.doc_type if m.doc_type in ("Policy", "Procedure", "Record", "Template", "Other") else "Other",
        version=m.version,
        status=m.status if m.status in ("Draft", "For Review", "Approved", "Superseded", "Obsolete") else m.status,
        owner=m.owner,
        approver=m.approver,
        confidentiality=getattr(m, "confidentiality", None),
        date_completed=getattr(m, "date_completed", None),
        next_review_date=getattr(m, "next_review_date", None),
        file_path=str(output_path),
        notes=None,
    )


def update_document_control_register(
    register_path: Path,
    model: DocumentModel,
    output_path: Path,
) -> None:
    """
    Upsert a row for this document into the Document Control Register CSV.

    - If doc_id already exists, its row is updated.
    - Otherwise, a new row is appended.
    """
    rows = _read_csv(register_path)

    dcr_row = _dcr_row_from_model(model, output_path)
    dcr_dict = dcr_row.model_dump()

    updated: List[Dict[str, str]] = []
    found = False
    for row in rows:
        if row.get("doc_id") == dcr_row.doc_id:
            # overwrite with current values
            updated.append({k: str(dcr_dict.get(k, "")) for k in DCR_FIELDNAMES})
            found = True
        else:
            # normalise existing row to current fieldnames
            updated.append({k: row.get(k, "") for k in DCR_FIELDNAMES})

    if not found:
        updated.append({k: str(dcr_dict.get(k, "")) for k in DCR_FIELDNAMES})

    _write_csv(register_path, DCR_FIELDNAMES, updated)


# ---------------------------------------------------------------------------
# Master Reference Register
# ---------------------------------------------------------------------------

MRR_FIELDNAMES = [
    "ref_id",
    "source_doc_id",
    "source_doc_title",
    "source_section_key",
    "ref_type",
    "target_identifier",
    "target_title",
    "target_version",
    "target_location",
    "notes",
]


def _next_ref_id(existing_rows: List[Dict[str, str]]) -> str:
    """
    Generate a simple monotonically increasing REF-000001 style ID.
    """
    prefix = "REF-"
    max_n = 0
    for row in existing_rows:
        ref_id = row.get("ref_id") or ""
        if not ref_id.startswith(prefix):
            continue
        tail = ref_id[len(prefix):]
        if tail.isdigit():
            n = int(tail)
            if n > max_n:
                max_n = n
    return f"{prefix}{max_n + 1:06d}"


def _mrr_entries_from_model(model: DocumentModel) -> List[ReferenceRegisterEntry]:
    """
    Very first cut:

    - Seed references from metadata.related_documents (strings).
    - Later we can extend this to scan dedicated JSON reference blocks
      at section level.
    """
    m = model.metadata
    entries: List[ReferenceRegisterEntry] = []

    for ref in m.related_documents or []:
        ref_str = str(ref).strip()
        if not ref_str:
            continue

        entries.append(
            ReferenceRegisterEntry(
                ref_id="",  # filled later
                source_doc_id=m.doc_id,
                source_doc_title=m.title,
                source_section_key=None,
                ref_type="InternalDocument",
                target_identifier=ref_str,
                target_title=None,
                target_version=None,
                target_location=None,
                notes=None,
            )
        )

    return entries


def update_master_reference_register(
    register_path: Path,
    model: DocumentModel,
) -> None:
    """
    Append new reference entries for this document into the Master Reference Register CSV.

    v0 behaviour:
    - Does NOT attempt to de-duplicate across runs.
      (Safe because REF IDs are unique; later we can add smarter merging.)
    """
    existing = _read_csv(register_path)

    # Convert existing rows back to dicts with all known fields
    normalised_existing: List[Dict[str, str]] = [
        {k: row.get(k, "") for k in MRR_FIELDNAMES} for row in existing
    ]

    entries = _mrr_entries_from_model(model)

    # Assign ref_ids and append
    for entry in entries:
        entry.ref_id = _next_ref_id(normalised_existing)
        normalised_existing.append(
            {k: str(entry.model_dump().get(k, "")) for k in MRR_FIELDNAMES}
        )

    _write_csv(register_path, MRR_FIELDNAMES, normalised_existing)
