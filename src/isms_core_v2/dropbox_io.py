"""
dropbox_io.py

Helpers for copying generated ISMS artefacts into a Dropbox-synced
folder structure.

v0 behaviour:
- No Dropbox API, just filesystem paths (assumes Dropbox client syncs a folder).
- Classify output by doc_type into 01_Policies / 02_Procedures / 03_Records / 04_Templates.
- Optionally copy input JSON / source Word into a 90_Imports area.
"""

from __future__ import annotations

from pathlib import Path
from shutil import copy2
from typing import Optional

from .models import DocumentModel


# ---------------------------------------------------------------------------
# Folder classification
# ---------------------------------------------------------------------------

def _doc_type_subfolder(doc_type: str) -> str:
    """
    Map DocMetadata.doc_type → ISMS subfolder name.
    """
    dt = (doc_type or "").strip().lower()

    if dt == "policy":
        return "01_Policies"
    if dt == "procedure":
        return "02_Procedures"
    if dt == "record":
        return "03_Records"
    if dt == "template":
        return "04_Templates"

    # Fallback: treat unknown types as Records
    return "03_Records"


def _safe_filename(name: str) -> str:
    """
    Make a filename-safe string from DocID or title.
    """
    import re

    base = name.strip()
    # Replace path separators and other odd characters
    base = re.sub(r"[\\/:*?\"<>|]+", "_", base)
    # Avoid empty name
    return base or "document"


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------

def build_dropbox_output_path(
    dropbox_root: Path,
    model: DocumentModel,
    local_output_path: Path,
) -> Path:
    """
    Compute the path under the *ISMS root* where this generated doc
    should be stored.

    v0 scheme (after this change):
    <dropbox_root>/<type_subfolder>/<doc_id>.docx

    where --dropbox-root is already the ISMS root folder.
    """
    doc_meta = model.metadata

    # Print metadata for debugging
    print(f"DEBUG: doc_id={doc_meta.doc_id}, title={doc_meta.title}")

    # treat dropbox_root as the ISMS root
    isms_root = dropbox_root
    subfolder = _doc_type_subfolder(doc_meta.doc_type)
    target_dir = isms_root / subfolder
    
    stem = _safe_filename(doc_meta.doc_id or doc_meta.title)
    ext = local_output_path.suffix or ".docx"
    
    if not stem:
        print("WARNING: No valid filename stem could be derived! Check doc_id and title.")
    
    target_path = target_dir / f"{stem}{ext}"
    return target_path





def copy_generated_document_to_dropbox(
    dropbox_root: Path,
    model: DocumentModel,
    local_output_path: Path,
) -> Path:
    """
    Copy the generated Word document into the appropriate Dropbox
    subfolder. Returns the final Dropbox path.
    """
    target_path = build_dropbox_output_path(dropbox_root, model, local_output_path)

    if target_path is None:
        raise ValueError("Generated target path is None. Check metadata or path construction.")
    
    print(f"DEBUG: Writing to Dropbox path: {target_path}")
    
    target_path.parent.mkdir(parents=True, exist_ok=True)
    copy2(local_output_path, target_path)

    return target_path


def copy_inputs_to_dropbox(
    dropbox_root: Path,
    json_input_path: Optional[Path] = None,
    source_word_path: Optional[Path] = None,
) -> None:
    """
    v0 helper to optionally copy input artefacts into
    <dropbox_root>/90_Imports, where dropbox_root is the ISMS root.
    """
    imports_dir = dropbox_root / "90_Imports"
    imports_dir.mkdir(parents=True, exist_ok=True)
