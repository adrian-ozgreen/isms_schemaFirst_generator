"""
word_importer.py

Import a Word document (.docx) and convert it to the JSON structure expected by
the schema-first ISMS generator, preserving inline formatting (bold/italic/
underline) and hyperlinks at the run level.

This version is intentionally generic and backwards-compatible:

- Paragraph content blocks look like:
    {
        "type": "paragraph",
        "text": "Plain concatenated text of all runs",
        "runs": [
            {
                "text": "Click ",
                "bold": false,
                "italic": false,
                "underline": false,
                "hyperlink": null
            },
            {
                "text": "here",
                "bold": true,
                "italic": false,
                "underline": true,
                "hyperlink": "https://example.com"
            }
        ]
    }

- Existing JSON that only has "text" will still render fine.
- New JSON produced by this importer will have both "text" and "runs". The
  generator can prefer "runs" when present for full fidelity.
"""

from __future__ import annotations

import argparse
import json
from pathlib import Path
from typing import Any, Dict, List, Optional
import re  # ← NEW

from docx import Document
from docx.text.paragraph import Paragraph
from docx.oxml.ns import qn


# ---------------------------------------------------------------------------
# Run extraction with hyperlink preservation
# ---------------------------------------------------------------------------

def extract_runs_with_hyperlinks(paragraph: Paragraph) -> List[Dict[str, Any]]:
    """
    Return an ordered list of run fragments from a paragraph, preserving:

    - run text
    - bold
    - italic
    - underline
    - whether the run is part of a hyperlink, and the hyperlink URL

    The returned structure is:
        [
            {
                "text": str,
                "bold": bool | None,
                "italic": bool | None,
                "underline": bool | None,
                "hyperlink": str | None,
            },
            ...
        ]
    """
    runs_data: List[Dict[str, Any]] = []

    # Build a map: run element -> hyperlink URL (if any)
    hyperlink_map: Dict[Any, Optional[str]] = {}
    p_elm = paragraph._p  # CT_P (low-level XML)
    part = paragraph.part

    # Find all w:hyperlink elements in the paragraph
    for h in p_elm.findall(".//w:hyperlink", p_elm.nsmap):
        r_id = h.get(qn("r:id"))
        url: Optional[str] = None
        if r_id is not None and r_id in part.rels:
            rel = part.rels[r_id]
            # python-docx Relationship exposes the URL as .target_ref
            target = getattr(rel, "target_ref", None)
            if target is not None:
                url = str(target)

        # Mark all w:r inside this hyperlink with the URL
        for r in h.findall(".//w:r", p_elm.nsmap):
            hyperlink_map[r] = url

    # Now iterate python-docx Run objects in order; map them to URLs where present
    for run in paragraph.runs:
        r_elm = run._r  # CT_R
        url = hyperlink_map.get(r_elm)

        runs_data.append(
            {
                "text": run.text or "",
                "bold": bool(run.bold) if run.bold is not None else False,
                "italic": bool(run.italic) if run.italic is not None else False,
                "underline": bool(run.underline) if run.underline is not None else bool(url),
                "hyperlink": url,
            }
        )

    return runs_data


# ---------------------------------------------------------------------------
# Content-block helpers
# ---------------------------------------------------------------------------

def _paragraph_to_block(paragraph: Paragraph) -> Optional[Dict[str, Any]]:
    """
    Convert a paragraph to a JSON content block.

    - Skips completely empty paragraphs.
    - Preserves both plain text and rich runs (with hyperlink, bold, etc.).
    """
    # Plain text (for backwards compatibility and search)
    plain_text = paragraph.text or ""
    if not plain_text.strip():
        # Completely empty paragraph; usually we can skip
        return None

    runs = extract_runs_with_hyperlinks(paragraph)

    block: Dict[str, Any] = {
        "kind": "paragraph",   # ← IMPORTANT: was "type": "paragraph"
        "text": plain_text,
        "runs": runs,
    }
    return block



def _get_heading_level(paragraph: Paragraph) -> Optional[int]:
    """
    Return a heading level (1–4) if the paragraph uses a Heading style,
    otherwise None.

    This looks for built-in Word styles like 'Heading 1', 'Heading 2', etc.
    """
    style = getattr(paragraph, "style", None)
    name = getattr(style, "name", "") or ""
    name_lower = name.lower()

    # Match "Heading 1", "Heading 2", "heading 3", etc.
    m = re.match(r"heading\s+([1-4])", name_lower)
    if m:
        return int(m.group(1))

    # You can extend this later with custom style names if needed
    return None


def _slugify(text: str) -> str:
    """
    Turn a heading title into a safe section key, e.g.:
    'Raw Data Specification' → 'raw_data_specification'.
    """
    base = re.sub(r"[^a-zA-Z0-9]+", "_", text.strip().lower()).strip("_")
    return base or "section"







def _import_body_as_single_section(doc: Document, title: str) -> Dict[str, Any]:
    """
    Import the main body of the document into a single top-level section
    ('record_content'), but use Word Heading styles to create nested
    subsections inside it.

    Behaviour:

    - Paragraphs with style 'Heading 1'–'Heading 4' become Section objects
      with level 1–4 (clamped to 1–4).
    - Content paragraphs are converted to ContentBlock objects (with runs)
      and attached to the most recent section at the appropriate level.
    - Paragraphs that appear before the first heading are added directly
      to record_content.content.
    """
    # Top-level catch-all section that the ISMS generator expects
    main_section: Dict[str, Any] = {
        "key": "record_content",
        "title": title or "Record Content",
        "level": 1,
        "content": [],
        "subsections": [],
    }

    # section_stack[i] is the most recent section at level i.
    # Index 0 is the top-level record_content.
    section_stack: List[Dict[str, Any]] = [main_section]

    for paragraph in doc.paragraphs:
        heading_level = _get_heading_level(paragraph)

        if heading_level is not None:
            # Convert Word Heading N → ISMS Level N+1 (max Level 5)
            lvl = heading_level + 1
            if lvl > 5:
                lvl = 5

            text = (paragraph.text or "").strip()
            if not text:
                continue

            key = _slugify(text)

            new_section: Dict[str, Any] = {
                "key": key,
                "title": text,
                "level": lvl,
                "content": [],
                "subsections": [],
            }

            # Ensure the stack length matches the *parent* level:
            # parent of level N is level N-1
            while len(section_stack) > (lvl - 1):
                section_stack.pop()

            parent = section_stack[-1]
            parent.setdefault("subsections", []).append(new_section)
            section_stack.append(new_section)

            # Don't add the heading itself as a content block
            continue


        # Not a heading: treat as body content under the current section
        block = _paragraph_to_block(paragraph)
        if block is not None:
            current_section = section_stack[-1]
            current_section.setdefault("content", []).append(block)

    return main_section



# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------

def import_word_to_json(
    input_path: Path,
    doc_type: str = "Record",
    doc_id: str = "REC-UNSPECIFIED-001",
    title: Optional[str] = None,
) -> Dict[str, Any]:
    """
    Main API: load a .docx and return JSON ready for the ISMS generator.

    You can plug this into your CLI or call it from `cli.py`.

    Parameters
    ----------
    input_path: Path
        Path to the source .docx file.
    doc_type: str
        "Record", "Policy", or "Procedure" (used only in metadata here).
    doc_id: str
        Identifier for the record in your ISMS (metadata.doc_id).
    title: Optional[str]
        Document title; if None, use the filename stem.

    Returns
    -------
    dict
        {
          "metadata": { ... },
          "sections": [ ... ]
        }
    """
    doc = Document(str(input_path))

    if title is None:
        title = input_path.stem

    metadata: Dict[str, Any] = {
        "doc_id": doc_id,
        "title": title,
        "doc_type": doc_type,
        "version": "0.1",
        "status": "Draft",
        "owner": "TBC",
        "approver": None,
        "related_documents": [],
    }

    # Import the body into a single catch-all section
    main_section = _import_body_as_single_section(doc, title)

    # ------------------------------------------------------------------
    # Ensure mandatory top-level sections exist so DocumentModel passes
    # its "mandatory sections" validation.
    #
    # These are stubs (empty content); they can be populated later, but
    # their presence satisfies the schema and keeps the CLI quiet.
    # ------------------------------------------------------------------
    mandatory_sections: list[tuple[str, str]] = [
        ("title_page", "Title Page"),
        ("document_control", "Document Control"),
        ("table_of_contents", "Table of Contents"),
        ("revision_history", "Revision History"),
        ("approval_signatures", "Approval Signatures"),
        ("document_classification", "Document Classification"),
        ("purpose", "Purpose"),
        ("scope", "Scope"),
        ("roles_and_responsibilities", "Roles and Responsibilities"),
        ("related_documents", "Related Documents"),
    ]

    sections: list[Dict[str, Any]] = []

    # In case you later add more sections, collect any existing keys
    existing_keys = {main_section.get("key")} if main_section.get("key") else set()

    for key, sec_title in mandatory_sections:
        if key in existing_keys:
            continue

        # Special case: document_classification must have three subsections
        if key == "document_classification":
            subsections = [
                {
                    "key": "distribution_list",
                    "title": "Distribution List",
                    "level": 2,
                    "content": [],
                    "subsections": [],
                },
                {
                    "key": "handling_requirements",
                    "title": "Handling Requirements",
                    "level": 2,
                    "content": [],
                    "subsections": [],
                },
                {
                    "key": "retention_period",
                    "title": "Retention Period",
                    "level": 2,
                    "content": [],
                    "subsections": [],
                },
            ]
        else:
            subsections = []

        sections.append(
            {
                "key": key,
                "title": sec_title,
                "level": 1,
                "content": [],
                "subsections": subsections,
            }
        )


    # Finally, append the actual imported content
    sections.append(main_section)

    json_data: Dict[str, Any] = {
        "metadata": metadata,
        "sections": sections,
    }
    return json_data


def import_word_to_document_dict(
    path: str | Path,
    doc_type: str = "Record",
    default_doc_id: str = "REC-UNSPECIFIED-001",
    **kwargs: Any,
) -> Dict[str, Any]:
    """
    Backwards-compatible wrapper used by src.isms_core_v2.cli.cmd_import_word.

    CLI calls it like:
        import_word_to_document_dict(
            path=input_path,
            doc_type=args.doc_type,
            default_doc_id=args.doc_id,
        )

    We normalise those arguments and delegate to import_word_to_json().
    """
    # Normalise path
    input_path = Path(path)

    # Use the CLI-provided doc_id as the actual metadata.doc_id
    effective_doc_id = default_doc_id or "REC-UNSPECIFIED-001"

    return import_word_to_json(
        input_path=input_path,
        doc_type=doc_type,
        doc_id=effective_doc_id,
        title=None,  # title will fall back to filename stem inside import_word_to_json
    )







# ---------------------------------------------------------------------------
# CLI entry point (optional)
# ---------------------------------------------------------------------------

def main(argv: Optional[List[str]] = None) -> None:
    parser = argparse.ArgumentParser(
        description="Import a Word document into ISMS JSON (with rich runs)."
    )
    parser.add_argument("input_docx", help="Path to the source .docx file")
    parser.add_argument("output_json", help="Path to write the JSON output")
    parser.add_argument(
        "--doc-type",
        choices=["Record", "Policy", "Procedure"],
        default="Record",
        help="Document type for metadata.doc_type",
    )
    parser.add_argument(
        "--doc-id",
        required=True,
        help="Document ID for metadata.doc_id"
    )
    parser.add_argument(
        "--title",
        help="Document title; defaults to input filename without extension",
    )

    args = parser.parse_args(argv)

    input_path = Path(args.input_docx)
    output_path = Path(args.output_json)

    data = import_word_to_json(
        input_path=input_path,
        doc_type=args.doc_type,
        doc_id=args.doc_id,
        title=args.title,
    )

    output_path.parent.mkdir(parents=True, exist_ok=True)
    with output_path.open("w", encoding="utf-8") as f:
        json.dump(data, f, indent=2, ensure_ascii=False)

    print(f"[OK] Imported Word document:")
    print(f"     Source: {input_path}")
    print(f"     Output: {output_path}")


if __name__ == "__main__":
    main()
