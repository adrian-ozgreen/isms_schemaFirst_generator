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
import re

from docx import Document
from docx.text.paragraph import Paragraph
from docx.table import Table
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
                "bold": bool,
                "italic": bool,
                "underline": bool,
                "hyperlink": str | None,
            },
            ...
        ]
    """
    runs_data: List[Dict[str, Any]] = []

    # ------------------------------------------------------------------
    # 1) Build a map: low-level w:r element -> hyperlink URL (if any)
    #
    #    We support two Word hyperlink encodings:
    #
    #    (a) Relationship-based hyperlinks:
    #        <w:hyperlink r:id="rId5"><w:r>...</w:r></w:hyperlink>
    #
    #    (b) Field-code-based hyperlinks ("HYPERLINK" fields):
    #        <w:fldChar w:fldCharType="begin"/>
    #        <w:instrText> HYPERLINK "https://..." </w:instrText>
    #        <w:fldChar w:fldCharType="separate"/>
    #        <w:r>Visible Text</w:r>
    #        <w:fldChar w:fldCharType="end"/>
    #
    #    We first collect (a), then add (b) only for runs that don't
    #    already have a mapping.
    # ------------------------------------------------------------------
    hyperlink_map: Dict[Any, Optional[str]] = {}
    p_elm = paragraph._p  # CT_P (low-level XML)
    part = paragraph.part

    # 1a) Relationship-based hyperlinks (<w:hyperlink r:id="...">)
    for h in p_elm.findall(".//w:hyperlink", p_elm.nsmap):
        r_id = h.get(qn("r:id"))
        url: Optional[str] = None
        if r_id is not None and r_id in part.rels:
            rel = part.rels[r_id]
            # python-docx exposes the target URL as .target_ref
            target = getattr(rel, "target_ref", None)
            if target is not None:
                url = str(target)

        # Mark all w:r inside this hyperlink with the URL
        for r in h.findall(".//w:r", p_elm.nsmap):
            hyperlink_map[r] = url

    # ------------------------------------------------------------------
    # 1b) Also detect field-code style hyperlinks (HYPERLINK fields)
    #     Word sometimes represents hyperlinks as fields instead of
    #     <w:hyperlink> elements. These typically look like:
    #
    #       <w:fldChar w:fldCharType="begin"/>
    #       <w:instrText> HYPERLINK "https://example.com" </w:instrText>
    #       <w:fldChar w:fldCharType="separate"/>
    #       <w:r>Visible text</w:r>
    #       <w:fldChar w:fldCharType="end"/>
    #
    #     We scan those field instructions and map all w:r nodes between
    #     'begin' and 'end' to the extracted URL, but ONLY if they do not
    #     already have a hyperlink from the relationship-based mapping
    #     above. That way, real <w:hyperlink> elements always take
    #     precedence.
    # ------------------------------------------------------------------
    current_field_link: Optional[str] = None
    inside_field = False

    for child in p_elm.iterchildren():
        tag = child.tag

        # Field begins
        if tag.endswith("fldChar") and child.get(qn("w:fldCharType")) == "begin":
            inside_field = True
            current_field_link = None

        # Field instructions – look for HYPERLINK "url"
        if tag.endswith("instrText") and inside_field:
            instr = (child.text or "").strip()
            if instr:
                m = re.search(r'HYPERLINK\s+"([^"]+)"', instr)
                if m:
                    current_field_link = m.group(1)

        # Field ends – stop applying the current link
        if tag.endswith("fldChar") and child.get(qn("w:fldCharType")) == "end":
            inside_field = False
            current_field_link = None

        # Apply the field-code hyperlink to runs that don't already
        # have a mapping from <w:hyperlink>.
        if tag.endswith("r") and inside_field and current_field_link:
            if child not in hyperlink_map:
                hyperlink_map[child] = current_field_link

    # ------------------------------------------------------------------
    # 2) Walk all w:r nodes in document order (NOT paragraph.runs)
    #    python-docx does not currently expose w:hyperlink runs via
    #    paragraph.runs, so we read directly from the XML tree.
    # ------------------------------------------------------------------
    nsmap = p_elm.nsmap
    for r in p_elm.findall(".//w:r", nsmap):
        # Collect the visible text for this run
        texts = [t.text or "" for t in r.findall(".//w:t", nsmap)]
        run_text = "".join(texts)
        if not run_text:
            continue  # skip empty runs

        # Run formatting from XML
        rPr = r.find("w:rPr", nsmap)
        bold = False
        italic = False
        underline = False

        if rPr is not None:
            if rPr.find("w:b", nsmap) is not None:
                bold = True
            if rPr.find("w:i", nsmap) is not None:
                italic = True
            if rPr.find("w:u", nsmap) is not None:
                underline = True

        url = hyperlink_map.get(r)
        if url:
            # Hyperlinks are usually underlined in the UI, but in case the
            # original formatting did something different, we still mark
            # underline=True whenever we know it's a hyperlink.
            underline = True

        runs_data.append(
            {
                "text": run_text,
                "bold": bold,
                "italic": italic,
                "underline": underline,
                "hyperlink": url,
            }
        )

    return runs_data


# ---------------------------------------------------------------------------
# Content-block helpers
# ---------------------------------------------------------------------------


def _detect_list_kind(paragraph: Paragraph) -> Optional[str]:
    """
    Best-effort detection of bullet vs numbered lists.

    Strategy:
    - Look at the style name (e.g. 'List Bullet', 'ISMS List Numbered').
    - If that fails, fall back to checking Word's numbering XML (w:numPr).
    """
    style = getattr(paragraph, "style", None)
    name = getattr(style, "name", "") or ""
    name_lower = name.lower()

    # 1) Style-based detection (works for 'List Bullet', 'ISMS List Bullet', etc.)
    if "bullet" in name_lower:
        return "bullet_list"
    if "number" in name_lower:
        return "numbered_list"

    # Treat generic 'List Paragraph' as bullet list
    if "list paragraph" in name_lower:
        return "bullet_list"

    # 2) Fallback: if the paragraph participates in a Word list
    p_elm = paragraph._p  # CT_P
    if p_elm is not None:
        num_pr = p_elm.find(".//w:numPr", p_elm.nsmap)
        if num_pr is not None:
            # If you want *all* lists to become bullet lists, change to "bullet_list"
            return "bullet_list"

    return None


def _paragraph_to_block(paragraph: Paragraph) -> Optional[Dict[str, Any]]:
    """
    Convert a paragraph to a JSON content block.

    - Skips completely empty paragraphs.
    - Detects bullet/numbered list items and emits 'bullet_list' /
      'numbered_list' blocks so the renderer can map them to ISMS list styles.
    - For any text-ish block we also capture rich runs, including hyperlinks.
    """
    plain_text = paragraph.text or ""
    if not plain_text.strip():
        # Completely empty paragraph; usually we can skip
        return None

    # Capture run fragments (with hyperlinks) up-front so we can reuse them
    runs = extract_runs_with_hyperlinks(paragraph)

    # 1) List items become 'bullet_list' / 'numbered_list' blocks.
    #    We keep the existing text representation so the renderer can
    #    still group and restart numbering, but now we also attach 'runs'
    #    so hyperlinks are preserved.
    list_kind = _detect_list_kind(paragraph)
    if list_kind is not None:
        return {
            "kind": list_kind,        # "bullet_list" or "numbered_list"
            "text": [plain_text],     # one item per block (as before)
            "runs": runs,
        }

    # 2) Normal paragraphs (non-list)
    return {
        "kind": "paragraph",
        "text": plain_text,
        "runs": runs,
    }





def _table_to_block(table: Table) -> Dict[str, Any]:
    """
    Convert a Word table into a 'table' content block.

    Structure matches what word_renderer._render_table_block expects:
      - block.header: List[str] for the header row
      - block.rows:   List[List[str]] for data rows
    """
    rows_data: List[List[str]] = []

    for row in table.rows:
        row_cells: List[str] = []
        for cell in row.cells:
            # Join all paragraph texts in the cell with line breaks
            texts = [
                (p.text or "").strip()
                for p in cell.paragraphs
                if (p.text or "").strip()
            ]
            row_cells.append("\n".join(texts))
        rows_data.append(row_cells)

    header: List[str] = rows_data[0] if rows_data else []
    body_rows: List[List[str]] = rows_data[1:] if len(rows_data) > 1 else []

    return {
        "kind": "table",
        "header": header,
        "rows": body_rows,
    }


def _iter_block_items(doc: Document):
    """
    Yield top-level block items (Paragraph or Table) from the document body
    in order.

    This is the standard python-docx pattern for iterating mixed content.
    """
    body = doc.element.body
    for child in body.iterchildren():
        if child.tag == qn("w:p"):
            yield Paragraph(child, doc)
        elif child.tag == qn("w:tbl"):
            yield Table(child, doc)


def _get_heading_level(paragraph: Paragraph) -> Optional[int]:
    """
    Return a heading level (1–5) if the paragraph uses a Heading style,
    otherwise None.

    Recognises both built-in 'Heading N' and custom 'ISMS Heading N'.
    """
    style = getattr(paragraph, "style", None)
    name = getattr(style, "name", "") or ""
    name_lower = name.lower()

    # Accept any style whose name contains "heading <n>", e.g.:
    #   "Heading 1"
    #   "ISMS Heading 2"
    #   "My Custom Heading 3"
    for lvl in range(1, 6):
        token = f"heading {lvl}"
        if token in name_lower:
            return lvl

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
    ('record_content'), but:

      - Use Word Heading styles (Heading N / ISMS Heading N) to create
        nested subsections inside it.
      - Preserve bullet / numbered lists via _paragraph_to_block().
      - Preserve tables via _table_to_block().
    """
    root_section: Dict[str, Any] = {
        "key": "record_content",
        "title": title,
        "level": 1,
        "content": [],
        "subsections": [],
    }

    # Stack of (heading_level, section_dict)
    section_stack: List[tuple[int, Dict[str, Any]]] = [(1, root_section)]

    def current_section() -> Dict[str, Any]:
        return section_stack[-1][1]

    for item in _iter_block_items(doc):
        if isinstance(item, Paragraph):
            heading_level = _get_heading_level(item)
            if heading_level is not None:
                # This paragraph is a heading: create a new section.
                heading_text = item.text.strip() or f"Heading {heading_level}"
                key = _slugify(heading_text)

                new_section: Dict[str, Any] = {
                    "key": key,
                    "title": heading_text,
                    "level": heading_level + 1,  # nested under record_content
                    "content": [],
                    "subsections": [],
                }

                # Attach to the appropriate parent based on heading level.
                while section_stack and section_stack[-1][0] >= heading_level + 1:
                    section_stack.pop()
                parent_section = section_stack[-1][1]
                parent_section["subsections"].append(new_section)

                section_stack.append((heading_level + 1, new_section))
                continue

            # Normal paragraph: convert to a content block
            block = _paragraph_to_block(item)
            if block is not None:
                current_section()["content"].append(block)

        elif isinstance(item, Table):
            table_block = _table_to_block(item)
            current_section()["content"].append(table_block)

    return root_section


# ---------------------------------------------------------------------------
# High-level import function
# ---------------------------------------------------------------------------


def import_word_to_json(
    input_path: Path,
    doc_type: str,
    doc_id: str,
    title: Optional[str] = None,
) -> Dict[str, Any]:
    """
    Import a Word document and return the ISMS JSON structure.

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
    mandatory_sections: List[tuple[str, str]] = [
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

    sections: List[Dict[str, Any]] = []

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
        help="Document ID for metadata.doc_id",
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

    print("[OK] Imported Word document:")
    print(f"     Source: {input_path}")
    print(f"     Output: {output_path}")


if __name__ == "__main__":
    main()
