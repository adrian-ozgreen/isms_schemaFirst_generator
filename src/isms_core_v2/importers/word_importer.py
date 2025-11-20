"""
Word → JSON importer for ISMS schema-first documents.

v0 profile: Record + PMDS-style specification
- Uses Word heading styles to build Section hierarchy.
- Uses list/body styles to map paragraphs to ContentBlocks.
"""

from __future__ import annotations

from pathlib import Path
from typing import List, Optional, Set

from docx import Document as DocxDocument
from docx.text.paragraph import Paragraph as DocxParagraph
from docx.table import Table as DocxTable
from docx.oxml.ns import qn
from docx.oxml.table import CT_Tbl
from docx.oxml.text.paragraph import CT_P

# If this import fails, change to: from isms_core_v2.models import Section, ContentBlock
from src.isms_core_v2.models import Section, ContentBlock





# ---------------------------------------------------------------------------
# Profile constants: Record + PMDS
# ---------------------------------------------------------------------------

# Map style names to heading levels
HEADING1_STYLES = ("ISMS Heading 1", "Heading 1")
HEADING2_STYLES = ("ISMS Heading 2", "Heading 2")
HEADING3_STYLES = ("ISMS Heading 3", "Heading 3")

# Map style names to body/list content types
BODY_STYLES = ("ISMS Body", "Normal")
BULLET_STYLES = ("ISMS List Bullet", "List Bullet")
NUMBERED_STYLES = ("ISMS List Numbered", "List Number")

# Mandatory ISMS sections expected by DocumentModel
MANDATORY_SECTION_TEMPLATES: dict[str, str] = {
    "title_page": "Title Page",
    "document_control": "Document Control",
    "table_of_contents": "Table of Contents",
    "revision_history": "Revision History",
    "approval_signatures": "Approval Signatures",
    "document_classification": "Document Classification",
    "purpose": "Purpose",
    "scope": "Scope",
    "roles_and_responsibilities": "Roles and Responsibilities",
    "related_documents": "Related Documents",
}

DOC_CLASS_CHILDREN_TEMPLATES: dict[str, str] = {
    "distribution_list": "Distribution List",
    "handling_requirements": "Handling Requirements",
    "retention_period": "Retention Period",
}

# ---------------------------------------------------------------------------
# Public entry point
# ---------------------------------------------------------------------------

def import_word_to_document_dict(
    path: Path,
    doc_type: str = "Record",
    default_doc_id: str = "REC-UNKNOWN-000",
) -> dict:
    """
    Load a Word document and produce a dict shaped like DocumentModel JSON:

        { "metadata": {...}, "sections": [ ... ] }

    v0:
    - Focused on Record-type PMDS-style docs.
    - Metadata is mostly heuristic / placeholder.
    - Sections + content derived from headings + paragraphs.
    """
    docx = DocxDocument(path)

    title = _guess_title(docx)

    metadata = {
        "doc_id": default_doc_id,
        "title": title or "Imported ISMS Record",
        "doc_type": doc_type,
        "version": "0.1",
        "status": "Draft",
        "owner": "TBD",
        "approver": None,
        "related_documents": [],
        "confidentiality": "Internal – TracWater Only",
        "date_completed": "",
        "next_review_date": "",
    }

    sections = _build_sections_from_doc(docx)
    sections = _ensure_mandatory_sections(sections)
    _ensure_document_classification_subsections(sections)


    return {
        "metadata": metadata,
        "sections": [s.model_dump() for s in sections],
    }


# ---------------------------------------------------------------------------
# Title / heading helpers
# ---------------------------------------------------------------------------

def _guess_title(docx: DocxDocument) -> Optional[str]:
    """
    Heuristic: use the first Heading 1 or Title paragraph as the document title.
    Falls back to the first non-empty paragraph.
    """
    # Try first Heading 1 / ISMS Heading 1 / Title
    for p in docx.paragraphs:
        style_name = (p.style.name or "").strip() if p.style else ""
        if style_name in HEADING1_STYLES or style_name.lower() == "title":
            text = p.text.strip()
            if text:
                return text

    # Fallback: first non-empty paragraph
    for p in docx.paragraphs:
        text = p.text.strip()
        if text:
            return text

    return None


def _get_heading_level(style_name: str) -> Optional[int]:
    """Map style name → section level (1–3) for this profile."""
    name = style_name.strip()
    if name in HEADING1_STYLES:
        return 1
    if name in HEADING2_STYLES:
        return 2
    if name in HEADING3_STYLES:
        return 3
    return None


def _get_list_kind_from_numbering(
    p: DocxParagraph, docx: DocxDocument
) -> Optional[str]:
    """
    Use Word's numbering (numPr) to determine if this paragraph is part of a list.

    Returns:
        "bullet_list"      if it looks like a bulleted list
        "numbered_list"    if it looks like a numbered list
        None               if it's not part of a list

    This works even when the paragraph style is "Normal", because it inspects
    the numbering.xml definitions without using XPath namespaces (for
    compatibility with python-docx' oxml wrapper).
    """
    # Paragraph properties
    pPr = p._p.pPr
    if pPr is None or pPr.numPr is None:
        return None

    numPr = pPr.numPr
    numId_elm = getattr(numPr, "numId", None)
    if numId_elm is None or numId_elm.val is None:
        return None

    num_id = str(numId_elm.val)

    # Underlying numbering.xml root element
    try:
        numbering_root = docx.part.numbering_part.element
    except AttributeError:
        # Document has no numbering part
        return None

    # 1) Find the <w:num> element with this numId and get its abstractNumId
    w_numId_attr = qn("w:numId")
    w_val_attr = qn("w:val")
    w_abstractNumId_attr = qn("w:abstractNumId")

    abstract_num_id = None

    for num in numbering_root.iter():
        # tag endswith 'num' catches '{...}num'
        if num.tag.endswith("num") and num.get(w_numId_attr) == num_id:
            for child in num.iter():
                if child.tag.endswith("abstractNumId"):
                    abstract_num_id = child.get(w_val_attr)
                    break
            if abstract_num_id:
                break

    if abstract_num_id is None:
        return None

    # 2) Find the <w:abstractNum> with this abstractNumId
    abstract_num = None
    for candidate in numbering_root.iter():
        if candidate.tag.endswith("abstractNum") and candidate.get(w_abstractNumId_attr) == abstract_num_id:
            abstract_num = candidate
            break

    if abstract_num is None:
        return None

    # 3) Within that abstractNum, find <w:numFmt w:val="...">
    num_format = None
    for child in abstract_num.iter():
        if child.tag.endswith("numFmt"):
            num_format = child.get(w_val_attr)
            break

    if num_format is None:
        return None

    if num_format == "bullet":
        return "bullet_list"

    # For our purposes, treat everything else as numbered
    return "numbered_list"




def _slugify_key(text: str) -> str:
    """Turn a heading text into a simple section key."""
    import re

    slug = re.sub(r"[^a-zA-Z0-9]+", "_", text.strip().lower()).strip("_")
    return slug or "section"



def _iter_block_items(docx: DocxDocument):
    """
    Yield block-level items (paragraphs and tables) from the document body
    in the order they appear.
    """
    body = docx.element.body
    for child in body.iterchildren():
        if isinstance(child, CT_P):
            yield DocxParagraph(child, docx)
        elif isinstance(child, CT_Tbl):
            yield DocxTable(child, docx)





# ---------------------------------------------------------------------------
# Section tree builder
# ---------------------------------------------------------------------------

# ---------------------------------------------------------------------------
# Section tree builder
# ---------------------------------------------------------------------------

def _build_sections_from_doc(docx: DocxDocument) -> List[Section]:
    """
    v0 algorithm (Record + PMDS), table-aware:

    - Walk through all block items in order (paragraphs + tables).
    - If a paragraph is a heading → start a new Section at that level.
    - Otherwise → attach content (paragraph / list / table) to the current section.
    """
    root_sections: List[Section] = []
    stack: List[Section] = []

    for block in _iter_block_items(docx):
        # -------------------------------------------------------------------
        # Paragraph handling (unchanged logic, just using block instead of p)
        # -------------------------------------------------------------------
        if isinstance(block, DocxParagraph):
            p = block
            style_name = (p.style.name or "").strip() if p.style else ""
            level = _get_heading_level(style_name)

            if level is not None:
                # New section
                new_section = Section(
                    key=_slugify_key(p.text or ""),
                    title=(p.text or "").strip(),
                    level=level,
                    content=[],
                    subsections=[],
                )
                _attach_section(root_sections, stack, new_section, level)
                continue

            text = (p.text or "").strip()
            if not text:
                # Empty paragraph → skip
                continue

            # Ensure we have a current section; if not, create a synthetic "Body"
            if not stack:
                if not root_sections:
                    root = Section(
                        key="body",
                        title="Body",
                        level=1,
                        content=[],
                        subsections=[],
                    )
                    root_sections.append(root)
                    stack.append(root)
            current = stack[-1]

            blocks = _paragraph_to_blocks(p, docx)
            current.content.extend(blocks)
            continue

        # -------------------------------------------------------------------
        # Table handling: attach as ContentBlock(kind="table")
        # -------------------------------------------------------------------
        if isinstance(block, DocxTable):
            tbl = block
            table_block = _table_to_block(tbl)
            if table_block is None:
                # Completely empty table → ignore
                continue

            # Ensure we have a current section; if not, create a synthetic "Body"
            if not stack:
                if not root_sections:
                    root = Section(
                        key="body",
                        title="Body",
                        level=1,
                        content=[],
                        subsections=[],
                    )
                    root_sections.append(root)
                    stack.append(root)
            current = stack[-1]
            current.content.append(table_block)
            continue

    return root_sections



def _attach_section(
    roots: List[Section],
    stack: List[Section],
    new_section: Section,
    level: int,
) -> None:
    """
    Attach new_section to the correct parent based on heading level.
    """
    while stack and stack[-1].level >= level:
        stack.pop()

    if not stack:
        roots.append(new_section)
    else:
        stack[-1].subsections.append(new_section)

    stack.append(new_section)


# ---------------------------------------------------------------------------
# Paragraph → ContentBlock mapping for Record + PMDS
# ---------------------------------------------------------------------------

def _paragraph_to_blocks(p: DocxParagraph, docx: DocxDocument) -> List[ContentBlock]:
    """
    Map a Word paragraph to one or more ContentBlocks.

    v0 rules (Record + PMDS):

    1. If numbering (numPr) says it's part of a list:
         - use that to classify as bullet_list or numbered_list.
    2. Else if style is in BULLET_STYLES / NUMBERED_STYLES:
         - classify accordingly.
    3. Else:
         - plain paragraph.
    """
    text = (p.text or "").strip()
    if not text:
        return []

    style_name = (p.style.name or "").strip() if p.style else ""

    # 1) Prefer Word numbering info if present
    list_kind = _get_list_kind_from_numbering(p, docx)
    if list_kind == "bullet_list":
        return [ContentBlock(kind="bullet_list", text=[text])]
    if list_kind == "numbered_list":
        return [ContentBlock(kind="numbered_list", text=[text])]

    # 2) Fall back to style-based detection (for docs that use explicit list styles)
    if style_name in BULLET_STYLES:
        return [ContentBlock(kind="bullet_list", text=[text])]

    if style_name in NUMBERED_STYLES:
        return [ContentBlock(kind="numbered_list", text=[text])]

    # 3) Default: body paragraph
    return [ContentBlock(kind="paragraph", text=text)]



def _table_to_block(tbl: DocxTable) -> Optional[ContentBlock]:
    """
    Convert a python-docx Table into a ContentBlock(kind='table').

    - First non-empty row is treated as the header.
    - Remaining non-empty rows become body rows.
    - Trailing empty cells/rows are trimmed.
    """
    raw_rows: List[List[str]] = []

    for row in tbl.rows:
        cells: List[str] = []
        for cell in row.cells:
            text = (cell.text or "").strip()
            cells.append(text)
        # Trim trailing empty cells
        while cells and not cells[-1]:
            cells.pop()
        raw_rows.append(cells)

    # Drop completely empty rows
    rows = [r for r in raw_rows if any(r)]
    if not rows:
        return None

    header = rows[0]
    body_rows = rows[1:] if len(rows) > 1 else []

    return ContentBlock(
        kind="table",
        header=header,
        rows=body_rows,
        # caption can be added later if we decide to parse it
    )




# ---------------------------------------------------------------------------
# Mandatory section helpers
# ---------------------------------------------------------------------------

def _collect_top_level_keys(sections: List[Section]) -> Set[str]:
    """
    Collect keys of top-level sections only.
    """
    return {s.key for s in sections}


def _ensure_mandatory_sections(sections: List[Section]) -> List[Section]:
    """
    Ensure that all mandatory ISMS sections exist as TOP-LEVEL sections,
    and order them so they appear first (in MANDATORY_SECTION_TEMPLATES order),
    followed by any other existing sections.

    We do NOT move/alter subsections or their content.
    """
    top_level_keys = _collect_top_level_keys(sections)

    # Index existing top-level sections by key and strip them out
    existing_by_key: dict[str, Section] = {}
    remaining: List[Section] = []
    for sec in sections:
        if sec.key in MANDATORY_SECTION_TEMPLATES and sec.key not in existing_by_key:
            existing_by_key[sec.key] = sec
        else:
            remaining.append(sec)

    ordered: List[Section] = []

    for key, title in MANDATORY_SECTION_TEMPLATES.items():
        if key in existing_by_key:
            # Use the existing section as-is
            ordered.append(existing_by_key[key])
            continue

        # Create a stub for missing mandatory sections
        content: List[ContentBlock] = []
        if key in ("purpose", "scope"):
            placeholder_text = (
                f"(Imported from existing business document. {title} has not been "
                "explicitly captured as a top-level section and should be reviewed "
                "and completed.)"
            )
            content.append(ContentBlock(kind="paragraph", text=placeholder_text))

        ordered.append(
            Section(
                key=key,
                title=title,
                level=1,
                content=content,
                subsections=[],
            )
        )

    # Then add all non-mandatory sections in their original order
    ordered.extend(remaining)
    return ordered


def _ensure_document_classification_subsections(sections: List[Section]) -> None:
    """
    Ensure that each 'document_classification' section has the mandatory
    child subsections: distribution_list, handling_requirements, retention_period.

    - We do NOT move or modify any existing subsections.
    - If a child key is missing, we add a stub subsection with placeholder content.
    """
    for sec in sections:
        # Recurse into tree
        _ensure_document_classification_subsections(sec.subsections)

        if sec.key != "document_classification":
            continue

        existing_child_keys = {child.key for child in sec.subsections}

        for child_key, child_title in DOC_CLASS_CHILDREN_TEMPLATES.items():
            if child_key in existing_child_keys:
                continue

            placeholder_text = (
                f"(Imported from existing business document. {child_title} has not been "
                "explicitly captured and should be reviewed and completed.)"
            )

            sec.subsections.append(
                Section(
                    key=child_key,
                    title=child_title,
                    level=sec.level + 1,
                    content=[ContentBlock(kind="paragraph", text=placeholder_text)],
                    subsections=[],
                )
            )