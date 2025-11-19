# src/isms_core_v2/models.py

from __future__ import annotations
from typing import List, Literal, ClassVar, Set
from pydantic import BaseModel, Field, validator


# ------------------------------------------------------------
# Content blocks: paragraphs, bullet lists, numbered lists
# ------------------------------------------------------------

class ContentBlock(BaseModel):
    kind: Literal["paragraph", "bullet_list", "numbered_list", "table"]

    # For paragraph / bullet / numbered list
    text: str | List[str] | None = None

    # For table
    header: List[str] | None = None        # Optional header row
    rows: List[List[str]] | None = None    # Body rows
    caption: str | None = None             # Optional caption

    @validator("text", always=True)
    def validate_text_for_non_table(cls, v, values):
        kind = values.get("kind")
        if kind in ("paragraph", "bullet_list", "numbered_list"):
            if v is None:
                raise ValueError(f"'text' is required for kind={kind}")
        return v

    @validator("rows", always=True)
    def validate_rows_for_table(cls, v, values):
        kind = values.get("kind")
        if kind == "table":
            if v is None or len(v) == 0:
                raise ValueError("'rows' is required and must be non-empty for kind=table")
        return v


# ------------------------------------------------------------
# Section (recursive) with up to 3 levels of subsections
# level = 1 → ISMS Heading 1 style
# ------------------------------------------------------------

class Section(BaseModel):
    key: str
    title: str
    level: int = 1
    content: List[ContentBlock] = Field(default_factory=list)
    subsections: List["Section"] = Field(default_factory=list)

    @validator("level")
    def level_must_be_between_1_and_4(cls, v):
        if v < 1 or v > 4:
            raise ValueError("Section.level must be between 1 and 4")
        return v

    @validator("subsections")
    def enforce_subsection_levels(cls, subs, values):
        """Subsections must have level = parent.level + 1"""
        level = values.get("level", 1)
        for s in subs:
            if s.level != level + 1:
                raise ValueError(
                    f"Subsection '{s.title}' must have level {level+1}, got {s.level}"
                )
        return subs


# ------------------------------------------------------------
# Document metadata (feeds DCR and MRR automation)
# ------------------------------------------------------------

class DocMetadata(BaseModel):
    doc_id: str              # e.g. REC-OPS-001
    title: str
    doc_type: Literal["Record", "Policy", "Procedure"]
    version: str             # e.g. "1.0"
    status: Literal["Draft", "For Review", "Approved"]
    owner: str               # e.g. "CTO", "ISMS Manager"
    approver: str | None = None
    related_documents: List[str] = Field(default_factory=list)

    # New optional fields tied to placeholders
    confidentiality: str | None = None
    date_completed: str | None = None
    next_review_date: str | None = None

    @validator("doc_id")
    def doc_id_format(cls, v):
        # very light rule; can tighten later
        if len(v.strip()) == 0:
            raise ValueError("doc_id cannot be empty")
        return v


# ------------------------------------------------------------
# DocumentModel: root schema object
# ------------------------------------------------------------

class DocumentModel(BaseModel):
    metadata: DocMetadata
    sections: List[Section]

    MANDATORY_KEYS: ClassVar[Set[str]] = {
        "title_page",
        "document_control",
        "table_of_contents",
        "revision_history",
        "approval_signatures",
        "document_classification",
        "purpose",
        "scope",
        "roles_and_responsibilities",
        "related_documents"
    }

    CLASSIFICATION_SUBKEYS: ClassVar[Set[str]] = {
        "distribution_list",
        "handling_requirements",
        "retention_period"
    }

    @validator("sections")
    def validate_mandatory_sections(cls, sections):
        keys = {s.key for s in sections}

        # Check top-level mandatory sections exist
        missing = cls.MANDATORY_KEYS - keys
        if missing:
            raise ValueError(f"Missing mandatory sections: {', '.join(sorted(missing))}")

        # Validate the document_classification subsections
        classification = next((s for s in sections if s.key == "document_classification"), None)
        if classification is None:
            raise ValueError("document_classification section missing")

        subkeys = {s.key for s in classification.subsections}
        missing_sub = cls.CLASSIFICATION_SUBKEYS - subkeys
        if missing_sub:
            raise ValueError(
                f"document_classification missing subsections: {', '.join(sorted(missing_sub))}"
            )

        return sections
