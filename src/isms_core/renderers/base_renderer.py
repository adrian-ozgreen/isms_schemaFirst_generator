from ..word_utils import add_body_under_heading, add_bullet_list_under_heading, populate_revision_history

class BaseRenderer:
    def __init__(self, doc, meta, sections, history):
        self.doc = doc
        self.meta = meta
        self.sections = sections or {}
        self.history = history or []
        self._rendered_section_keys = set()

    def add_body(self, heading, text):
        add_body_under_heading(self.doc, heading, text)

    def add_list(self, heading, items):
        add_bullet_list_under_heading(self.doc, heading, items)

    def render_common_sections(self):
        if "purpose" in self.sections:
            self.add_body("Purpose", self.sections["purpose"])
            self._rendered_section_keys.add("purpose")

        if "scope" in self.sections:
            self.add_body("Scope", self.sections["scope"])
            self._rendered_section_keys.add("scope")

        if "roles_and_responsibilities" in self.sections:
            lines = []
            for r in self.sections.get("roles_and_responsibilities", []):
                for resp in r.get("responsibilities", []):
                    lines.append(f"{r.get('role','Role')} – {resp}")
            if lines:
                self.add_list("Roles and Responsibilities", lines)
                self._rendered_section_keys.add("roles_and_responsibilities")

        if "related_documents" in self.sections:
            rel = self.sections.get("related_documents")
            items = []
            if isinstance(rel, list):
                if all(isinstance(d, dict) for d in rel):
                    # dict form: id/title/type
                    items = [
                        f"{d.get('id','ID')} – {d.get('title','Title')} ({d.get('type','Doc')})"
                        for d in rel
                    ]
                else:
                    # fallback: accept plain strings
                    items = [str(d) for d in rel]
            elif isinstance(rel, str):
                items = [rel]
            if items:
                self.add_list("Related Documents", items)
                self._rendered_section_keys.add("related_documents")

        if "definitions_and_acronyms" in self.sections:
            defs = [
                f"{d.get('term','')} – {d.get('definition','')}"
                for d in self.sections.get("definitions_and_acronyms", [])
            ]
            if defs:
                self.add_list("Definitions and Acronyms", defs)
                self._rendered_section_keys.add("definitions_and_acronyms")


    def render_remaining_sections(self):
        """
        Generic fallback: render any section keys that were not explicitly
        handled in render_common_sections() as headings + body or bullets.
        """
        rendered = getattr(self, "_rendered_section_keys", set())

        for key, val in self.sections.items():
            if key in rendered:
                continue
            if val is None or val == "" or val == []:
                continue

            heading = key.replace("_", " ").title()

            # String -> heading + paragraph
            if isinstance(val, str):
                self.add_body(heading, val)

            # List -> heading + bullet list
            elif isinstance(val, list):
                items = []
                for item in val:
                    if isinstance(item, str):
                        items.append(item)
                    elif isinstance(item, dict):
                        # flatten dict into a readable line
                        items.append("; ".join(f"{k}: {v}" for k, v in item.items()))
                    else:
                        items.append(str(item))
                if items:
                    self.add_list(heading, items)

            # Dict -> heading + bullet list of key: value
            elif isinstance(val, dict):
                items = [f"{k}: {v}" for k, v in val.items()]
                self.add_list(heading, items)

            # Any other type -> heading + stringified body
            else:
                self.add_body(heading, str(val))



from .base_renderer import BaseRenderer
from ..word_utils import (
    add_body_under_heading,
    add_bullet_list_under_heading,
    add_numbered_list_under_heading,
)

class RecordRenderer(BaseRenderer):
    def render(self):
        # 1) Purpose and Scope – attach content blocks to existing headings
        # if "purpose" in self.sections:
        #     add_body_under_heading(self.doc, "Purpose", self.sections["purpose"])
        #     self._rendered_section_keys.add("purpose")

        # if "scope" in self.sections:
        #     add_body_under_heading(self.doc, "Scope", self.sections["scope"])
        #     self._rendered_section_keys.add("scope")

        # 2) Record Content – handled below (see section 4)

        # 3) Other record-common sections
        super().render_common_sections()

        # 4) Manual-specific sections rendered explicitly (next points)
        ...
        # 5) Finally, render any remaining sections generically
        self.render_remaining_sections()

