
from .base_renderer import BaseRenderer

class ProcedureRenderer(BaseRenderer):
    def render(self):
        # Profiles-driven order for Procedure
        prof = {
            "required_sections": [
                "Purpose",
                "Scope",
                "Procedure Steps",
                "Roles and Responsibilities"
            ],
            "optional_sections": [
                "Related Documents",
                "Definitions and Acronyms",
                "Normative References"
            ]
        }
        # Map canonical headings to JSON keys
        keymap = {
            "Purpose": "purpose",
            "Scope": "scope",
            "Procedure Steps": "procedure_steps",
            "Roles and Responsibilities": "roles_and_responsibilities",
            "Related Documents": "related_documents",
            "Definitions and Acronyms": "definitions_and_acronyms",
            "Normative References": "normative_references",
        }

        def render_section(h):
            k = keymap.get(h)
            if not k: return
            if k == "procedure_steps":
                steps = self.sections.get(k) or []
                if steps: self.add_list("Procedure Steps", steps)
                return
            if k == "roles_and_responsibilities":
                lines = []
                for r in self.sections.get(k, []):
                    for resp in r.get("responsibilities", []):
                        lines.append(f"{r.get('role','Role')} – {resp}")
                if lines: self.add_list("Roles and Responsibilities", lines)
                return
            # default paragraph
            val = self.sections.get(k)
            if val:
                if isinstance(val, str):
                    self.add_body(h, val)
                else:
                    # flatten other structures into bullets
                    if isinstance(val, list):
                        items = []
                        for item in val:
                            if isinstance(item, dict):
                                items.append(" – ".join([str(v) for v in item.values() if v]))
                            else:
                                items.append(str(item))
                        if items: self.add_list(h, items)

        order = prof["required_sections"] + [h for h in prof["optional_sections"] if self.sections.get(keymap.get(h))]
        for h in order:
            render_section(h)
