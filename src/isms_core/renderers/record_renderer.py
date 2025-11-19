from .base_renderer import BaseRenderer
from ..word_utils import (
    add_body_under_heading,
    add_numbered_list_under_heading,
    add_bullet_list_under_heading,
)


class RecordRenderer(BaseRenderer):
    def debug_scan(self):
        print("\n=== RUNTIME SCAN ===")
        for i, p in enumerate(self.doc.paragraphs):
            txt = p.text.strip()
            if txt:
                print(f"#{i}: '{txt}' | style={p.style.name}")


    def render(self):

        # TEMP DEBUG — inspect the REAL document in memory
        #self.debug_scan()

        
        # 1) Render common sections (Purpose, Scope, Roles, Related Docs, Definitions, etc.)
        super().render_common_sections()

        # 2) Render the Record-specific core section
        # Record Content -> bullets 
        if "record_content" in self.sections:
            rc = self.sections["record_content"]
            if isinstance(rc, list):
                add_bullet_list_under_heading(self.doc, "Record Content", rc)
            else:
                add_body_under_heading(self.doc, "Record Content", rc)
            self._rendered_section_keys.add("record_content")

        # Json Input Structure -> bullets
        if "json_input_structure" in self.sections:
            lines = [
                l.strip() for l in str(self.sections["json_input_structure"]).splitlines()
                if l.strip()
            ]
            add_bullet_list_under_heading(self.doc, "Json Input Structure", lines)
            self._rendered_section_keys.add("json_input_structure")

        # Governance -> bullets
        if "governance" in self.sections:
            lines = [
                l.strip() for l in str(self.sections["governance"]).splitlines()
                if l.strip()
            ]
            add_bullet_list_under_heading(self.doc, "Governance", lines)
            self._rendered_section_keys.add("governance")


        # Generating A New Document -> numbered list
        if "generating_a_new_document" in self.sections:
            steps = self.sections["generating_a_new_document"]
            if isinstance(steps, list):
                add_numbered_list_under_heading(self.doc, "Generating A New Document", steps)
            else:
                add_body_under_heading(self.doc, "Generating A New Document", steps)
            self._rendered_section_keys.add("generating_a_new_document")

        # Creating A New Version -> numbered list
        if "creating_a_new_version" in self.sections:
            steps = self.sections["creating_a_new_version"]
            if isinstance(steps, list):
                add_numbered_list_under_heading(self.doc, "Creating A New Version", steps)
            else:
                add_body_under_heading(self.doc, "Creating A New Version", steps)
            self._rendered_section_keys.add("creating_a_new_version")



        # 3) Generic fallback for any remaining sections
        #    (prerequisites, directory_overview, json_input_structure, etc.)
        self.render_remaining_sections()
