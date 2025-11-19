from .base_renderer import BaseRenderer

class PolicyRenderer(BaseRenderer):
    def render(self):
        super().render_common_sections()
        if "policy_statements" in self.sections:
            for ps in self.sections.get("policy_statements", []):
                heading = ps.get("heading", "")
                text = ps.get("text", "")
                blob = f"{heading}\n\n{text}" if heading else text
                self.add_body("Policy Statements", blob)
        if "compliance_and_enforcement" in self.sections:
            self.add_body("Compliance and Enforcement", self.sections["compliance_and_enforcement"])
