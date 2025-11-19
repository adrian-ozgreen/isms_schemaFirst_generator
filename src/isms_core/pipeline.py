from pathlib import Path
import json
from docx import Document
from .docx_props import set_doc_properties
from .doc_control import populate_document_control_table
from .front_matter import (
    populate_distribution_list,
    populate_approval_signatures,
    populate_document_classification,
    populate_handling_requirements,
    populate_retention_period,
)
from ..config_loader import load_profiles
from .renderers import load_renderer
from .dynamic_tables import populate_dynamic_table
from .content_blocks import load_blocks, merge_content_blocks
from .schema import validate_input_payload
from ..config_loader import load_profiles  # you already have this in v9+


def _validate_sections(doc_type: str, sections: dict, profiles: dict) -> list[tuple[str, str]]:
    """
    Return a list of validation messages (severity, message).
    """
    msgs = []
    prof = profiles.get(doc_type, {})
    required = [s.strip() for s in prof.get("required_sections", [])]
    optional = [s.strip() for s in prof.get("optional_sections", [])]

    # Map canonical headings to JSON keys we currently support
    keymap = {
        "Purpose": "purpose",
        "Scope": "scope",
        "Roles and Responsibilities": "roles_and_responsibilities",
        "Related Documents": "related_documents",
        "Definitions and Acronyms": "definitions_and_acronyms",
        "Policy Statements": "policy_statements",
        "Compliance and Enforcement": "compliance_and_enforcement",
        "Procedure Steps": "procedure_steps",
        "Record Content": "record_content",
    }

    present_headings = {h for h,k in keymap.items() if k in sections and sections.get(k)}
    for h in required:
        if h not in present_headings:
            msgs.append(("ERROR", f"Missing required section: {h}"))
    for h in present_headings:
        if h not in required and h not in optional:
            msgs.append(("WARN", f"Section present but not in profile for {doc_type}: {h}"))

    return msgs, keymap, required, optional

def generate_isms_doc(json_path: Path, template_path: Path, output_dir: Path) -> Path:
    data = json.loads(Path(json_path).read_text(encoding="utf-8"))
    # Step 6: Load and merge content blocks (Markdown boilerplates)
    blocks = load_blocks(Path(__file__).resolve().parents[2] / "data" / "content_blocks")
    data = merge_content_blocks(data, blocks)

    profiles = load_profiles(Path(__file__).resolve().parents[2] / 'data' / 'config' / 'document_profiles.yaml')






    # NEW: lightweight schema validation
    errors, warns = validate_input_payload(data, profiles, strict=True)
    for w in warns:
        print(f"[WARN] {w}")
    if errors:
        for e in errors:
            print(f"[ERROR] {e}")
        # fail fast to avoid half-baked docs; flip to False to allow graceful build
        raise SystemExit(1)







    meta = data.get("metadata", {})
    doc_type = meta.get("document_type", "Procedure")
    sections = data.get("sections", {})
    history = data.get("revision_history", [])

    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)
    output_name = f"{meta.get('doc_id','DOC')} {meta.get('title','Document')} v{meta.get('version','1.0')}.docx"
    output_path = output_dir / output_name

    # Start from the master base template
    output_path.write_bytes(Path(template_path).read_bytes())
    doc = Document(str(output_path))

    # Validation & ordering
    msgs, keymap, required, optional = _validate_sections(doc_type, sections, profiles)
    for sev, m in msgs:
        print(f"[{sev}] {m}")

    # Render body
    renderer_cls = load_renderer(doc_type)
    renderer = renderer_cls(doc, meta, sections, history)
    renderer.render()

    # Populate Document Control table first (so content appears near front)
    populate_document_control_table(doc, meta)

    # Generic dynamic tables (optional)
    for spec in data.get('dynamic_tables', []):
        try:
            populate_dynamic_table(doc, spec)
        except Exception as e:
            print('Dynamic table error:', e)


    doc.save(str(output_path))
    # After saving body, update core/custom properties so fields & headers/footers pick them up
    set_doc_properties(output_path, meta)
    return output_path
