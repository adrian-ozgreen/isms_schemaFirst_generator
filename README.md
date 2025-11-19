# ISMS Hybrid Artefact Generator (Skeleton)

This is a first-step scaffold for the hybrid architecture we agreed on:
- One common engine (`isms_core.pipeline`), with type-specific renderers (`Policy`, `Procedure`, `Record`).
- A single visual master template (`data/templates/ISMS_Master_Base.docx`) ensures consistent styles.
- JSON inputs per artefact, plus YAML profiles to keep section ordering/requirements per type.

## Quick start

1. Create a venv and install requirements:

```bash
python -m venv .venv
source .venv/bin/activate  # on Windows: .venv\Scripts\activate
pip install -r requirements.txt
```

2. Generate the sample document:

```bash
python -m src.main
```

The output DOCX will be written to `outputs/isms_docs/`.

## Notes

- This skeleton *does not* yet update DOCX Custom Properties (DocID, Version, etc.). We will add a small XML patcher later.
- `word_utils.py` is intentionally minimal and focuses on inserting content under known headings (e.g., Purpose, Scope).
- The renderer map lives in `src/isms_core/renderers/__init__.py`. Add new types (e.g., `Standard`, `Guideline`) easily.


## Step 2: Profiles-driven ordering & validation
- The pipeline now loads `data/config/document_profiles.yaml`, validates required/optional sections, and prints any issues to the console.
- The `Procedure` renderer follows the profile order when rendering sections.
- The Document Control table detection is more robust (finds table after the "Document Control" heading).


## Step 3: Front-matter tables
- New module `src/isms_core/front_matter.py` populates:
  - Approval Signatures (Name, Role, Signature, Date)
  - Distribution List (Recipient, Role/Dept, Method, Notes)
  - Document Classification (KV table)
  - Handling Requirements (KV table)
  - Retention Period (KV table)
- The pipeline fills these **before** rendering body content.
- The sample JSON includes example data for each table.


## Dynamic tables (JSON-driven, optional headings)

```json
"dynamic_tables": [
  {
    "target": {"after_heading": "Equipment Register"},
    "columns": ["Item", "Serial", "Installed At", "Notes"],
    "rows": [
      ["PS3 Meter", "SN-123", "Main Switchboard", ""],
      {"Item": "H2S Sensor", "Installed At": "Blower Room", "Notes": "Baseline offset applied"}
    ],
    "create_if_missing": true,
    "heading": null,
    "heading_style": "Heading 2"
  }
]
