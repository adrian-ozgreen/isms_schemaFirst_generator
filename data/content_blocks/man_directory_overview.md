├─ src/
│  ├─ isms_core/                    ← pipeline, renderers, schema, tables, props
│  └─ main.py                       ← entry point (builds all *.json in data/sample/)
├─ data/
│  ├─ config/document_profiles.yaml ← required/optional sections map
│  ├─ templates/ISMS_Master_Base.docx
│  ├─ content_blocks/               ← reusable boilerplates (.md)
│  └─ sample/                       ← one JSON per artefact
└─ outputs/isms_docs/               ← generated .docx
