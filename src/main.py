# src/main.py
from pathlib import Path
from .isms_core.pipeline import generate_isms_doc

BASE = Path(__file__).resolve().parents[1]
samples = (BASE / "data" / "sample").glob("*.json")
template = BASE / "data" / "templates" / "ISMS_Master_Base.docx"
outdir = BASE / "outputs" / "isms_docs"

for json_path in samples:
    out = generate_isms_doc(json_path, template, outdir)
    print("Generated ISMS document:", out)
