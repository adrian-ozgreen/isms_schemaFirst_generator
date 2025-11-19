from docx import Document

def scan(doc_path):
    doc = Document(doc_path)
    print("\n=== HEADING SCAN ===")
    for i, p in enumerate(doc.paragraphs):
        txt = p.text.strip()
        if not txt:
            continue
        if "Purpose" in txt or "Scope" in txt or "Roles" in txt:
            print(f"#{i} | '{txt}' | style={p.style.name}")

if __name__ == "__main__":
    scan("data/templates/ISMS_Master_Base.docx")
