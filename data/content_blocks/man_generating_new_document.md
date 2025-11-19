1. Duplicate a similar JSON in data/sample/ and rename with the next DocID.
2. Update metadata (doc_id, title, version, document_type).
3. Edit sections: use inline text or {"use_block": "<name>"} and add dynamic tables if needed.
4. Run python -m src.main.
5. Open the DOCX and Ctrl + A → F9 to refresh fields (Document Control, TOC).