# document_processor.py

from docx import Document

def process_document(doc_path):
    try:
        return Document(doc_path)
    except Exception as e:
        raise Exception(f"Nem sikerült megnyitni a dokumentumot: {str(e)}")
