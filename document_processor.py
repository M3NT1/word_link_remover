# document_processor.py

from docx import Document
import logging

def process_document(doc_path):
    logging.info(f"Dokumentum feldolgozása kezdődik: {doc_path}")
    try:
        document = Document(doc_path)
        logging.info(f"Dokumentum sikeresen betöltve: {doc_path}")
        return document
    except Exception as e:
        logging.error(f"Hiba a dokumentum megnyitása során: {str(e)}")
        raise Exception(f"Nem sikerült megnyitni a dokumentumot: {str(e)}")
