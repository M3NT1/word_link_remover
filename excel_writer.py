# excel_writer.py

from openpyxl import Workbook
import logging


def write_to_excel(hyperlinks, excel_path):
    logging.info(f"Excel fájl írása kezdődik: {excel_path}")
    wb = Workbook()
    ws = wb.active
    ws.title = "Hivatkozások"

    # Fejléc hozzáadása
    ws.append(["Szöveg", "Cél", "Típus", "Hivatkozás szövege", "Kontextus"])
    logging.debug("Excel fejléc hozzáadva")

    # Adatok hozzáadása
    for link in hyperlinks:
        ws.append([
            link["text"],
            link["target"],
            link["type"],
            link.get("link_text", ""),
            link["context"]
        ])
        logging.debug(f"Hivatkozás hozzáadva az Excelhez: {link['text']} -> {link['target']}")

    # Excel fájl mentése
    wb.save(excel_path)
    logging.info(f"Excel fájl sikeresen mentve: {excel_path}")
