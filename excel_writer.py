# excel_writer.py

from openpyxl import Workbook


def write_to_excel(hyperlinks, excel_path):
    wb = Workbook()
    ws = wb.active
    ws.title = "Hivatkozások"

    # Fejléc hozzáadása
    ws.append(["Szöveg", "Cél", "Típus"])

    # Adatok hozzáadása
    for link in hyperlinks:
        ws.append([link["text"], link["target"], link["type"]])

    # Excel fájl mentése
    wb.save(excel_path)
