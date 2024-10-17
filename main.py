# main.py

import os
import datetime
import logging
import tkinter as tk
from tkinter import filedialog
from docx import Document
from openpyxl import Workbook
from document_processor import process_document
from hyperlink_extractor import extract_hyperlinks
from excel_writer import write_to_excel
from file_utils import get_file_path, create_output_directory

logging.basicConfig(filename='main.log', level=logging.DEBUG,
                    format='%(asctime)s - %(levelname)s - %(message)s')


def select_file(title, filetypes):
    root = tk.Tk()
    root.withdraw()
    file_path = filedialog.askopenfilename(title=title, filetypes=filetypes)
    logging.info(f"Kiválasztott fájl: {file_path}")
    print(f"Kiválasztott fájl: {file_path}")
    return file_path


def select_directory(title, initial_dir=None):
    root = tk.Tk()
    root.withdraw()
    directory = filedialog.askdirectory(title=title, initialdir=initial_dir)
    logging.info(f"Kiválasztott könyvtár: {directory}")
    print(f"Kiválasztott könyvtár: {directory}")
    return directory


def main():
    logging.info("Program indítása")
    print("Program indítása")

    try:
        # Dokumentum bekérése
        doc_path = select_file("Válassza ki a Word dokumentumot", [("Word dokumentumok", "*.docx")])
        if not doc_path:
            logging.warning("Nem választott ki fájlt. A program leáll.")
            print("Nem választott ki fájlt. A program leáll.")
            return

        # Kimeneti mappa bekérése
        initial_dir = os.path.dirname(doc_path)
        output_dir = select_directory("Válassza ki a kimeneti mappát", initial_dir)
        if not output_dir:
            logging.warning("Nem választott ki kimeneti mappát. A program leáll.")
            print("Nem választott ki kimeneti mappát. A program leáll.")
            return

        create_output_directory(output_dir)

        # Időbélyeg generálása
        timestamp = datetime.datetime.now().strftime("%Y%m%d%H%M")

        # Excel fájl elérési útjának generálása
        excel_path = get_file_path(output_dir, f"hivatkozasok_{timestamp}.xlsx")

        # Log fájl elérési útjának generálása
        log_path = get_file_path(output_dir, f"log_{timestamp}.txt")

        logging.info("Dokumentum feldolgozásának kezdete")
        print("Dokumentum feldolgozásának kezdete")
        # Dokumentum feldolgozása
        document = process_document(doc_path)

        logging.info("Hivatkozások kinyerésének kezdete")
        print("Hivatkozások kinyerésének kezdete")
        # Hivatkozások kinyerése
        hyperlinks = extract_hyperlinks(document)

        logging.info("Excel fájl írásának kezdete")
        print("Excel fájl írásának kezdete")
        # Excel fájl létrehozása és írása
        write_to_excel(hyperlinks, excel_path)

        print(f"A hivatkozások sikeresen kimentve: {excel_path}")
        logging.info(f"A hivatkozások sikeresen kimentve: {excel_path}")

        # Log írása
        with open(log_path, "w", encoding="utf-8") as log_file:
            log_file.write(f"Feldolgozás sikeres. Időpont: {timestamp}\n")
            log_file.write(f"Bemeneti fájl: {doc_path}\n")
            log_file.write(f"Kimeneti fájl: {excel_path}\n")
            log_file.write(f"Talált hivatkozások száma: {len(hyperlinks)}\n")

        print(f"Log fájl létrehozva: {log_path}")
        logging.info(f"Log fájl létrehozva: {log_path}")

    except Exception as e:
        print(f"Hiba történt: {str(e)}")
        logging.error(f"Hiba történt: {str(e)}", exc_info=True)
        with open(log_path, "w", encoding="utf-8") as log_file:
            log_file.write(f"Hiba történt. Időpont: {timestamp}\n")
            log_file.write(f"Hibaüzenet: {str(e)}\n")


if __name__ == "__main__":
    main()
    print("Program befejezve. Nyomjon Enter-t a kilépéshez.")
    input()
