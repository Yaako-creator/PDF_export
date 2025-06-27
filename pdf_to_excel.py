# Export text from PDFs in this directory into an Excel file.
#
# Dependencies (install locally or via pip):
#     pip install pypdf openpyxl

import glob
import os
from pypdf import PdfReader
from openpyxl import Workbook


def parse_pdfs(pdf_dir: str):
    """Extract text from all PDF files in the given directory."""
    records = []
    for pdf_path in glob.glob(os.path.join(pdf_dir, "*.pdf")):
        reader = PdfReader(pdf_path)
        pages = [page.extract_text() or "" for page in reader.pages]
        text = "\n".join(pages)
        records.append({"filename": os.path.basename(pdf_path), "text": text})
    return records


def export_to_excel(records, excel_path: str):
    """Write extracted PDF text to an Excel file."""
    wb = Workbook()
    ws = wb.active
    ws.append(["filename", "text"])
    for rec in sorted(records, key=lambda x: x["filename"]):
        ws.append([rec["filename"], rec["text"]])
    wb.save(excel_path)


def main():
    pdf_dir = os.path.dirname(__file__)
    excel_out = os.path.join(pdf_dir, "pdf_text.xlsx")
    data = parse_pdfs(pdf_dir)
    export_to_excel(data, excel_out)
    print(f"Exported {len(data)} PDFs to {excel_out}")


if __name__ == "__main__":
    main()
