# Dependencies:
# pip install pdfminer.six pandas openpyxl

import glob
import os
import pandas as pd
from pdfminer.high_level import extract_text


def parse_pdfs(pdf_dir: str):
    """Extract text from all PDF files in the given directory."""
    data = []
    for pdf_path in glob.glob(os.path.join(pdf_dir, "*.pdf")):
        text = extract_text(pdf_path)
        data.append({"filename": os.path.basename(pdf_path), "text": text})
    return data


def export_to_excel(records, excel_path: str):
    """Write extracted PDF text to an Excel file."""
    df = pd.DataFrame(records)
    df.sort_values("filename", inplace=True)
    df.to_excel(excel_path, index=False)


def main():
    pdf_dir = os.path.dirname(__file__)
    excel_out = os.path.join(pdf_dir, "pdf_text.xlsx")
    data = parse_pdfs(pdf_dir)
    export_to_excel(data, excel_out)
    print(f"Exported {len(data)} PDFs to {excel_out}")


if __name__ == "__main__":
    main()
