import re
from pathlib import Path

import win32com.client as win32


# --- SETTINGS ---------------------------------------------------------

# Path to your template workbook (from step 1)
TEMPLATE_PATH = Path(r"C:\Users\Administrator\Desktop\ExcelProcess\EXCEL\template_pdf_imports.xlsx")

# Folder that contains PDF files to convert
PDF_FOLDER = Path(r"C:\Users\Administrator\Desktop\ExcelProcess\PDF")

# Folder to save the resulting Excel files – here: same folder as this script
OUTPUT_FOLDER = Path(__file__).parent   # or Path(r"C:\Projects\PdfToExcel\output")

# Name of the query in the template (see Queries & Connections pane)
QUERY_NAME = "Query1" 

# ---------------------------------------------------------------------


def update_query_pdf_path(query, new_pdf_path: Path):
    """Replace the File.Contents(...) path in the Power Query M formula."""
    formula = query.Formula

    # Excel is happy with forward slashes, which avoids backslash-escaping issues
    pdf_str = new_pdf_path.resolve().as_posix()

    # Replace the part File.Contents("...") with the new path
    new_formula = re.sub(
        r'File\.Contents\(".*?"\)',
        f'File.Contents("{pdf_str}")',
        formula,
    )
    query.Formula = new_formula


def main():
    OUTPUT_FOLDER.mkdir(parents=True, exist_ok=True)

    excel = win32.Dispatch("Excel.Application")
    excel.Visible = False  # set True while debugging if you want to see Excel

    try:
        for pdf_path in PDF_FOLDER.glob("*.pdf"):
            print(f"Processing {pdf_path.name}...")

            # Open a fresh copy of the template each time
            wb = excel.Workbooks.Open(str(TEMPLATE_PATH))

            # Get the Power Query object
            query = wb.Queries(QUERY_NAME)

            # Point the query to the current PDF
            update_query_pdf_path(query, pdf_path)

            # Refresh all queries in this workbook
            wb.RefreshAll()

            # Wait until queries finish (Power Query is async)
            excel.CalculateUntilAsyncQueriesDone()

            # Build output path: same name as PDF but .xlsx
            out_path = OUTPUT_FOLDER / f"{pdf_path.stem}.xlsx"

            # 51 = xlOpenXMLWorkbook (.xlsx)
            wb.SaveAs(str(out_path), FileFormat=51)

            wb.Close(SaveChanges=False)

            print(f"Saved -> {out_path}")

    finally:
        excel.Quit()


if __name__ == "__main__":
    main()
