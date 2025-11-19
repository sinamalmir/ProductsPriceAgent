# from pathlib import Path
# import re

# import win32com.client as win32
# from openpyxl import load_workbook


# # ------------- SETTINGS (تنظیمات را اینجا مطابق سیستم خودت عوض کن) -------------

# # 1) فایل تمپلیت Power Query که از PDF وارد می‌کند (قبلی که ساختیم)
# # TEMPLATE_IMPORT_XLSX = Path(r"C:\Users\Administrator\Desktop\ExcelProcess\EXCEL\template_pdf_imports.xlsx")
# TEMPLATE_IMPORT_XLSX = Path(r"C:\Users\Administrator\Desktop\pdfConvertor\ProductsPriceAgent\EXCEL\template_pdf_imports.xlsx")

# # 2) پوشه‌ای که PDFها داخل آن هستند
# # PDF_FOLDER = Path(r"C:\Users\Administrator\Desktop\ExcelProcess\PDF")
# PDF_FOLDER = Path(r"C:\Users\Administrator\Desktop\pdfConvertor\ProductsPriceAgent\PDF")

# # 3) پوشه خروجی فایل‌های Excel تبدیل‌شده
# # OUTPUT_FOLDER = Path(r"C:\Users\Administrator\Desktop\ExcelProcess\converted_excels")
# OUTPUT_FOLDER = Path(r"C:\Users\Administrator\Desktop\pdfConvertor\ProductsPriceAgent\converted_excels")
# # 4) نام Query در Excel (از پنجره Queries & Connections)
# QUERY_NAME = "Query1"

# # 5) فایل تمپلیت نهایی که لیست مدل‌ها را دارد (همان عکسی که فرستادی)
# # FINAL_TEMPLATE_PATH = Path(r"C:\Users\Administrator\Desktop\ExcelProcess\excel_data\price_of_products.xlsx")
# FINAL_TEMPLATE_PATH = Path(r"C:\Users\Administrator\Desktop\pdfConvertor\ProductsPriceAgent\excel_data\price_of_products.xlsx")

# # 6) نام ستونی که می‌خواهی در آن قیمت‌ها نوشته شود
# TARGET_COLUMN_HEADER = "باطری روکاری"

# # ---------------------------------------------------------------------


# def update_query_pdf_path(query, new_pdf_path: Path):
#     """مسیر فایل PDF را در فرمول Power Query عوض می‌کند."""
#     formula = query.Formula
#     pdf_str = new_pdf_path.resolve().as_posix()

#     new_formula = re.sub(
#         r'File\.Contents\(".*?"\)',
#         f'File.Contents("{pdf_str}")',
#         formula,
#     )
#     query.Formula = new_formula


# # نرمال‌سازی اسم‌ها برای مقایسه بدون حساسیت به فاصله/حروف بزرگ
# def normalize_name(s: str) -> str:
#     return re.sub(r"\s+", " ", s.strip().lower())


# # نگاشت بین اسم‌هایی که در فایل دیتا می‌آید و اسم‌هایی که در تمپلیت داری
# MODEL_MAPPING = {
#     # کلیدها: دقیقاً همان متن سلول در فایل دیتا
#     # مقدار: لیست نام‌هایی که در ستون A تمپلیت وجود دارد
#     "8/SE 2020/2022": ["iPhone 8", "iPhone SE (1st Gen)", "iPhone SE (3rd Gen)"],
#     "XR": ["iPhone XR"],
#     "XS": ["iPhone XS"],
#     "XS MAX": ["iPhone XS Max"],
#     "11": ["iPhone 11"],
#     "11 PRO": ["iPhone 11 Pro"],
#     "11 PRO MAX": ["iPhone 11 Pro Max"],
#     "12/12 PRO": ["iPhone 12", "iPhone 12 Pro"],
#     "12 MINI": ["iPhone 12 mini"],
#     "12 PRO MAX": ["iPhone 12 Pro Max"],
#     "13": ["iPhone 13"],
#     "13 MINI": ["iPhone 13 mini"],
#     "13 PRO": ["iPhone 13 Pro"],
#     "13 PRO MAX": ["iPhone 13 Pro Max"],
#     "14": ["iPhone 14"],
#     "14 PLUS": ["iPhone 14 Plus"],
#     "14 PRO": ["iPhone 14 Pro"],
#     "14 PRO MAX": ["iPhone 14 Pro Max"],
#     # اگر مدل‌های دیگری هم داری، مشابه همین اینجا اضافه کن
# }


# def convert_pdfs_to_excels():
#     """همه PDFهای پوشه را با Excel به xlsx تبدیل می‌کند."""
#     OUTPUT_FOLDER.mkdir(parents=True, exist_ok=True)

#     excel = win32.Dispatch("Excel.Application")
#     excel.Visible = False

#     try:
#         for pdf_path in PDF_FOLDER.glob("*.pdf"):
#             print(f"Processing PDF -> {pdf_path.name}")

#             wb = excel.Workbooks.Open(str(TEMPLATE_IMPORT_XLSX))
#             query = wb.Queries(QUERY_NAME)

#             update_query_pdf_path(query, pdf_path)

#             wb.RefreshAll()
#             excel.CalculateUntilAsyncQueriesDone()

#             out_path = OUTPUT_FOLDER / f"{pdf_path.stem}.xlsx"
#             wb.SaveAs(str(out_path), FileFormat=51)
#             wb.Close(SaveChanges=False)

#             print(f"Saved Excel -> {out_path}")

#     finally:
#         excel.Quit()


# def fill_template_from_converted_excel(converted_xlsx: Path):
#     """
#     از فایل Excel تبدیل‌شده (مثل cell_high_capacity.xlsx)
#     مدل/قیمت را می‌خواند و در template.xlsx در ستون 'باطری روکاری' می‌نویسد.
#     """
#     print(f"Reading data from: {converted_xlsx}")

#     # 1) خواندن فایل دیتا
#     data_wb = load_workbook(converted_xlsx, data_only=True)
#     data_ws = data_wb.active

#     # 2) باز کردن تمپلیت نهایی
#     tmpl_wb = load_workbook(FINAL_TEMPLATE_PATH)
#     tmpl_ws = tmpl_wb.active

#     # 3) پیدا کردن شماره ستون "باطری روکاری"
#     target_col_idx = None
#     for cell in tmpl_ws[1]:
#         if str(cell.value).strip() == TARGET_COLUMN_HEADER:
#             target_col_idx = cell.column
#             break

#     if target_col_idx is None:
#         raise ValueError(f"ستونی با عنوان '{TARGET_COLUMN_HEADER}' در ردیف اول تمپلیت پیدا نشد.")

#     # 4) دیکشنری: اسم مدل (نرمال‌شده) → شماره ردیف در تمپلیت
#     template_model_rows = {}
#     for row in range(2, tmpl_ws.max_row + 1):
#         val = tmpl_ws.cell(row=row, column=1).value
#         if not val:
#             continue
#         template_model_rows[normalize_name(str(val))] = row

#     # 5) خواندن رنج C6:D... از فایل دیتا و نوشتن در تمپلیت
#     for model_val, price_val in data_ws.iter_rows(
#         min_row=6, min_col=3, max_col=4, values_only=True
#     ):
#         if not model_val or not price_val:
#             continue

#         if str(price_val).strip() == "-":
#             # قیمت ناموجود
#             continue

#         key = str(model_val).strip().upper()

#         # اول از نگاشت استفاده می‌کنیم
#         mapped_models = MODEL_MAPPING.get(key)

#         # اگر در دیکشنری نبود، سعی می‌کنیم مستقیم با "iPhone + مدل" مچ کنیم
#         if not mapped_models:
#             generic_name = "iPhone " + str(model_val).strip()
#             norm = normalize_name(generic_name)
#             if norm in template_model_rows:
#                 mapped_models = [generic_name]
#             else:
#                 print(f"مدل ناشناخته در دیتا: {model_val}  (رد شد)")
#                 continue

#         # الان برای هر مدل تمپلیت، قیمت را در ستون هدف می‌نویسیم
#         for tmpl_model_name in mapped_models:
#             row_idx = template_model_rows.get(normalize_name(tmpl_model_name))
#             if not row_idx:
#                 print(f"مدل '{tmpl_model_name}' در تمپلیت پیدا نشد (از روی {model_val})")
#                 continue

#             tmpl_ws.cell(row=row_idx, column=target_col_idx).value = price_val

#     # 6) ذخیره تمپلیت
#     tmpl_wb.save(FINAL_TEMPLATE_PATH)
#     print(f"Template updated with prices from {converted_xlsx.name}")


# def main():
#     convert_pdfs_to_excels()
#     # اگر می‌خواهی اول PDFها را هم خودکار تبدیل کنی، این خط را باز کن:
#     # convert_pdfs_to_excels()

#     # فرض: فایل تبدیل‌شده‌ای که گفتی نامش مثلاً این است:
#     converted_file = OUTPUT_FOLDER / "cell  HIGH CAPACITY.xlsx"
#     fill_template_from_converted_excel(converted_file)


# if __name__ == "__main__":
#     main()


from pathlib import Path
import re

import win32com.client as win32
from openpyxl import load_workbook


# ------------- SETTINGS (تنظیمات را اینجا مطابق سیستم خودت عوض کن) -------------

TEMPLATE_IMPORT_XLSX = Path(r"C:\Users\Administrator\Desktop\pdfConvertor\ProductsPriceAgent\EXCEL\template_pdf_imports.xlsx")
PDF_FOLDER           = Path(r"C:\Users\Administrator\Desktop\pdfConvertor\ProductsPriceAgent\PDF")
OUTPUT_FOLDER        = Path(r"C:\Users\Administrator\Desktop\pdfConvertor\ProductsPriceAgent\converted_excels")

# نام Query در Excel
QUERY_NAME = "Query1"

# فایل تمپلیت نهایی (جدول بزرگ مدل‌ها)
FINAL_TEMPLATE_PATH = Path(r"C:\Users\Administrator\Desktop\pdfConvertor\ProductsPriceAgent\excel_data\FinalTemplate.xlsx")

# ستونی که قیمت باتری روکاری در آن نوشته می‌شود
TARGET_COLUMN_HEADER = "باطری روکاری"

# --- جدید: فایل JC PRODUCTS و ستون «تگ باطری» ---
JC_PRODUCTS_PATH   = Path(r"C:\Users\Administrator\Desktop\pdfConvertor\ProductsPriceAgent\converted_excels\JC PRODUCTS NORMAL.xlsx")
TAG_COLUMN_HEADER  = "تگ باطری"

# ---------------------------------------------------------------------


def update_query_pdf_path(query, new_pdf_path: Path):
    """مسیر فایل PDF را در فرمول Power Query عوض می‌کند."""
    formula = query.Formula
    pdf_str = new_pdf_path.resolve().as_posix()

    new_formula = re.sub(
        r'File\.Contents\(".*?"\)',
        f'File.Contents("{pdf_str}")',
        formula,
    )
    query.Formula = new_formula


def normalize_name(s: str) -> str:
    """نرمال‌سازی اسم‌ها برای مقایسه بدون حساسیت به فاصله/حروف بزرگ"""
    return re.sub(r"\s+", " ", s.strip().lower())


# نگاشت بین اسم‌هایی که در فایل دیتا می‌آید و اسم‌هایی که در تمپلیت داری
MODEL_MAPPING = {
    "8/SE 2020/2022": ["iPhone 8", "iPhone SE (1st Gen)", "iPhone SE (3rd Gen)"],
    "XR": ["iPhone XR"],
    "XS": ["iPhone XS"],
    "XS MAX": ["iPhone XS Max"],
    "11": ["iPhone 11"],
    "11 PRO": ["iPhone 11 Pro"],
    "11 PRO MAX": ["iPhone 11 Pro Max"],
    "12/12 PRO": ["iPhone 12", "iPhone 12 Pro"],
    "12 MINI": ["iPhone 12 mini"],
    "12 PRO MAX": ["iPhone 12 Pro Max"],
    "13": ["iPhone 13"],
    "13 MINI": ["iPhone 13 mini"],
    "13 PRO": ["iPhone 13 Pro"],
    "13 PRO MAX": ["iPhone 13 Pro Max"],
    "14": ["iPhone 14"],
    "14 PLUS": ["iPhone 14 Plus"],
    "14 PRO": ["iPhone 14 Pro"],
    "14 PRO MAX": ["iPhone 14 Pro Max"],
}


def convert_pdfs_to_excels():
    """همه PDFهای پوشه را با Excel به xlsx تبدیل می‌کند."""
    OUTPUT_FOLDER.mkdir(parents=True, exist_ok=True)

    excel = win32.Dispatch("Excel.Application")
    excel.Visible = False

    try:
        for pdf_path in PDF_FOLDER.glob("*.pdf"):
            print(f"Processing PDF -> {pdf_path.name}")

            wb = excel.Workbooks.Open(str(TEMPLATE_IMPORT_XLSX))
            query = wb.Queries(QUERY_NAME)

            update_query_pdf_path(query, pdf_path)

            wb.RefreshAll()
            excel.CalculateUntilAsyncQueriesDone()

            out_path = OUTPUT_FOLDER / f"{pdf_path.stem}.xlsx"
            wb.SaveAs(str(out_path), FileFormat=51)
            wb.Close(SaveChanges=False)

            print(f"Saved Excel -> {out_path}")

    finally:
        excel.Quit()


def fill_template_from_converted_excel(converted_xlsx: Path):
    """
    از فایل Excel تبدیل‌شده (مثل cell  HIGH CAPACITY.xlsx)
    مدل/قیمت را می‌خواند و در price_of_products.xlsx در ستون 'باطری روکاری' می‌نویسد.
    """
    print(f"Reading data from: {converted_xlsx}")

    # 1) خواندن فایل دیتا
    data_wb = load_workbook(converted_xlsx, data_only=True)
    data_ws = data_wb.active

    # 2) باز کردن تمپلیت نهایی
    tmpl_wb = load_workbook(FINAL_TEMPLATE_PATH)
    tmpl_ws = tmpl_wb.active

    # 3) پیدا کردن شماره ستون "باطری روکاری"
    target_col_idx = None
    for cell in tmpl_ws[1]:
        if str(cell.value).strip() == TARGET_COLUMN_HEADER:
            target_col_idx = cell.column
            break

    if target_col_idx is None:
        raise ValueError(f"ستونی با عنوان '{TARGET_COLUMN_HEADER}' در ردیف اول تمپلیت پیدا نشد.")

    # 4) دیکشنری: اسم مدل (نرمال‌شده) → شماره ردیف در تمپلیت
    template_model_rows = {}
    for row in range(2, tmpl_ws.max_row + 1):
        val = tmpl_ws.cell(row=row, column=1).value
        if not val:
            continue
        template_model_rows[normalize_name(str(val))] = row

    # 5) خواندن رنج C6:D... از فایل دیتا و نوشتن در تمپلیت
    for model_val, price_val in data_ws.iter_rows(
        min_row=6, min_col=3, max_col=4, values_only=True
    ):
        if not model_val or not price_val:
            continue

        if str(price_val).strip() == "-":
            # قیمت ناموجود
            continue

        key = str(model_val).strip().upper()

        mapped_models = MODEL_MAPPING.get(key)

        # اگر در دیکشنری نبود، سعی می‌کنیم مستقیم با "iPhone + مدل" مچ کنیم
        if not mapped_models:
            generic_name = "iPhone " + str(model_val).strip()
            norm = normalize_name(generic_name)
            if norm in template_model_rows:
                mapped_models = [generic_name]
            else:
                print(f"مدل ناشناخته در دیتا (فایل قیمت): {model_val}  (رد شد)")
                continue

        # الان برای هر مدل تمپلیت، قیمت را در ستون هدف می‌نویسیم
        for tmpl_model_name in mapped_models:
            row_idx = template_model_rows.get(normalize_name(tmpl_model_name))
            if not row_idx:
                print(f"مدل '{tmpl_model_name}' در تمپلیت پیدا نشد (از روی {model_val})")
                continue

            tmpl_ws.cell(row=row_idx, column=target_col_idx).value = price_val

    # 6) ذخیره تمپلیت
    tmpl_wb.save(FINAL_TEMPLATE_PATH)
    print(f"Template updated with prices from {converted_xlsx.name}")


# --- تابع جدید: خواندن JC PRODUCTS و پر کردن ستون «تگ باطری» ---
def fill_template_from_jc_products(jc_xlsx: Path):
    """
    از فایل JC PRODUCTS.xlsx سلول‌های E7:F28 را می‌خواند
    و مقدار ستون F را در price_of_products.xlsx در ستون 'تگ باطری'
    روبروی مدل مربوطه می‌نویسد.
    """
    print(f"Reading JC Products from: {jc_xlsx}")

    # 1) خواندن فایل JC PRODUCTS
    jc_wb = load_workbook(jc_xlsx, data_only=True)
    jc_ws = jc_wb.active

    # 2) باز کردن تمپلیت نهایی
    tmpl_wb = load_workbook(FINAL_TEMPLATE_PATH)
    tmpl_ws = tmpl_wb.active

    # 3) پیدا کردن شماره ستون "تگ باطری"
    tag_col_idx = None
    for cell in tmpl_ws[1]:
        if str(cell.value).strip() == TAG_COLUMN_HEADER:
            tag_col_idx = cell.column
            break

    if tag_col_idx is None:
        raise ValueError(f"ستونی با عنوان '{TAG_COLUMN_HEADER}' در ردیف اول تمپلیت پیدا نشد.")

    # 4) دیکشنری: اسم مدل (نرمال‌شده) → شماره ردیف در تمپلیت
    template_model_rows = {}
    for row in range(2, tmpl_ws.max_row + 1):
        val = tmpl_ws.cell(row=row, column=1).value
        if not val:
            continue
        template_model_rows[normalize_name(str(val))] = row

    # 5) خواندن رنج E7:F28 از JC PRODUCTS و نوشتن تگ‌ها در تمپلیت
    for model_val, tag_val in jc_ws.iter_rows(
        min_row=7, max_row=28, min_col=5, max_col=6, values_only=True
    ):
        if not model_val or not tag_val:
            continue

        norm_model = normalize_name(str(model_val))

        row_idx = template_model_rows.get(norm_model)

        # اگر پیدا نشد، یکبار هم با "iPhone " امتحان کن
        if not row_idx:
            generic_name = "iPhone " + str(model_val).strip()
            row_idx = template_model_rows.get(normalize_name(generic_name))

        if not row_idx:
            print(f"مدل '{model_val}' در تمپلیت برای تگ باتری پیدا نشد (JC PRODUCTS)")
            continue

        tmpl_ws.cell(row=row_idx, column=tag_col_idx).value = tag_val

    # 6) ذخیره تمپلیت
    tmpl_wb.save(FINAL_TEMPLATE_PATH)
    print(f"Template updated with tags from {jc_xlsx.name}")




def main():
    # تبدیل PDFها (در صورت نیاز)
    # convert_pdfs_to_excels()

    # فایل قیمتی که از PDF cell HIGH CAPACITY ساخته شده
    converted_file = OUTPUT_FOLDER / "cell  HIGH CAPACITY.xlsx"   # اگر اسم دقیق فرق دارد، اینجا را عوض کن
    fill_template_from_converted_excel(converted_file)

    # پر کردن ستون «تگ باطری» از JC PRODUCTS
    fill_template_from_jc_products(JC_PRODUCTS_PATH)


if __name__ == "__main__":
    main()
