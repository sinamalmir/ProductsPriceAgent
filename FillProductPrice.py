
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
TAG_COLUMN_HEADER    = "تگ باطری"              # از JC PRODUCTS
CAM_FRONT_HEADER     = "دوربین جلو پک کامل"    # از apple parts NORMAL (1)
SPEACKERS_HEADER      = "اسپیکر بالا"         # از apple parts NORMAL (1)
DOWN_SPEACKERS_HEADER = "اسپیکر پایین"      # از apple parts NORMAL (1)
FLAT_POWER_HEADER     = "فلت پاور"           # از apple parts NORMAL (1)
FLAT_VOLUME_HEADER    = "فلت ولوم"           # از apple parts NORMAL (1)    
FLAT_FLASH_CAMERA_HEADER = "فلت فلش دوربین"     # از apple parts NORMAL (1)
VIBRATION_HEADER      = "موتور ویبره"           # از apple parts NORMAL (1)
FRAME_HEADER          = "شاسی با فلت"                # از apple parts NORMAL (1)
FPC_FLAT_HEADER       = "فلت اف پی سی"                # از apple parts NORMAL (1)
FPC_RECEIVER_HEADER   = "فلت اف پی سی جی سی"                # از apple parts NORMAL (1)
FLAT_ANTENNA_HEADER    = "فلت آنتن"                # از apple parts NORMAL (1)
FLAT_POWER_WIRELESS_HEADER = "فلت وایرلس شارژ"    # از apple parts NORMAL (1)

LENZ_GLASS_HEADER       = "شیشه لنز بازاری"    # از apple parts NORMAL (1)
FACE_ID_TAG_HEADER     = "تگ فیس ایدی"    # از apple parts NORMAL (1)
FIX_FACE_ID_HEADER     = "تعمیر فیس آیدی"    # از apple parts NORMAL (1)
FIX_CAMERA_ERROR_HEADER  = "رفع ارور دوربین"    # از apple parts NORMAL (1)
SHIELD_HEADER          = "شیلد"    # از apple parts NORMAL (1)
ICLOUD_MOTHERBOARD_HEADER = "مادربرد آیکلود کامل"    # از apple parts NORMAL (1)


# --- جدید: فایل JC PRODUCTS و ستون «تگ باطری» ---
JC_PRODUCTS_PATH   = Path(r"C:\Users\Administrator\Desktop\pdfConvertor\ProductsPriceAgent\converted_excels\JC PRODUCTS NORMAL.xlsx")
APPLE_PARTS_NORMAL_PATH  = Path(r"C:\Users\Administrator\Desktop\pdfConvertor\ProductsPriceAgent\converted_excels\apple parts NORMAL (1).xlsx")
APPLE_PARTS_RAYAN_PATH  = Path(r"C:\Users\Administrator\Desktop\pdfConvertor\ProductsPriceAgent\converted_excels\Apple_Parts_rayan.xlsx")

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

def convert_pdfs_to_excels():
   
    OUTPUT_FOLDER.mkdir(parents=True, exist_ok=True)

   
    for f in OUTPUT_FOLDER.glob("*.xlsx"):
        try:
            f.unlink()
            print(f"Deleted old file -> {f}")
        except Exception as e:
            print(f"Could not delete file {f}: {e}")

    excel = win32.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False  # 🔹 جلوگیری از هرگونه پیغام مزاحم

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
        # برگرداندن تنظیمات اکسل به حالت عادی و بستن برنامه
        try:
            excel.DisplayAlerts = True
        except Exception:
            pass
        excel.Quit()


def build_template_model_index(tmpl_ws):
    """دیکشنری: اسم مدل نرمال‌شده → شماره ردیف در تمپلیت"""
    index = {}
    for row in range(2, tmpl_ws.max_row + 1):
        val = tmpl_ws.cell(row=row, column=1).value
        if not val:
            continue
        index[normalize_name(str(val))] = row
    return index


def find_column_by_header(ws, header_text: str):
    """شماره ستون را بر اساس متن هدر در ردیف ۱ برمی‌گرداند."""
    for cell in ws[1]:
        if str(cell.value).strip() == header_text:
            return cell.column
    return None


# --------- 1) پر کردن ستون «باطری روکاری» از cell HIGH CAPACITY ----------

def fill_template_from_converted_excel(converted_xlsx: Path):
    print(f"Reading data from: {converted_xlsx}")

    data_wb = load_workbook(converted_xlsx, data_only=True)
    data_ws = data_wb.active

    tmpl_wb = load_workbook(FINAL_TEMPLATE_PATH)
    tmpl_ws = tmpl_wb.active

    target_col_idx = find_column_by_header(tmpl_ws, TARGET_COLUMN_HEADER)
    if target_col_idx is None:
        raise ValueError(f"ستونی با عنوان '{TARGET_COLUMN_HEADER}' در ردیف اول تمپلیت پیدا نشد.")

    template_model_rows = build_template_model_index(tmpl_ws)

    for model_val, price_val in data_ws.iter_rows(
        min_row=6, min_col=3, max_col=4, values_only=True
    ):
        if not model_val or not price_val:
            continue

        if str(price_val).strip() == "-":
            continue

        key = str(model_val).strip().upper()
        mapped_models = MODEL_MAPPING.get(key)

        if not mapped_models:
            generic_name = "iPhone " + str(model_val).strip()
            norm = normalize_name(generic_name)
            if norm in template_model_rows:
                mapped_models = [generic_name]
            else:
                print(f"مدل ناشناخته در دیتا (cell HIGH CAPACITY): {model_val}  (رد شد)")
                continue

        for tmpl_model_name in mapped_models:
            row_idx = template_model_rows.get(normalize_name(tmpl_model_name))
            if not row_idx:
                print(f"مدل '{tmpl_model_name}' در تمپلیت پیدا نشد (از روی {model_val})")
                continue

            tmpl_ws.cell(row=row_idx, column=target_col_idx).value = price_val

    tmpl_wb.save(FINAL_TEMPLATE_PATH)
    print(f"Template updated with prices from {converted_xlsx.name}")


# --------- 2) پر کردن ستون «تگ باطری» از JC PRODUCTS ----------

def fill_template_from_jc_products(jc_xlsx: Path):
    print(f"Reading JC Products from: {jc_xlsx}")

    jc_wb = load_workbook(jc_xlsx, data_only=True)
    jc_ws = jc_wb.active

    tmpl_wb = load_workbook(FINAL_TEMPLATE_PATH)
    tmpl_ws = tmpl_wb.active

    tag_col_idx = find_column_by_header(tmpl_ws, TAG_COLUMN_HEADER)
    if tag_col_idx is None:
        raise ValueError(f"ستونی با عنوان '{TAG_COLUMN_HEADER}' در ردیف اول تمپلیت پیدا نشد.")

    template_model_rows = build_template_model_index(tmpl_ws)

    # E7:F28
    for model_val, tag_val in jc_ws.iter_rows(
        min_row=7, max_row=28, min_col=5, max_col=6, values_only=True
    ):
        if not model_val or not tag_val:
            continue

        norm_model = normalize_name(str(model_val))
        row_idx = template_model_rows.get(norm_model)

        if not row_idx:
            generic_name = "iPhone " + str(model_val).strip()
            row_idx = template_model_rows.get(normalize_name(generic_name))

        if not row_idx:
            print(f"مدل '{model_val}' در تمپلیت برای تگ باتری پیدا نشد (JC PRODUCTS)")
            continue

        tmpl_ws.cell(row=row_idx, column=tag_col_idx).value = tag_val

    tmpl_wb.save(FINAL_TEMPLATE_PATH)
    print(f"Template updated with tags from {jc_xlsx.name}")


# --------- 3) پر کردن ستون «دوربین جلو پک کامل» از apple parts NORMAL (1) ----------

def fill_template_from_apple_parts_normal(apple_xlsx: Path):
    """
    از فایل apple parts NORMAL (1).xlsx سلول‌های C4:D43 را می‌خواند
    و مقدار ستون D را در ستون «دوربین جلو پک کامل» تمپلیت می‌نویسد.
    """
    print(f"Reading Apple Parts Normal from: {apple_xlsx}")

    ap_wb = load_workbook(apple_xlsx, data_only=True)
    ap_ws = ap_wb.active

    tmpl_wb = load_workbook(FINAL_TEMPLATE_PATH)
    tmpl_ws = tmpl_wb.active

    cam_col_idx = find_column_by_header(tmpl_ws, CAM_FRONT_HEADER)
    if cam_col_idx is None:
        raise ValueError(f"ستونی با عنوان '{CAM_FRONT_HEADER}' در ردیف اول تمپلیت پیدا نشد.")

    template_model_rows = build_template_model_index(tmpl_ws)

    # C4:D43
    for model_val, price_val in ap_ws.iter_rows(
        min_row=4, max_row=43, min_col=3, max_col=4, values_only=True
    ):
        if not model_val or not price_val:
            continue

        norm_model = normalize_name(str(model_val))
        row_idx = template_model_rows.get(norm_model)

        if not row_idx:
            generic_name = "iPhone " + str(model_val).strip()
            row_idx = template_model_rows.get(normalize_name(generic_name))

        if not row_idx:
            print(f"مدل '{model_val}' در تمپلیت برای دوربین جلو پیدا نشد (apple parts NORMAL)")
            continue

        tmpl_ws.cell(row=row_idx, column=cam_col_idx).value = price_val

    tmpl_wb.save(FINAL_TEMPLATE_PATH)
    print(f"Template updated with front camera prices from {apple_xlsx.name}")


def fill_template_from_apple_parts_normal_downSpeackers(apple_xlsx: Path):

    print(f"Reading Apple Parts Normal from: {apple_xlsx}")

    ap_wb = load_workbook(apple_xlsx, data_only=True)
    ap_ws = ap_wb.active

    tmpl_wb = load_workbook(FINAL_TEMPLATE_PATH)
    tmpl_ws = tmpl_wb.active

    cam_col_idx = find_column_by_header(tmpl_ws, SPEACKERS_HEADER)
    if cam_col_idx is None:
        raise ValueError(f"ستونی با عنوان '{SPEACKERS_HEADER}' در ردیف اول تمپلیت پیدا نشد.")

    template_model_rows = build_template_model_index(tmpl_ws)

    # S4:T43
    for model_val, price_val in ap_ws.iter_rows(
        min_row=4, max_row=43, min_col=19, max_col=20, values_only=True
    ):
        if not model_val or not price_val:
            continue

        norm_model = normalize_name(str(model_val))
        row_idx = template_model_rows.get(norm_model)

        if not row_idx:
            generic_name = "iPhone " + str(model_val).strip()
            row_idx = template_model_rows.get(normalize_name(generic_name))

        if not row_idx:
            print(f"مدل '{model_val}' در تمپلیت برای دوربین جلو پیدا نشد (apple parts NORMAL)")
            continue

        tmpl_ws.cell(row=row_idx, column=cam_col_idx).value = price_val

    tmpl_wb.save(FINAL_TEMPLATE_PATH)
    print(f"Template updated with front camera prices from {apple_xlsx.name}")


def fill_template_from_apple_parts_normal_speakers(apple_xlsx: Path):

    print(f"Reading Apple Parts Normal from: {apple_xlsx}")

    ap_wb = load_workbook(apple_xlsx, data_only=True)
    ap_ws = ap_wb.active

    tmpl_wb = load_workbook(FINAL_TEMPLATE_PATH)
    tmpl_ws = tmpl_wb.active

    cam_col_idx = find_column_by_header(tmpl_ws, DOWN_SPEACKERS_HEADER)
    if cam_col_idx is None:
        raise ValueError(f"ستونی با عنوان '{DOWN_SPEACKERS_HEADER}' در ردیف اول تمپلیت پیدا نشد.")

    template_model_rows = build_template_model_index(tmpl_ws)

    # Q4:R43
    for model_val, price_val in ap_ws.iter_rows(
        min_row=4, max_row=43, min_col=17, max_col=18, values_only=True
    ):
        if not model_val or not price_val:
            continue

        norm_model = normalize_name(str(model_val))
        row_idx = template_model_rows.get(norm_model)

        if not row_idx:
            generic_name = "iPhone " + str(model_val).strip()
            row_idx = template_model_rows.get(normalize_name(generic_name))

        if not row_idx:
            print(f"مدل '{model_val}' در تمپلیت برای دوربین جلو پیدا نشد (apple parts NORMAL)")
            continue

        tmpl_ws.cell(row=row_idx, column=cam_col_idx).value = price_val

    tmpl_wb.save(FINAL_TEMPLATE_PATH)
    print(f"Template updated with front camera prices from {apple_xlsx.name}")


def fill_template_from_apple_parts_normal_flat_power(apple_xlsx: Path):

    print(f"Reading Apple Parts Normal from: {apple_xlsx}")

    ap_wb = load_workbook(apple_xlsx, data_only=True)
    ap_ws = ap_wb.active

    tmpl_wb = load_workbook(FINAL_TEMPLATE_PATH)
    tmpl_ws = tmpl_wb.active

    cam_col_idx = find_column_by_header(tmpl_ws, FLAT_POWER_HEADER)
    if cam_col_idx is None:
        raise ValueError(f"ستونی با عنوان '{FLAT_POWER_HEADER}' در ردیف اول تمپلیت پیدا نشد.")

    template_model_rows = build_template_model_index(tmpl_ws)

    # G4:H43
    for model_val, price_val in ap_ws.iter_rows(
        min_row=4, max_row=43, min_col=7, max_col=8, values_only=True
    ):
        if not model_val or not price_val:
            continue

        norm_model = normalize_name(str(model_val))
        row_idx = template_model_rows.get(norm_model)

        if not row_idx:
            generic_name = "iPhone " + str(model_val).strip()
            row_idx = template_model_rows.get(normalize_name(generic_name))

        if not row_idx:
            print(f"This model : '{model_val}' Not found in file (apple parts NORMAL)")
            continue

        tmpl_ws.cell(row=row_idx, column=cam_col_idx).value = price_val

    tmpl_wb.save(FINAL_TEMPLATE_PATH)
    print(f"Template updated with front camera prices from {apple_xlsx.name}")

def fill_template_from_apple_parts_normal_flat_power_volume(apple_xlsx: Path):

    print(f"Reading Apple Parts Normal from: {apple_xlsx}")

    ap_wb = load_workbook(apple_xlsx, data_only=True)
    ap_ws = ap_wb.active

    tmpl_wb = load_workbook(FINAL_TEMPLATE_PATH)
    tmpl_ws = tmpl_wb.active

    cam_col_idx = find_column_by_header(tmpl_ws, FLAT_VOLUME_HEADER)
    if cam_col_idx is None:
        raise ValueError(f"ستونی با عنوان '{FLAT_VOLUME_HEADER}' در ردیف اول تمپلیت پیدا نشد.")

    template_model_rows = build_template_model_index(tmpl_ws)

    # W4:X43
    for model_val, price_val in ap_ws.iter_rows(
        min_row=4, max_row=43, min_col=23, max_col=24, values_only=True
    ):
        if not model_val or not price_val:
            continue

        norm_model = normalize_name(str(model_val))
        row_idx = template_model_rows.get(norm_model)

        if not row_idx:
            generic_name = "iPhone " + str(model_val).strip()
            row_idx = template_model_rows.get(normalize_name(generic_name))

        if not row_idx:
            print(f"This model : '{model_val}' Not found in file (apple parts NORMAL)")
            continue

        tmpl_ws.cell(row=row_idx, column=cam_col_idx).value = price_val

    tmpl_wb.save(FINAL_TEMPLATE_PATH)
    print(f"Template updated with front camera prices from {apple_xlsx.name}")




def fill_template_from_apple_parts_normal_flat_flash_camera(apple_xlsx: Path):

    print(f"Reading Apple Parts Normal from: {apple_xlsx}")

    ap_wb = load_workbook(apple_xlsx, data_only=True)
    ap_ws = ap_wb.active

    tmpl_wb = load_workbook(FINAL_TEMPLATE_PATH)
    tmpl_ws = tmpl_wb.active

    cam_col_idx = find_column_by_header(tmpl_ws, FLAT_FLASH_CAMERA_HEADER)
    if cam_col_idx is None:
        raise ValueError(f"ستونی با عنوان '{FLAT_FLASH_CAMERA_HEADER}' در ردیف اول تمپلیت پیدا نشد.")

    template_model_rows = build_template_model_index(tmpl_ws)

    # AE4:AF43
    for model_val, price_val in ap_ws.iter_rows(
        min_row=4, max_row=43, min_col=31, max_col=32, values_only=True
    ):
        if not model_val or not price_val:
            continue

        norm_model = normalize_name(str(model_val))
        row_idx = template_model_rows.get(norm_model)

        if not row_idx:
            generic_name = "iPhone " + str(model_val).strip()
            row_idx = template_model_rows.get(normalize_name(generic_name))

        if not row_idx:
            print(f"This model : '{model_val}' Not found in file (apple parts NORMAL)")
            continue

        tmpl_ws.cell(row=row_idx, column=cam_col_idx).value = price_val

    tmpl_wb.save(FINAL_TEMPLATE_PATH)
    print(f"Template updated with front camera prices from {apple_xlsx.name}")



def fill_template_from_apple_parts_normal_vibration(apple_xlsx: Path):

    print(f"Reading Apple Parts Normal from: {apple_xlsx}")

    ap_wb = load_workbook(apple_xlsx, data_only=True)
    ap_ws = ap_wb.active

    tmpl_wb = load_workbook(FINAL_TEMPLATE_PATH)
    tmpl_ws = tmpl_wb.active

    cam_col_idx = find_column_by_header(tmpl_ws, VIBRATION_HEADER)
    if cam_col_idx is None:
        raise ValueError(f"ستونی با عنوان '{VIBRATION_HEADER}' در ردیف اول تمپلیت پیدا نشد.")

    template_model_rows = build_template_model_index(tmpl_ws)

    # U4:V43
    for model_val, price_val in ap_ws.iter_rows(
        min_row=4, max_row=43, min_col=21, max_col=22, values_only=True
    ):
        if not model_val or not price_val:
            continue

        norm_model = normalize_name(str(model_val))
        row_idx = template_model_rows.get(norm_model)

        if not row_idx:
            generic_name = "iPhone " + str(model_val).strip()
            row_idx = template_model_rows.get(normalize_name(generic_name))

        if not row_idx:
            print(f"This model : '{model_val}' Not found in file (apple parts NORMAL)")
            continue

        tmpl_ws.cell(row=row_idx, column=cam_col_idx).value = price_val

    tmpl_wb.save(FINAL_TEMPLATE_PATH)
    print(f"Template updated with front camera prices from {apple_xlsx.name}")


#rayan pdf file
def fill_template_from_apple_parts_rayan_frame(apple_xlsx: Path):

    print(f"Reading Apple Parts rayan from: {apple_xlsx}")

    ap_wb = load_workbook(apple_xlsx, data_only=True)
    ap_ws = ap_wb.active

    tmpl_wb = load_workbook(FINAL_TEMPLATE_PATH)
    tmpl_ws = tmpl_wb.active

    cam_col_idx = find_column_by_header(tmpl_ws, FRAME_HEADER)
    if cam_col_idx is None:
        raise ValueError(f"ستونی با عنوان '{FRAME_HEADER}' در ردیف اول تمپلیت پیدا نشد.")

    template_model_rows = build_template_model_index(tmpl_ws)

    # M4:N43
    for model_val, price_val in ap_ws.iter_rows(
        min_row=4, max_row=38, min_col=13, max_col=14, values_only=True
    ):
        if not model_val or not price_val:
            continue

        norm_model = normalize_name(str(model_val))
        row_idx = template_model_rows.get(norm_model)

        if not row_idx:
            generic_name = "iPhone " + str(model_val).strip()
            row_idx = template_model_rows.get(normalize_name(generic_name))

        if not row_idx:
            print(f"This model : '{model_val}' Not found in file (apple parts NORMAL)")
            continue

        tmpl_ws.cell(row=row_idx, column=cam_col_idx).value = price_val

    tmpl_wb.save(FINAL_TEMPLATE_PATH)
    print(f"Template updated with front camera prices from {apple_xlsx.name}")


def fill_template_from_apple_parts_normal_fpc_flat(apple_xlsx: Path):

    print(f"Reading Apple Parts Normal from: {apple_xlsx}")

    ap_wb = load_workbook(apple_xlsx, data_only=True)
    ap_ws = ap_wb.active

    tmpl_wb = load_workbook(FINAL_TEMPLATE_PATH)
    tmpl_ws = tmpl_wb.active

    cam_col_idx = find_column_by_header(tmpl_ws, VIBRATION_HEADER)
    if cam_col_idx is None:
        raise ValueError(f"ستونی با عنوان '{VIBRATION_HEADER}' در ردیف اول تمپلیت پیدا نشد.")

    template_model_rows = build_template_model_index(tmpl_ws)

    # AM4:AN3
    for model_val, price_val in ap_ws.iter_rows(
        min_row=4, max_row=43, min_col=39, max_col=40, values_only=True
    ):
        if not model_val or not price_val:
            continue

        norm_model = normalize_name(str(model_val))
        row_idx = template_model_rows.get(norm_model)

        if not row_idx:
            generic_name = "iPhone " + str(model_val).strip()
            row_idx = template_model_rows.get(normalize_name(generic_name))

        if not row_idx:
            print(f"This model : '{model_val}' Not found in file (apple parts NORMAL)")
            continue

        tmpl_ws.cell(row=row_idx, column=cam_col_idx).value = price_val

    tmpl_wb.save(FINAL_TEMPLATE_PATH)
    print(f"Template updated with front camera prices from {apple_xlsx.name}")


def fill_template_from_apple_parts_normal_fpc_receiver(apple_xlsx: Path):

    print(f"Reading Apple Parts Normal from: {apple_xlsx}")

    ap_wb = load_workbook(apple_xlsx, data_only=True)
    ap_ws = ap_wb.active

    tmpl_wb = load_workbook(FINAL_TEMPLATE_PATH)
    tmpl_ws = tmpl_wb.active

    cam_col_idx = find_column_by_header(tmpl_ws, FPC_RECEIVER_HEADER)
    if cam_col_idx is None:
        raise ValueError(f"ستونی با عنوان '{FPC_RECEIVER_HEADER}' در ردیف اول تمپلیت پیدا نشد.")

    template_model_rows = build_template_model_index(tmpl_ws)

    # G4:HN3
    for model_val, price_val in ap_ws.iter_rows(
        min_row=7, max_row=28, min_col=7, max_col=8, values_only=True
    ):
        if not model_val or not price_val:
            continue

        norm_model = normalize_name(str(model_val))
        row_idx = template_model_rows.get(norm_model)

        if not row_idx:
            generic_name = "iPhone " + str(model_val).strip()
            row_idx = template_model_rows.get(normalize_name(generic_name))

        if not row_idx:
            print(f"This model : '{model_val}' Not found in file (apple parts NORMAL)")
            continue

        tmpl_ws.cell(row=row_idx, column=cam_col_idx).value = price_val

    tmpl_wb.save(FINAL_TEMPLATE_PATH)
    print(f"Template updated with front camera prices from {apple_xlsx.name}")


def fill_template_from_apple_parts_normal_flat_antenna(apple_xlsx: Path):

    print(f"Reading Apple Parts Normal from: {apple_xlsx}")

    ap_wb = load_workbook(apple_xlsx, data_only=True)
    ap_ws = ap_wb.active

    tmpl_wb = load_workbook(FINAL_TEMPLATE_PATH)
    tmpl_ws = tmpl_wb.active

    cam_col_idx = find_column_by_header(tmpl_ws, FLAT_ANTENNA_HEADER)
    if cam_col_idx is None:
        raise ValueError(f"ستونی با عنوان '{FLAT_ANTENNA_HEADER}' در ردیف اول تمپلیت پیدا نشد.")

    template_model_rows = build_template_model_index(tmpl_ws)

    # AW4:AX28
    for model_val, price_val in ap_ws.iter_rows(
        min_row=5, max_row=43, min_col=49, max_col=50, values_only=True
    ):
        if not model_val or not price_val:
            continue

        norm_model = normalize_name(str(model_val))
        row_idx = template_model_rows.get(norm_model)

        if not row_idx:
            generic_name = "iPhone " + str(model_val).strip()
            row_idx = template_model_rows.get(normalize_name(generic_name))

        if not row_idx:
            print(f"This model : '{model_val}' Not found in file (apple parts NORMAL)")
            continue

        tmpl_ws.cell(row=row_idx, column=cam_col_idx).value = price_val

    tmpl_wb.save(FINAL_TEMPLATE_PATH)
    print(f"Template updated with front camera prices from {apple_xlsx.name}")



def fill_template_from_apple_parts_normal_lenz_glass(apple_xlsx: Path):

    print(f"Reading Apple Parts Normal from: {apple_xlsx}")

    ap_wb = load_workbook(apple_xlsx, data_only=True)
    ap_ws = ap_wb.active

    tmpl_wb = load_workbook(FINAL_TEMPLATE_PATH)
    tmpl_ws = tmpl_wb.active

    cam_col_idx = find_column_by_header(tmpl_ws, FLAT_POWER_WIRELESS_HEADER)
    if cam_col_idx is None:
        raise ValueError(f"ستونی با عنوان '{FLAT_POWER_WIRELESS_HEADER}' در ردیف اول تمپلیت پیدا نشد.")

    template_model_rows = build_template_model_index(tmpl_ws)

    # K6:L43
    for model_val, price_val in ap_ws.iter_rows(
        min_row=6, max_row=43, min_col=11, max_col=12, values_only=True
    ):
        if not model_val or not price_val:
            continue

        norm_model = normalize_name(str(model_val))
        row_idx = template_model_rows.get(norm_model)

        if not row_idx:
            generic_name = "iPhone " + str(model_val).strip()
            row_idx = template_model_rows.get(normalize_name(generic_name))

        if not row_idx:
            print(f"This model : '{model_val}' Not found in file (apple parts NORMAL)")
            continue

        tmpl_ws.cell(row=row_idx, column=cam_col_idx).value = price_val

    tmpl_wb.save(FINAL_TEMPLATE_PATH)
    print(f"Template updated with front camera prices from {apple_xlsx.name}")


def fill_template_from_apple_parts_normal_flat_wireless(apple_xlsx: Path):

    print(f"Reading Apple Parts Normal from: {apple_xlsx}")

    ap_wb = load_workbook(apple_xlsx, data_only=True)
    ap_ws = ap_wb.active

    tmpl_wb = load_workbook(FINAL_TEMPLATE_PATH)
    tmpl_ws = tmpl_wb.active

    cam_col_idx = find_column_by_header(tmpl_ws, LENZ_GLASS_HEADER)
    if cam_col_idx is None:
        raise ValueError(f"ستونی با عنوان '{LENZ_GLASS_HEADER}' در ردیف اول تمپلیت پیدا نشد.")

    template_model_rows = build_template_model_index(tmpl_ws)

    # O6:P43
    for model_val, price_val in ap_ws.iter_rows(
        min_row=6, max_row=43, min_col=15, max_col=16, values_only=True
    ):
        if not model_val or not price_val:
            continue

        norm_model = normalize_name(str(model_val))
        row_idx = template_model_rows.get(norm_model)

        if not row_idx:
            generic_name = "iPhone " + str(model_val).strip()
            row_idx = template_model_rows.get(normalize_name(generic_name))

        if not row_idx:
            print(f"This model : '{model_val}' Not found in file (apple parts NORMAL)")
            continue

        tmpl_ws.cell(row=row_idx, column=cam_col_idx).value = price_val

    tmpl_wb.save(FINAL_TEMPLATE_PATH)
    print(f"Template updated with front camera prices from {apple_xlsx.name}")


def fill_template_from_apple_parts_normal_face_id_tag(apple_xlsx: Path):

    print(f"Reading Apple Parts Normal from: {apple_xlsx}")

    ap_wb = load_workbook(apple_xlsx, data_only=True)
    ap_ws = ap_wb.active

    tmpl_wb = load_workbook(FINAL_TEMPLATE_PATH)
    tmpl_ws = tmpl_wb.active

    cam_col_idx = find_column_by_header(tmpl_ws, FACE_ID_TAG_HEADER)
    if cam_col_idx is None:
        raise ValueError(f"ستونی با عنوان '{FACE_ID_TAG_HEADER}' در ردیف اول تمپلیت پیدا نشد.")

    template_model_rows = build_template_model_index(tmpl_ws)

    # Q7:R28
    for model_val, price_val in ap_ws.iter_rows(
        min_row=7, max_row=28, min_col=17, max_col=18, values_only=True
    ):
        if not model_val or not price_val:
            continue

        norm_model = normalize_name(str(model_val))
        row_idx = template_model_rows.get(norm_model)

        if not row_idx:
            generic_name = "iPhone " + str(model_val).strip()
            row_idx = template_model_rows.get(normalize_name(generic_name))

        if not row_idx:
            print(f"This model : '{model_val}' Not found in file (apple parts NORMAL)")
            continue

        tmpl_ws.cell(row=row_idx, column=cam_col_idx).value = price_val

    tmpl_wb.save(FINAL_TEMPLATE_PATH)
    print(f"Template updated with front camera prices from {apple_xlsx.name}")



def fill_template_from_apple_parts_normal_fix_face_id(apple_xlsx: Path):

    print(f"Reading Apple Parts Normal from: {apple_xlsx}")

    ap_wb = load_workbook(apple_xlsx, data_only=True)
    ap_ws = ap_wb.active

    tmpl_wb = load_workbook(FINAL_TEMPLATE_PATH)
    tmpl_ws = tmpl_wb.active

    cam_col_idx = find_column_by_header(tmpl_ws, FIX_FACE_ID_HEADER)
    if cam_col_idx is None:
        raise ValueError(f"ستونی با عنوان '{FIX_FACE_ID_HEADER}' در ردیف اول تمپلیت پیدا نشد.")

    template_model_rows = build_template_model_index(tmpl_ws)

    # C7:D28
    for model_val, price_val in ap_ws.iter_rows(
        min_row=7, max_row=28, min_col=3, max_col=4, values_only=True
    ):
        if not model_val or not price_val:
            continue

        norm_model = normalize_name(str(model_val))
        row_idx = template_model_rows.get(norm_model)

        if not row_idx:
            generic_name = "iPhone " + str(model_val).strip()
            row_idx = template_model_rows.get(normalize_name(generic_name))

        if not row_idx:
            print(f"This model : '{model_val}' Not found in file (apple parts NORMAL)")
            continue

        tmpl_ws.cell(row=row_idx, column=cam_col_idx).value = price_val

    tmpl_wb.save(FINAL_TEMPLATE_PATH)
    print(f"Template updated with front camera prices from {apple_xlsx.name}")


def fill_template_from_apple_parts_normal_fix_camera_error(apple_xlsx: Path):

    print(f"Reading Apple Parts Normal from: {apple_xlsx}")

    ap_wb = load_workbook(apple_xlsx, data_only=True)
    ap_ws = ap_wb.active

    tmpl_wb = load_workbook(FINAL_TEMPLATE_PATH)
    tmpl_ws = tmpl_wb.active

    cam_col_idx = find_column_by_header(tmpl_ws, FIX_CAMERA_ERROR_HEADER)
    if cam_col_idx is None:
        raise ValueError(f"ستونی با عنوان '{FIX_CAMERA_ERROR_HEADER}' در ردیف اول تمپلیت پیدا نشد.")

    template_model_rows = build_template_model_index(tmpl_ws)

    # M7:N28
    for model_val, price_val in ap_ws.iter_rows(
        min_row=7, max_row=28, min_col=13, max_col=14, values_only=True
    ):
        if not model_val or not price_val:
            continue

        norm_model = normalize_name(str(model_val))
        row_idx = template_model_rows.get(norm_model)

        if not row_idx:
            generic_name = "iPhone " + str(model_val).strip()
            row_idx = template_model_rows.get(normalize_name(generic_name))

        if not row_idx:
            print(f"This model : '{model_val}' Not found in file (apple parts NORMAL)")
            continue

        tmpl_ws.cell(row=row_idx, column=cam_col_idx).value = price_val

    tmpl_wb.save(FINAL_TEMPLATE_PATH)
    print(f"Template updated with front camera prices from {apple_xlsx.name}")


def fill_template_from_apple_parts_normal_shield(apple_xlsx: Path):

    print(f"Reading Apple Parts Normal from: {apple_xlsx}")

    ap_wb = load_workbook(apple_xlsx, data_only=True)
    ap_ws = ap_wb.active

    tmpl_wb = load_workbook(FINAL_TEMPLATE_PATH)
    tmpl_ws = tmpl_wb.active

    cam_col_idx = find_column_by_header(tmpl_ws, SHIELD_HEADER)
    if cam_col_idx is None:
        raise ValueError(f"ستونی با عنوان '{SHIELD_HEADER}' در ردیف اول تمپلیت پیدا نشد.")

    template_model_rows = build_template_model_index(tmpl_ws)

    # Y6:Z43
    for model_val, price_val in ap_ws.iter_rows(
        min_row=6, max_row=43, min_col=25, max_col=26, values_only=True
    ):
        if not model_val or not price_val:
            continue

        norm_model = normalize_name(str(model_val))
        row_idx = template_model_rows.get(norm_model)

        if not row_idx:
            generic_name = "iPhone " + str(model_val).strip()
            row_idx = template_model_rows.get(normalize_name(generic_name))

        if not row_idx:
            print(f"This model : '{model_val}' Not found in file (apple parts NORMAL)")
            continue

        tmpl_ws.cell(row=row_idx, column=cam_col_idx).value = price_val

    tmpl_wb.save(FINAL_TEMPLATE_PATH)
    print(f"Template updated with front camera prices from {apple_xlsx.name}")




def main():
   # if u need to convert pdfs to excel every time --> uncommnet this code 

    convert_pdfs_to_excels()

   
    # converted_file = OUTPUT_FOLDER / "cell  HIGH CAPACITY.xlsx"  
    # fill_template_from_converted_excel(converted_file)

   
    # fill_template_from_jc_products(JC_PRODUCTS_PATH)

  
    # fill_template_from_apple_parts_normal(APPLE_PARTS_NORMAL_PATH)

    # fill_template_from_apple_parts_normal_speakers(APPLE_PARTS_NORMAL_PATH)

    # fill_template_from_apple_parts_normal_downSpeackers(APPLE_PARTS_NORMAL_PATH)
    # fill_template_from_apple_parts_normal_flat_power_volume(APPLE_PARTS_NORMAL_PATH)
    # fill_template_from_apple_parts_normal_flat_flash_camera(APPLE_PARTS_NORMAL_PATH)
    # fill_template_from_apple_parts_normal_vibration(APPLE_PARTS_NORMAL_PATH)

    # fill_template_from_apple_parts_rayan_frame(APPLE_PARTS_RAYAN_PATH)

    # fill_template_from_apple_parts_normal_fpc_flat(APPLE_PARTS_NORMAL_PATH)

    # fill_template_from_apple_parts_normal_fpc_receiver(JC_PRODUCTS_PATH)
    # fill_template_from_apple_parts_normal_flat_antenna(APPLE_PARTS_NORMAL_PATH)

    # fill_template_from_apple_parts_normal_flat_wireless(APPLE_PARTS_NORMAL_PATH)

    # fill_template_from_apple_parts_normal_flat_wireless(APPLE_PARTS_NORMAL_PATH)

    # fill_template_from_apple_parts_normal_face_id_tag(JC_PRODUCTS_PATH)

    # fill_template_from_apple_parts_normal_fix_face_id(JC_PRODUCTS_PATH)

    # fill_template_from_apple_parts_normal_fix_camera_error(JC_PRODUCTS_PATH)

    # fill_template_from_apple_parts_normal_shield(APPLE_PARTS_NORMAL_PATH)

if __name__ == "__main__":
    main()