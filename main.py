from pathlib import Path
import re
import logging

import win32com.client as win32
from openpyxl import load_workbook

from telegram import Update
from telegram.ext import (
    ApplicationBuilder,
    CommandHandler,
    MessageHandler,
    ContextTypes,
    filters,
)

# ====================== TELEGRAM SETTINGS ======================

# Put your bot token here
BOT_TOKEN = "7266866129:AAHkXCVpqoy4uFLZzHlSTS-ycI_C-1FKKmQ"

# Configure logging
logging.basicConfig(
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    level=logging.INFO,
)
logger = logging.getLogger(__name__)

# ======================== FILE SETTINGS ========================

TEMPLATE_IMPORT_XLSX = Path(
    r"D:\Projects\pdfConvertor\ProductsPriceAgent\EXCEL\template_pdf_imports.xlsx"
)
PDF_FOLDER = Path(
    r"D:\Projects\pdfConvertor\ProductsPriceAgent\PDF"
)
OUTPUT_FOLDER = Path(
    r"D:\Projects\pdfConvertor\ProductsPriceAgent\converted_excels"
)

# Excel Power Query name
QUERY_NAME = "Query1"

# Final template file (main product list)
FINAL_TEMPLATE_PATH = Path(
    r"D:\Projects\pdfConvertor\ProductsPriceAgent\excel_data\FinalTemplate.xlsx"
)

# Column headers in final template
TARGET_COLUMN_HEADER = "باطری روکاری"
TAG_COLUMN_HEADER = "تگ باطری"
CAM_FRONT_HEADER = "دوربین جلو پک کامل"
SPEACKERS_HEADER = "اسپیکر بالا"
DOWN_SPEACKERS_HEADER = "اسپیکر پایین"
FLAT_POWER_HEADER = "فلت پاور"
FLAT_VOLUME_HEADER = "فلت ولوم"
FLAT_FLASH_CAMERA_HEADER = "فلت فلش دوربین"
VIBRATION_HEADER = "موتور ویبره"
FRAME_HEADER = "شاسی با فلت"
FPC_FLAT_HEADER = "فلت اف پی سی"
FPC_RECEIVER_HEADER = "فلت اف پی سی جی سی"
FLAT_ANTENNA_HEADER = "فلت آنتن"
FLAT_POWER_WIRELESS_HEADER = "فلت وایرلس شارژ"

LENZ_GLASS_HEADER = "شیشه لنز بازاری"
FACE_ID_TAG_HEADER = "تگ فیس ایدی"
FIX_FACE_ID_HEADER = "تعمیر فیس آیدی"
FIX_CAMERA_ERROR_HEADER = "رفع ارور دوربین"
SHIELD_HEADER = "شیلد"
ICLOUD_MOTHERBOARD_HEADER = "مادربرد آیکلود کامل"

# External Excel sources (expected converted files)
JC_PRODUCTS_PATH = Path(
    r"D:\Projects\pdfConvertor\ProductsPriceAgent\converted_excels\JC PRODUCTS NORMAL.xlsx"
)
APPLE_PARTS_NORMAL_PATH = Path(
    r"D:\Projects\pdfConvertor\ProductsPriceAgent\converted_excels\apple parts NORMAL (1).xlsx"
)
APPLE_PARTS_RAYAN_PATH = Path(
    r"D:\Projects\pdfConvertor\ProductsPriceAgent\converted_excels\Apple_Parts_rayan.xlsx"
)

# ======================== CORE HELPERS =========================

def update_query_pdf_path(query, new_pdf_path: Path):
    """Update Power Query M formula with a new PDF file path."""
    formula = query.Formula
    pdf_str = new_pdf_path.resolve().as_posix()

    new_formula = re.sub(
        r'File\.Contents\(".*?"\)',   # pattern
        f'File.Contents("{pdf_str}")',# replacement
        formula,                      # <-- missing argument (string)
    )
    query.Formula = new_formula


def normalize_name(s: str) -> str:
    """Normalize model names (lowercase, collapse spaces) for robust matching."""
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


def convert_pdfs_to_excels():
    """
    Convert all PDFs in PDF_FOLDER to Excel files in OUTPUT_FOLDER
    using Excel Power Query template.
    Existing .xlsx files in OUTPUT_FOLDER will be deleted first.
    """
    OUTPUT_FOLDER.mkdir(parents=True, exist_ok=True)

    # Delete old converted Excel files
    for f in OUTPUT_FOLDER.glob("*.xlsx"):
        try:
            f.unlink()
            logger.info(f"Deleted old file -> {f}")
        except Exception as e:
            logger.warning(f"Could not delete file {f}: {e}")

    excel = win32.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False

    try:
        for pdf_path in PDF_FOLDER.glob("*.pdf"):
            logger.info(f"Processing PDF -> {pdf_path.name}")

            wb = excel.Workbooks.Open(str(TEMPLATE_IMPORT_XLSX))
            query = wb.Queries(QUERY_NAME)

            update_query_pdf_path(query, pdf_path)

            wb.RefreshAll()
            excel.CalculateUntilAsyncQueriesDone()

            out_path = OUTPUT_FOLDER / f"{pdf_path.stem}.xlsx"
            wb.SaveAs(str(out_path), FileFormat=51)
            wb.Close(SaveChanges=False)

            logger.info(f"Saved Excel -> {out_path}")

    finally:
        try:
            excel.DisplayAlerts = True
        except Exception:
            pass
        excel.Quit()


def build_template_model_index(tmpl_ws):
    """Build a dict: normalized model name -> row index in final template."""
    index = {}
    for row in range(2, tmpl_ws.max_row + 1):
        val = tmpl_ws.cell(row=row, column=1).value
        if not val:
            continue
        index[normalize_name(str(val))] = row
    return index


def find_column_by_header(ws, header_text: str):
    """Return column index for a given header text in row 1."""
    for cell in ws[1]:
        if str(cell.value).strip() == header_text:
            return cell.column
    return None


# ========================= FILL FUNCTIONS =========================
# (same logic as your code, with English comments/logs)


def fill_template_from_converted_excel(converted_xlsx: Path):
    """
    Read model/price from a converted Excel file (e.g. cell HIGH CAPACITY.xlsx)
    and fill the 'باطری روکاری' column in final template.
    """
    logger.info(f"Reading data from: {converted_xlsx}")

    data_wb = load_workbook(converted_xlsx, data_only=True)
    data_ws = data_wb.active

    tmpl_wb = load_workbook(FINAL_TEMPLATE_PATH)
    tmpl_ws = tmpl_wb.active

    target_col_idx = find_column_by_header(tmpl_ws, TARGET_COLUMN_HEADER)
    if target_col_idx is None:
        raise ValueError(
            f"Header '{TARGET_COLUMN_HEADER}' not found in final template."
        )

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
                logger.warning(
                    f"Unknown model in cell HIGH CAPACITY: {model_val}  (skipped)"
                )
                continue

        for tmpl_model_name in mapped_models:
            row_idx = template_model_rows.get(normalize_name(tmpl_model_name))
            if not row_idx:
                logger.warning(
                    f"Model '{tmpl_model_name}' not found in final template (from {model_val})"
                )
                continue

            tmpl_ws.cell(row=row_idx, column=target_col_idx).value = price_val

    tmpl_wb.save(FINAL_TEMPLATE_PATH)
    logger.info(f"Template updated with prices from {converted_xlsx.name}")


def fill_template_from_jc_products(jc_xlsx: Path):
    """
    Read JC PRODUCTS Excel (E7:F28) and fill 'تگ باطری' column in final template.
    """
    logger.info(f"Reading JC Products from: {jc_xlsx}")

    jc_wb = load_workbook(jc_xlsx, data_only=True)
    jc_ws = jc_wb.active

    tmpl_wb = load_workbook(FINAL_TEMPLATE_PATH)
    tmpl_ws = tmpl_wb.active

    tag_col_idx = find_column_by_header(tmpl_ws, TAG_COLUMN_HEADER)
    if tag_col_idx is None:
        raise ValueError(f"Header '{TAG_COLUMN_HEADER}' not found in final template.")

    template_model_rows = build_template_model_index(tmpl_ws)

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
            logger.warning(
                f"Model '{model_val}' for battery tag not found in final template (JC PRODUCTS)"
            )
            continue

        tmpl_ws.cell(row=row_idx, column=tag_col_idx).value = tag_val

    tmpl_wb.save(FINAL_TEMPLATE_PATH)
    logger.info(f"Template updated with tags from {jc_xlsx.name}")


def fill_template_from_apple_parts_normal(apple_xlsx: Path):
    """Fill 'دوربین جلو پک کامل' from C4:D43."""
    logger.info(f"Reading Apple Parts Normal from: {apple_xlsx}")

    ap_wb = load_workbook(apple_xlsx, data_only=True)
    ap_ws = ap_wb.active

    tmpl_wb = load_workbook(FINAL_TEMPLATE_PATH)
    tmpl_ws = tmpl_wb.active

    col_idx = find_column_by_header(tmpl_ws, CAM_FRONT_HEADER)
    if col_idx is None:
        raise ValueError(f"Header '{CAM_FRONT_HEADER}' not found in final template.")

    template_model_rows = build_template_model_index(tmpl_ws)

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
            logger.warning(
                f"Model '{model_val}' for front camera not found in final template (apple parts NORMAL)"
            )
            continue

        tmpl_ws.cell(row=row_idx, column=col_idx).value = price_val

    tmpl_wb.save(FINAL_TEMPLATE_PATH)
    logger.info(
        f"Template updated with front camera prices from {apple_xlsx.name}"
    )


def fill_template_from_apple_parts_normal_downSpeackers(apple_xlsx: Path):
    """Fill 'اسپیکر بالا' from S4:T43."""
    logger.info(f"Reading Apple Parts Normal (upper speaker) from: {apple_xlsx}")

    ap_wb = load_workbook(apple_xlsx, data_only=True)
    ap_ws = ap_wb.active

    tmpl_wb = load_workbook(FINAL_TEMPLATE_PATH)
    tmpl_ws = tmpl_wb.active

    col_idx = find_column_by_header(tmpl_ws, SPEACKERS_HEADER)
    if col_idx is None:
        raise ValueError(f"Header '{SPEACKERS_HEADER}' not found in final template.")

    template_model_rows = build_template_model_index(tmpl_ws)

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
            logger.warning(
                f"Model '{model_val}' for upper speaker not found in final template (apple parts NORMAL)"
            )
            continue

        tmpl_ws.cell(row=row_idx, column=col_idx).value = price_val

    tmpl_wb.save(FINAL_TEMPLATE_PATH)
    logger.info(
        f"Template updated with upper speaker prices from {apple_xlsx.name}"
    )


def fill_template_from_apple_parts_normal_speakers(apple_xlsx: Path):
    """Fill 'اسپیکر پایین' from Q4:R43."""
    logger.info(f"Reading Apple Parts Normal (lower speaker) from: {apple_xlsx}")

    ap_wb = load_workbook(apple_xlsx, data_only=True)
    ap_ws = ap_wb.active

    tmpl_wb = load_workbook(FINAL_TEMPLATE_PATH)
    tmpl_ws = tmpl_wb.active

    col_idx = find_column_by_header(tmpl_ws, DOWN_SPEACKERS_HEADER)
    if col_idx is None:
        raise ValueError(
            f"Header '{DOWN_SPEACKERS_HEADER}' not found in final template."
        )

    template_model_rows = build_template_model_index(tmpl_ws)

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
            logger.warning(
                f"Model '{model_val}' for lower speaker not found in final template (apple parts NORMAL)"
            )
            continue

        tmpl_ws.cell(row=row_idx, column=col_idx).value = price_val

    tmpl_wb.save(FINAL_TEMPLATE_PATH)
    logger.info(
        f"Template updated with lower speaker prices from {apple_xlsx.name}"
    )


def fill_template_from_apple_parts_normal_flat_power(apple_xlsx: Path):
    """Fill 'فلت پاور' from G4:H43."""
    logger.info(f"Reading Apple Parts Normal (flat power) from: {apple_xlsx}")

    ap_wb = load_workbook(apple_xlsx, data_only=True)
    ap_ws = ap_wb.active

    tmpl_wb = load_workbook(FINAL_TEMPLATE_PATH)
    tmpl_ws = tmpl_wb.active

    col_idx = find_column_by_header(tmpl_ws, FLAT_POWER_HEADER)
    if col_idx is None:
        raise ValueError(f"Header '{FLAT_POWER_HEADER}' not found in final template.")

    template_model_rows = build_template_model_index(tmpl_ws)

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
            logger.warning(
                f"Model '{model_val}' for flat power not found in final template (apple parts NORMAL)"
            )
            continue

        tmpl_ws.cell(row=row_idx, column=col_idx).value = price_val

    tmpl_wb.save(FINAL_TEMPLATE_PATH)
    logger.info(
        f"Template updated with flat power prices from {apple_xlsx.name}"
    )


def fill_template_from_apple_parts_normal_flat_power_volume(apple_xlsx: Path):
    """Fill 'فلت ولوم' from W4:X43."""
    logger.info(f"Reading Apple Parts Normal (flat volume) from: {apple_xlsx}")

    ap_wb = load_workbook(apple_xlsx, data_only=True)
    ap_ws = ap_wb.active

    tmpl_wb = load_workbook(FINAL_TEMPLATE_PATH)
    tmpl_ws = tmpl_wb.active

    col_idx = find_column_by_header(tmpl_ws, FLAT_VOLUME_HEADER)
    if col_idx is None:
        raise ValueError(
            f"Header '{FLAT_VOLUME_HEADER}' not found in final template."
        )

    template_model_rows = build_template_model_index(tmpl_ws)

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
            logger.warning(
                f"Model '{model_val}' for flat volume not found in final template (apple parts NORMAL)"
            )
            continue

        tmpl_ws.cell(row=row_idx, column=col_idx).value = price_val

    tmpl_wb.save(FINAL_TEMPLATE_PATH)
    logger.info(
        f"Template updated with flat volume prices from {apple_xlsx.name}"
    )


def fill_template_from_apple_parts_normal_flat_flash_camera(apple_xlsx: Path):
    """Fill 'فلت فلش دوربین' from AE4:AF43."""
    logger.info(f"Reading Apple Parts Normal (flash flat) from: {apple_xlsx}")

    ap_wb = load_workbook(apple_xlsx, data_only=True)
    ap_ws = ap_wb.active

    tmpl_wb = load_workbook(FINAL_TEMPLATE_PATH)
    tmpl_ws = tmpl_wb.active

    col_idx = find_column_by_header(tmpl_ws, FLAT_FLASH_CAMERA_HEADER)
    if col_idx is None:
        raise ValueError(
            f"Header '{FLAT_FLASH_CAMERA_HEADER}' not found in final template."
        )

    template_model_rows = build_template_model_index(tmpl_ws)

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
            logger.warning(
                f"Model '{model_val}' for flash flat not found in final template (apple parts NORMAL)"
            )
            continue

        tmpl_ws.cell(row=row_idx, column=col_idx).value = price_val

    tmpl_wb.save(FINAL_TEMPLATE_PATH)
    logger.info(
        f"Template updated with flash flat prices from {apple_xlsx.name}"
    )


def fill_template_from_apple_parts_normal_vibration(apple_xlsx: Path):
    """Fill 'موتور ویبره' from U4:V43."""
    logger.info(f"Reading Apple Parts Normal (vibration) from: {apple_xlsx}")

    ap_wb = load_workbook(apple_xlsx, data_only=True)
    ap_ws = ap_wb.active

    tmpl_wb = load_workbook(FINAL_TEMPLATE_PATH)
    tmpl_ws = tmpl_wb.active

    col_idx = find_column_by_header(tmpl_ws, VIBRATION_HEADER)
    if col_idx is None:
        raise ValueError(
            f"Header '{VIBRATION_HEADER}' not found in final template."
        )

    template_model_rows = build_template_model_index(tmpl_ws)

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
            logger.warning(
                f"Model '{model_val}' for vibration not found in final template (apple parts NORMAL)"
            )
            continue

        tmpl_ws.cell(row=row_idx, column=col_idx).value = price_val

    tmpl_wb.save(FINAL_TEMPLATE_PATH)
    logger.info(
        f"Template updated with vibration prices from {apple_xlsx.name}"
    )


def fill_template_from_apple_parts_rayan_frame(apple_xlsx: Path):
    """Fill 'شاسی با فلت' from M4:N38 in Apple_Parts_rayan.xlsx."""
    logger.info(f"Reading Apple Parts rayan (frame) from: {apple_xlsx}")

    ap_wb = load_workbook(apple_xlsx, data_only=True)
    ap_ws = ap_wb.active

    tmpl_wb = load_workbook(FINAL_TEMPLATE_PATH)
    tmpl_ws = tmpl_wb.active

    col_idx = find_column_by_header(tmpl_ws, FRAME_HEADER)
    if col_idx is None:
        raise ValueError(f"Header '{FRAME_HEADER}' not found in final template.")

    template_model_rows = build_template_model_index(tmpl_ws)

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
            logger.warning(
                f"Model '{model_val}' for frame not found in final template (Apple_Parts_rayan)"
            )
            continue

        tmpl_ws.cell(row=row_idx, column=col_idx).value = price_val

    tmpl_wb.save(FINAL_TEMPLATE_PATH)
    logger.info(
        f"Template updated with frame prices from {apple_xlsx.name}"
    )

# (You can add the rest of your fill_* functions here in same style if needed)


# ==================== HIGH-LEVEL PIPELINE ====================


def run_full_price_pipeline():
    """
    Run the full pipeline:
    - Convert all PDFs in PDF_FOLDER to Excel files
    - For each known Excel source (if exists), apply corresponding fill_* functions
    - Overwrite values in final template when data is repeated
    """
    logger.info("Starting full price pipeline...")

    if not TEMPLATE_IMPORT_XLSX.exists():
        raise FileNotFoundError(
            f"Template import XLSX not found: {TEMPLATE_IMPORT_XLSX}"
        )

    if not FINAL_TEMPLATE_PATH.exists():
        raise FileNotFoundError(
            f"Final template XLSX not found: {FINAL_TEMPLATE_PATH}"
        )

    # 1) Convert all PDFs
    convert_pdfs_to_excels()

    # 2) Fill from cell HIGH CAPACITY (if exists)
    cell_file = OUTPUT_FOLDER / "cell  HIGH CAPACITY.xlsx"
    if cell_file.exists():
        fill_template_from_converted_excel(cell_file)
    else:
        logger.info("cell  HIGH CAPACITY.xlsx not found, skipping battery prices.")

    # 3) JC PRODUCTS (if exists)
    if JC_PRODUCTS_PATH.exists():
        fill_template_from_jc_products(JC_PRODUCTS_PATH)
    else:
        logger.info("JC PRODUCTS NORMAL.xlsx not found, skipping battery tag column.")

    # 4) Apple parts normal (if exists)
    if APPLE_PARTS_NORMAL_PATH.exists():
        fill_template_from_apple_parts_normal(APPLE_PARTS_NORMAL_PATH)
        fill_template_from_apple_parts_normal_speakers(APPLE_PARTS_NORMAL_PATH)
        fill_template_from_apple_parts_normal_downSpeackers(APPLE_PARTS_NORMAL_PATH)
        fill_template_from_apple_parts_normal_flat_power(APPLE_PARTS_NORMAL_PATH)
        fill_template_from_apple_parts_normal_flat_power_volume(APPLE_PARTS_NORMAL_PATH)
        fill_template_from_apple_parts_normal_flat_flash_camera(APPLE_PARTS_NORMAL_PATH)
        fill_template_from_apple_parts_normal_vibration(APPLE_PARTS_NORMAL_PATH)
    else:
        logger.info("apple parts NORMAL (1).xlsx not found, skipping some parts.")

    # 5) Apple parts rayan (if exists)
    if APPLE_PARTS_RAYAN_PATH.exists():
        fill_template_from_apple_parts_rayan_frame(APPLE_PARTS_RAYAN_PATH)
    else:
        logger.info("Apple_Parts_rayan.xlsx not found, skipping rayan frame.")

    if not FINAL_TEMPLATE_PATH.exists():
        raise FileNotFoundError("Final template was not found after processing.")

    logger.info("Price pipeline finished successfully.")


# ====================== TELEGRAM BOT PART ======================


async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "سلام 👋\n"
        "Please send one or more PDF price files.\n"
        "I will convert them to Excel, update the final template, and send it back to you."
    )

async def handle_document(update: Update, context: ContextTypes.DEFAULT_TYPE):
    message = update.message
    doc = message.document

    if not doc:
        return

    # Accept only PDFs
    if doc.mime_type != "application/pdf" and not doc.file_name.lower().endswith(".pdf"):
        await message.reply_text("Please send PDF files only.")
        return

    PDF_FOLDER.mkdir(parents=True, exist_ok=True)

    filename = doc.file_name or f"{doc.file_unique_id}.pdf"
    pdf_path = PDF_FOLDER / filename

    await message.reply_text(
        f"PDF file '{filename}' received.\nProcessing, please wait..."
    )

    try:
        # ✅ IMPORTANT: get File object first, then download
        tg_file = await doc.get_file()
        await tg_file.download_to_drive(str(pdf_path))
        logger.info(f"Downloaded PDF to {pdf_path}")

        # Run full pipeline (all PDFs in folder, overwrite template values)
        run_full_price_pipeline()

        # Send final template back
        if FINAL_TEMPLATE_PATH.exists():
            with open(FINAL_TEMPLATE_PATH, "rb") as f:
                await message.reply_document(
                    document=f,
                    filename=FINAL_TEMPLATE_PATH.name,
                    caption="Final Excel template with updated prices is ready ✅",
                )
        else:
            await message.reply_text(
                "Error: Final template file was not found after processing."
            )

    except FileNotFoundError as e:
        logger.exception("File not found error in pipeline")
        await message.reply_text(
            f"File error during processing:\n{e}\n\n"
            "Please check that all required Excel templates and import files exist on the server."
        )
    except Exception as e:
        logger.exception("Unexpected error in pipeline")
        await message.reply_text(
            "An unexpected error occurred during processing.\n"
            f"Details: {e}"
        )


def main():
    application = ApplicationBuilder().token(BOT_TOKEN).build()

    application.add_handler(CommandHandler("start", start))
    application.add_handler(MessageHandler(filters.Document.ALL, handle_document))

    logger.info("Bot is starting...")
    application.run_polling()


if __name__ == "__main__":
    main()
