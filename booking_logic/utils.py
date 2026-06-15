import re
from datetime import datetime
from pathlib import Path

import pandas as pd
from werkzeug.utils import secure_filename

BATCH_PATTERN = re.compile(r"(\S+)\((\d{2}/\d{2}/\d{2})-(\d{2}/\d{2}/\d{2})\)")
EXTRA_COLUMNS = [
    "Ngày SX",
    "HSD",
    "Số ngày trong hạn",
    "Số ngày còn hạn",
    "Tỷ lệ (%)",
]

BASE_DIR = Path(__file__).resolve().parent.parent
UPLOAD_FOLDER = BASE_DIR / "uploads"
OUTPUT_FOLDER = BASE_DIR / "outputs"
RESULT_FILE = OUTPUT_FOLDER / "result.txt"
BOOKING_RESULT_FILE = OUTPUT_FOLDER / "Booking_result.xlsx"
STOCK_RESULT_FILE = OUTPUT_FOLDER / "Stock_result.xlsx"

UPLOAD_FOLDER.mkdir(exist_ok=True)
OUTPUT_FOLDER.mkdir(exist_ok=True)


def parse_quantity(value) -> int:
    if pd.isna(value):
        return 0
    try:
        return int(float(value))
    except (ValueError, TypeError):
        return 0


def calc_expiry_info(manufacture_date: datetime, expiry_date: datetime, ref_date: datetime) -> dict:
    total_days = (expiry_date - manufacture_date).days
    remaining_days = (expiry_date - ref_date).days
    ratio_percent = round(remaining_days / total_days * 100, 2) if total_days > 0 else 0

    return {
        "total_days": total_days,
        "remaining_days": remaining_days,
        "ratio_percent": ratio_percent,
    }


def get_product_code(value) -> str | None:
    if pd.isna(value):
        return None

    text = str(value).strip()
    bracket_match = re.search(r"\[(.*?)\]", text)
    if bracket_match:
        return bracket_match.group(1).strip()

    return text if text else None


def display_upload_name(filepath: str) -> str:
    filename = Path(filepath).name
    for prefix in ("stock_", "booking_"):
        if filename.startswith(prefix):
            return filename[len(prefix):]
    return filename


def save_upload(file_storage, prefix: str) -> str:
    filename = secure_filename(file_storage.filename)
    if not filename:
        raise ValueError("Tên file không hợp lệ")

    saved_path = UPLOAD_FOLDER / f"{prefix}_{filename}"
    file_storage.save(saved_path)
    return str(saved_path)
