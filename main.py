from datetime import datetime, timedelta
import re
import sys
import pandas as pd

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

# Mã lô + (ngày sản xuất - ngày hết hạn), ví dụ: 092(20/05/25-20/08/26)
BATCH_PATTERN = re.compile(r'(\S+)\((\d{2}/\d{2}/\d{2})-(\d{2}/\d{2}/\d{2})\)')

FORECAST_DAYS = 15


def calc_for_date(manufacture_date, expiry_date, ref_date) -> dict:
    total_days = (expiry_date - manufacture_date).days
    remaining_days = (expiry_date - ref_date).days
    ratio = remaining_days / total_days if total_days > 0 else 0

    return {
        "total_days": total_days,
        "remaining_days": remaining_days,
        "ratio_percent": round(ratio * 100, 2),
    }


def build_extra_columns(ref_date: datetime, forecast_days: int) -> list[str]:
    columns = ["Ngày SX", "HSD", "Số ngày trong hạn"]
    for day_offset in range(forecast_days):
        date_str = (ref_date + timedelta(days=day_offset)).strftime("%d/%m/%y")
        columns.append(f"Còn hạn {date_str}")
        columns.append(f"Tỷ lệ {date_str} (%)")
    return columns


def enrich_excel(
    filepath: str,
    output_path: str | None = None,
    input_date: datetime | None = None,
    forecast_days: int = FORECAST_DAYS,
) -> pd.DataFrame:
    if output_path is None:
        output_path = filepath

    if input_date is None:
        input_date = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)

    df = pd.read_excel(filepath, header=None)
    extra_columns = build_extra_columns(input_date, forecast_days)
    total_columns = 1 + len(extra_columns)

    while df.shape[1] < total_columns:
        df[df.shape[1]] = None

    for col_idx, header in enumerate(extra_columns, start=1):
        df.iloc[0, col_idx] = header

    batch_count = 0
    for row_idx in range(len(df)):
        value = df.iloc[row_idx, 0]
        if pd.isna(value):
            continue

        match = BATCH_PATTERN.search(str(value).strip())
        if not match:
            continue

        manufacture_str = match.group(2)
        expiry_str = match.group(3)
        manufacture_date = datetime.strptime(manufacture_str, "%d/%m/%y")
        expiry_date = datetime.strptime(expiry_str, "%d/%m/%y")
        total_days = (expiry_date - manufacture_date).days

        df.iloc[row_idx, 1] = manufacture_str
        df.iloc[row_idx, 2] = expiry_str
        df.iloc[row_idx, 3] = total_days

        col_idx = 4
        for day_offset in range(forecast_days):
            ref_date = input_date + timedelta(days=day_offset)
            info = calc_for_date(manufacture_date, expiry_date, ref_date)
            df.iloc[row_idx, col_idx] = info["remaining_days"]
            df.iloc[row_idx, col_idx + 1] = info["ratio_percent"]
            col_idx += 2

        batch_count += 1

    df.to_excel(output_path, index=False, header=False)
    return df, batch_count, input_date


def process_file(filepath: str):
    df, batch_count, ref_date = enrich_excel(filepath, filepath)

    print(f"\n--- {filepath} ---")
    print(f"Ngày tham chiếu: {ref_date.strftime('%d/%m/%y')}")
    print(f"Đã ghi {batch_count} dòng mã lô")
    print(f"Số cột mới: {len(build_extra_columns(ref_date, FORECAST_DAYS))}")
    print("Mẫu 3 dòng đầu:")
    sample = df[df.iloc[:, 0].astype(str).str.contains(r"\(\d{2}/\d{2}/\d{2}-\d{2}/\d{2}/\d{2}\)", na=False)].head(3)
    print(sample.to_string())


if __name__ == "__main__":
    FILES = ["HCM.xlsx"]

    for file in FILES:
        process_file(file)
