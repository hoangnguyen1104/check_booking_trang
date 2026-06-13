from datetime import datetime
import re
import sys
import pandas as pd

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

# Mã lô + (ngày sản xuất - ngày hết hạn), ví dụ: 092(20/05/25-20/08/26)
BATCH_PATTERN = re.compile(r'(\S+)\((\d{2}/\d{2}/\d{2})-(\d{2}/\d{2}/\d{2})\)')

REF_DATE = datetime(2026, 6, 16)

EXTRA_COLUMNS = [
    "Ngày SX",
    "HSD",
    "Số ngày trong hạn",
    "Số ngày còn hạn",
    "Tỷ lệ (%)",
]


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


def enrich_excel(filepath: str, output_path: str | None = None) -> tuple[pd.DataFrame, int]:
    if output_path is None:
        output_path = filepath

    df = pd.read_excel(filepath, header=None, dtype=object)
    total_columns = 2 + len(EXTRA_COLUMNS)

    # Khôi phục tiêu đề cột B nếu đã bị ghi
    df.iloc[1, 1] = "HCM/Stock/MT"
    df.iloc[2, 1] = "Số lượng"

    while df.shape[1] < total_columns:
        df[df.shape[1]] = None

    # Tiêu đề cột (dòng 2, cùng hàng với "Số lượng")
    for col_idx, header in enumerate(EXTRA_COLUMNS, start=2):
        df.iloc[2, col_idx] = header

    batch_count = 0
    for row_idx in range(len(df)):
        # Chuẩn hóa số lượng từ dòng dữ liệu (bỏ qua 3 dòng tiêu đề)
        if row_idx >= 3:
            df.iloc[row_idx, 1] = parse_quantity(df.iloc[row_idx, 1])

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
        info = calc_expiry_info(manufacture_date, expiry_date, REF_DATE)

        df.iloc[row_idx, 2] = manufacture_str
        df.iloc[row_idx, 3] = expiry_str
        df.iloc[row_idx, 4] = info["total_days"]
        df.iloc[row_idx, 5] = info["remaining_days"]
        df.iloc[row_idx, 6] = info["ratio_percent"]
        batch_count += 1

    df.to_excel(output_path, index=False, header=False)
    return df, batch_count


def process_file(filepath: str):
    df, batch_count = enrich_excel(filepath, filepath)

    print(f"\n--- {filepath} ---")
    print(f"Ngày tham chiếu: {REF_DATE.strftime('%d/%m/%y')}")
    print(f"Đã ghi {batch_count} dòng mã lô")
    print(f"Cột mới: {', '.join(EXTRA_COLUMNS)}")
    print("Mẫu 5 dòng mã lô đầu:")
    sample = df[df.iloc[:, 0].astype(str).str.contains(
        r"\(\d{2}/\d{2}/\d{2}-\d{2}/\d{2}/\d{2}\)", na=False
    )].head(5)
    print(sample.to_string())


def read_booking(filepath: str):
    product_map = {}

    df = pd.read_excel(filepath, header=None, dtype=object)

    for product_code, amount in df.iloc[1:, [0, 1]].itertuples(index=False):
        product_map[product_code] = product_map.get(product_code, 0) + amount

    return product_map

if __name__ == "__main__":
    process_file("HCM_1306.xlsx")
    df = pd.read_excel("HCM_1306.xlsx", header=None, dtype=object)
    p_orders = read_booking("Booking.xlsx")
    for key, sl in p_orders.items():
        print(f"{key}: {sl}")
        for row_idx in range(len(df)):
            value = df.iloc[row_idx, 0]
            if str(key) in str(value):
                print(value)
                p_remain = 0
                p_date = False
                for i in range(row_idx+1, len(df)):
                    value = df.iloc[i, 0]
                    match = BATCH_PATTERN.search(str(value).strip())
                    if not match:
                        break
                    row_values = df.iloc[i].tolist()
                    if float(row_values[6]) >= 50:
                        print(row_values[0], "Tồn kho:", row_values[1], "  Date:", row_values[6])
                        p_date = True
                        p_remain += int(row_values[1])
                if not p_date:
                    print("Không có lô nào còn đủ date")
                if p_remain < sl:
                    print("Không còn đủ tồn kho", p_remain)
                break
        print("\n\n")
