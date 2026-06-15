from datetime import datetime

from ..utils import BATCH_PATTERN, EXTRA_COLUMNS, display_upload_name, parse_quantity


def build_legacy_result_text(
    stock_file: str,
    booking_file: str,
    stock_df,
    batch_count: int,
    ref_date: datetime,
    threshold_percent: float,
) -> str:
    lines = [
        f"--- {display_upload_name(stock_file)} ---",
        f"Ngày tham chiếu: {ref_date.strftime('%d/%m/%y')}",
        f"Ngưỡng date: {threshold_percent:g}%",
        f"Đã ghi {batch_count} dòng mã lô",
        f"Cột mới: {', '.join(EXTRA_COLUMNS)}",
        "Mẫu 5 dòng mã lô đầu:",
    ]

    sample = stock_df[
        stock_df.iloc[:, 0].astype(str).str.contains(
            r"\(\d{2}/\d{2}/\d{2}-\d{2}/\d{2}/\d{2}\)",
            na=False,
            regex=True,
        )
    ].head(5)
    lines.append(sample.to_string())

    p_orders = read_booking_totals(booking_file)
    for key, sl in p_orders.items():
        lines.append(f"{key}: {sl}")

        for row_idx in range(len(stock_df)):
            value = stock_df.iloc[row_idx, 0]
            if str(key) not in str(value):
                continue

            lines.append(str(value))
            p_remain = 0
            p_date = False

            for i in range(row_idx + 1, len(stock_df)):
                batch_value = stock_df.iloc[i, 0]
                match = BATCH_PATTERN.search(str(batch_value).strip())
                if not match:
                    break

                row_values = stock_df.iloc[i].tolist()
                ratio = float(row_values[6] or 0)
                if ratio >= threshold_percent:
                    lines.append(f"{row_values[0]} Tồn kho: {row_values[1]}   Date: {row_values[6]}")
                    p_date = True
                    p_remain += parse_quantity(row_values[1])

            if not p_date:
                lines.append("Không có lô nào còn đủ date")
            if p_remain < sl:
                lines.append(f"Không còn đủ tồn kho {p_remain}")
            break

        lines.extend(["", ""])

    return "\n".join(lines)


def read_booking_totals(filepath: str) -> dict:
    import pandas as pd

    product_map = {}
    df = pd.read_excel(filepath, header=None, dtype=object)

    for product_code, amount in df.iloc[1:, [0, 1]].itertuples(index=False):
        if pd.isna(product_code):
            continue
        product_map[product_code] = product_map.get(product_code, 0) + parse_quantity(amount)

    return product_map
