import re
from datetime import datetime

import pandas as pd

from .utils import BATCH_PATTERN, EXTRA_COLUMNS, calc_expiry_info, get_product_code, parse_quantity


def enrich_excel(filepath: str, ref_date: datetime) -> tuple[pd.DataFrame, int]:
    df = pd.read_excel(filepath, header=None, dtype=object)
    total_columns = 2 + len(EXTRA_COLUMNS)

    df.iloc[1, 1] = "HCM/Stock/MT"
    df.iloc[2, 1] = "Số lượng"

    while df.shape[1] < total_columns:
        df[df.shape[1]] = None

    for col_idx, header in enumerate(EXTRA_COLUMNS, start=2):
        df.iloc[2, col_idx] = header

    batch_count = 0
    for row_idx in range(len(df)):
        if row_idx >= 3:
            df.iloc[row_idx, 1] = parse_quantity(df.iloc[row_idx, 1])

        value = df.iloc[row_idx, 0]
        if pd.isna(value):
            continue

        match = BATCH_PATTERN.search(str(value).strip())
        if not match:
            continue

        manufacture_date = datetime.strptime(match.group(2), "%d/%m/%y")
        expiry_date = datetime.strptime(match.group(3), "%d/%m/%y")
        info = calc_expiry_info(manufacture_date, expiry_date, ref_date)

        df.iloc[row_idx, 2] = match.group(2)
        df.iloc[row_idx, 3] = match.group(3)
        df.iloc[row_idx, 4] = info["total_days"]
        df.iloc[row_idx, 5] = info["remaining_days"]
        df.iloc[row_idx, 6] = info["ratio_percent"]
        batch_count += 1

    return df, batch_count


def build_available_batches(stock_df: pd.DataFrame, threshold_percent: float) -> dict[str, list[dict]]:
    available_batches = {}

    for row_idx in range(len(stock_df)):
        product_code = get_product_code(stock_df.iloc[row_idx, 0])
        if not product_code or BATCH_PATTERN.search(str(stock_df.iloc[row_idx, 0]).strip()):
            continue

        available_batches.setdefault(product_code, [])

        for batch_idx in range(row_idx + 1, len(stock_df)):
            batch_value = stock_df.iloc[batch_idx, 0]
            match = BATCH_PATTERN.search(str(batch_value).strip())
            if not match:
                break

            row_values = stock_df.iloc[batch_idx].tolist()
            ratio = float(row_values[6] or 0)
            quantity = parse_quantity(row_values[1])

            if ratio >= threshold_percent and quantity > 0:
                available_batches[product_code].append(
                    {
                        "batch": str(row_values[0]),
                        "quantity": quantity,
                        "ratio": ratio,
                        "row_order": batch_idx,
                    }
                )

    for batches in available_batches.values():
        batches.sort(key=lambda batch: (batch["ratio"], batch["row_order"]))

    return available_batches


def total_batch_quantity(batches: list[dict]) -> int:
    return sum(batch["quantity"] for batch in batches)


def allocate_from_batches(product_code: str, order_quantity: int, available_batches: dict[str, list[dict]]) -> tuple[int, int, int, str]:
    batches = available_batches.get(product_code, [])
    stock_before = total_batch_quantity(batches)
    remaining_order = order_quantity
    fulfilled_quantity = 0
    note_lines = []

    for batch in batches:
        if remaining_order <= 0:
            break
        if batch["quantity"] <= 0:
            continue

        picked_quantity = min(remaining_order, batch["quantity"])
        batch["quantity"] -= picked_quantity
        remaining_order -= picked_quantity
        fulfilled_quantity += picked_quantity
        note_lines.append(f"{batch['batch']}: {picked_quantity} (Date {batch['ratio']}%)")

    stock_after = total_batch_quantity(batches)
    return stock_before, fulfilled_quantity, stock_after, "\n".join(note_lines)
