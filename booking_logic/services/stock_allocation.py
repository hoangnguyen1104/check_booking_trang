from ..utils import BATCH_PATTERN, get_product_code, parse_quantity


def build_available_batches(stock_df, threshold_percent: float) -> dict[str, list[dict]]:
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
