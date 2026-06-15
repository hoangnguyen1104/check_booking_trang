import pandas as pd

from ..utils import BOOKING_RESULT_FILE, get_product_code, parse_quantity
from .stock_allocation import allocate_from_batches


def update_booking_fulfillment(booking_file: str, available_batches: dict[str, list[dict]]) -> None:
    booking_df = pd.read_excel(booking_file, header=None, dtype=object)
    fulfillment_col = booking_df.shape[1]
    note_col = fulfillment_col + 1
    booking_df[fulfillment_col] = None
    booking_df[note_col] = None
    booking_df.iloc[0, fulfillment_col] = "Số lượng đáp ứng"
    booking_df.iloc[0, note_col] = "Ghi chú chọn lô"

    for row_idx in range(1, len(booking_df)):
        product_code = get_product_code(booking_df.iloc[row_idx, 0])
        if not product_code:
            continue

        order_quantity = parse_quantity(booking_df.iloc[row_idx, 1])
        _, fulfilled_quantity, _, note = allocate_from_batches(
            product_code,
            order_quantity,
            available_batches,
        )

        booking_df.iloc[row_idx, fulfillment_col] = fulfilled_quantity
        booking_df.iloc[row_idx, note_col] = note

    booking_df.to_excel(BOOKING_RESULT_FILE, index=False, header=False)
