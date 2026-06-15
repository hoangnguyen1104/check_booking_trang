from __future__ import annotations

from dataclasses import dataclass


@dataclass
class BookingOrder:
    product_code: str
    order_quantity: int
    fulfilled_quantity: int = 0
    notes: str = ""
