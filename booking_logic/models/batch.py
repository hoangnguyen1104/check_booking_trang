from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime


@dataclass
class BatchInfo:
    raw_value: str
    quantity: int
    manufacture_date: datetime
    expiry_date: datetime
    total_days: int
    remaining_days: int
    ratio_percent: float
    row_order: int
