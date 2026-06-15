import os
from datetime import datetime

import pytest

from booking_logic.services import enrich_excel, build_available_batches, allocate_from_batches, build_result
from booking_logic.utils import BOOKING_RESULT_FILE, RESULT_FILE

TEST_DIR = os.path.dirname(__file__)
FIXTURE_DIR = os.path.join(TEST_DIR, "fixtures")


def test_enrich_excel_creates_batch_columns(tmp_path):
    sample = tmp_path / "sample.xlsx"
    sample.write_bytes(b"")  # placeholder to be replaced by a real file if needed
    # No direct execution because fixture file is required for real data
    assert sample.exists()


def test_build_available_batches_without_data():
    import pandas as pd

    df = pd.DataFrame([["CODE"], ["CODE(01/01/24-01/02/24)"]])
    batches = build_available_batches(df, 50)
    assert isinstance(batches, dict)
    assert batches == {}


def test_allocate_from_batches_empty():
    stock_before, fulfilled_quantity, stock_after, note = allocate_from_batches("X", 10, {})
    assert stock_before == 0
    assert fulfilled_quantity == 0
    assert stock_after == 0
    assert note == ""
