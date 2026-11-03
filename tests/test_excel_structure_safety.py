
import hashlib

import pytest
from openpyxl import Workbook, load_workbook
from openpyxl.worksheet.table import Table, TableStyleInfo

from src.synthetic_excel_writer import (
    TEST_MARKER,
    write_synthetic_copy,
)


# ==========================================
# 1. FAKE TRANSACTION
# ==========================================

def fake_sale(bill_no):
    return {
        "bill_no": bill_no,
        "product": "TEST_SHIRT",
        "sold_qty": 1,
        "cost_price": 1000,
        "sales_price": 1500,
        "cost_amount": 1000,
        "sales_amount": 1500,
        "date": "2026-10-10",
    }


# ==========================================
# 2. CREATE SYNTHETIC WORKBOOK
# ==========================================

def create_test_workbook(path):

    wb = Workbook()
    ws = wb.active
    ws.title = "SalesRepItem"

    ws["A1"] = TEST_MARKER
    ws["B5"] = "=AT5*AQ5"

    record = fake_sale("TEST001")

    columns = {
        "bill_no": "Z",
        "product": "AO",
        "cost_price": "AP",
        "sales_price": "AQ",
        "sold_qty": "AT",
        "cost_amount": "AU",
        "sales_amount": "AV",
        "date": "AZ",
    }

    for field, column in columns.items():
        ws[f"{column}5"] = record[field]

    wb.save(path)
    wb.close()


def file_hash(path):
    return hashlib.sha256(
        path.read_bytes()
    ).hexdigest()


# ==========================================
# TEST 1: REJECT FOOTER ROWS
# ==========================================

def test_writer_rejects_footer_rows(tmp_path):

    original = tmp_path / "original.xlsx"
    output = tmp_path / "updated.xlsx"

    create_test_workbook(original)

    wb = load_workbook(original)

    try:
        ws = wb["SalesRepItem"]

        # A totals/footer row below transactions
        ws["Z6"] = "TOTAL"
        ws["AV6"] = "=SUM(AV5:AV5)"

        wb.save(original)

    finally:
        wb.close()

    before_hash = file_hash(original)

    with pytest.raises(ValueError):
        write_synthetic_copy(
            original,
            output,
            [
                fake_sale("TEST001"),
                fake_sale("TEST002"),
            ],
        )

    assert not output.exists()
    assert file_hash(original) == before_hash


# ==========================================
# TEST 2: REJECT STRUCTURED EXCEL TABLES
# ==========================================

def test_writer_rejects_excel_tables(tmp_path):

    original = tmp_path / "original.xlsx"
    output = tmp_path / "updated.xlsx"

    create_test_workbook(original)

    wb = load_workbook(original)

    try:
        ws = wb["SalesRepItem"]

        # Add a small structured table.
        ws["Z4"] = "Bill Number"
        ws["AA4"] = "Location"
        ws["AA5"] = "TEST_STORE"

        table = Table(
            displayName="TestSalesTable",
            ref="Z4:AA5",
        )

        table.tableStyleInfo = TableStyleInfo(
            name="TableStyleMedium2",
            showRowStripes=True,
        )

        ws.add_table(table)
        wb.save(original)

    finally:
        wb.close()

    before_hash = file_hash(original)

    with pytest.raises(ValueError):
        write_synthetic_copy(
            original,
            output,
            [
                fake_sale("TEST001"),
                fake_sale("TEST002"),
            ],
        )

    assert not output.exists()
    assert file_hash(original) == before_hash


# ==========================================
# TEST 3: REJECT MERGED TRANSACTION CELLS
# ==========================================

def test_writer_rejects_merged_cells(tmp_path):

    original = tmp_path / "original.xlsx"
    output = tmp_path / "updated.xlsx"

    create_test_workbook(original)

    wb = load_workbook(original)

    try:
        ws = wb["SalesRepItem"]

        # Merge cells within the sales row.
        ws.merge_cells("D5:E5")

        wb.save(original)

    finally:
        wb.close()

    before_hash = file_hash(original)

    with pytest.raises(ValueError):
        write_synthetic_copy(
            original,
            output,
            [
                fake_sale("TEST001"),
                fake_sale("TEST002"),
            ],
        )

    assert not output.exists()
    assert file_hash(original) == before_hash
