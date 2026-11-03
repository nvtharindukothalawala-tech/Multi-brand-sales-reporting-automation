
import hashlib

import pytest
from openpyxl import Workbook, load_workbook

from src.synthetic_excel_writer import (
    TEST_MARKER,
    write_synthetic_copy,
)


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


def make_workbook(path, marker=TEST_MARKER):
    wb = Workbook()
    ws = wb.active
    ws.title = "SalesRepItem"

    ws["A1"] = marker
    ws["B5"] = "=1+1"

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


def sha256(path):
    return hashlib.sha256(
        path.read_bytes()
    ).hexdigest()


def test_writer_adds_only_new_rows(tmp_path):
    source = tmp_path / "original.xlsx"
    output = tmp_path / "updated.xlsx"

    make_workbook(source)

    before_hash = sha256(source)

    incoming = [
        fake_sale("TEST001"),
        fake_sale("TEST002"),
    ]

    result = write_synthetic_copy(
        source, output, incoming
    )

    assert result["matched"] == 1
    assert result["inserted"] == 1
    assert result["existing_after"] == 2

    # Original workbook remains unchanged.
    assert sha256(source) == before_hash

    wb = load_workbook(output)

    try:
        ws = wb["SalesRepItem"]

        assert ws["Z5"].value == "TEST001"
        assert ws["Z6"].value == "TEST002"

        # Original formula is preserved.
        assert ws["B5"].value == "=1+1"

    finally:
        wb.close()


def test_writer_blocks_real_workbook(tmp_path):
    source = tmp_path / "unmarked.xlsx"
    output = tmp_path / "updated.xlsx"

    make_workbook(
        source,
        marker="COMPANY_REPORT",
    )

    before_hash = sha256(source)

    with pytest.raises(ValueError):
        write_synthetic_copy(
            source,
            output,
            [fake_sale("TEST002")],
        )

    assert not output.exists()
    assert sha256(source) == before_hash


def test_writer_rejects_overwrite(tmp_path):
    source = tmp_path / "original.xlsx"
    output = tmp_path / "updated.xlsx"

    make_workbook(source)
    output.write_bytes(b"existing file")

    before = sha256(output)

    with pytest.raises(FileExistsError):
        write_synthetic_copy(
            source,
            output,
            [fake_sale("TEST002")],
        )

    assert sha256(output) == before
