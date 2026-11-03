
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


# ==========================================
# TEST 4: REPEATED EXECUTION
# ==========================================

def test_writer_is_idempotent(tmp_path):

    original = tmp_path / "original.xlsx"
    first_output = tmp_path / "first_output.xlsx"
    second_output = tmp_path / "second_output.xlsx"

    make_workbook(original)

    incoming = [
        fake_sale("TEST001"),
        fake_sale("TEST002"),
    ]

    # First execution
    first_result = write_synthetic_copy(
        original,
        first_output,
        incoming,
    )

    assert first_result["matched"] == 1
    assert first_result["inserted"] == 1

    # Second execution using the first output
    second_result = write_synthetic_copy(
        first_output,
        second_output,
        incoming,
    )

    assert second_result["matched"] == 2
    assert second_result["inserted"] == 0
    assert second_result["existing_after"] == 2

    # Check transaction cells in second output
    wb = load_workbook(second_output)

    try:
        ws = wb["SalesRepItem"]

        assert ws["Z5"].value == "TEST001"
        assert ws["Z6"].value == "TEST002"
        assert ws["Z7"].value is None

    finally:
        wb.close()


# ==========================================
# TEST 5: REPEATED IDENTICAL TRANSACTIONS
# ==========================================

def test_writer_preserves_duplicate_occurrences(tmp_path):

    original = tmp_path / "original.xlsx"
    output = tmp_path / "updated.xlsx"

    make_workbook(original)

    # Original workbook contains one TEST001.
    # Incoming data contains three TEST001 lines.
    incoming = [
        fake_sale("TEST001"),
        fake_sale("TEST001"),
        fake_sale("TEST001"),
    ]

    result = write_synthetic_copy(
        original,
        output,
        incoming,
    )

    assert result["matched"] == 1
    assert result["inserted"] == 2
    assert result["existing_after"] == 3

    wb = load_workbook(output)

    try:
        ws = wb["SalesRepItem"]

        assert ws["Z5"].value == "TEST001"
        assert ws["Z6"].value == "TEST001"
        assert ws["Z7"].value == "TEST001"

    finally:
        wb.close()


# ==========================================
# TEST 6: SOURCE AND OUTPUT MUST DIFFER
# ==========================================

def test_writer_rejects_same_input_output(tmp_path):

    original = tmp_path / "original.xlsx"

    make_workbook(original)

    original_hash = sha256(original)

    with pytest.raises(ValueError):
        write_synthetic_copy(
            original,
            original,
            [fake_sale("TEST002")],
        )

    # Original workbook must remain unchanged
    assert sha256(original) == original_hash


# ==========================================
# LESSON 32: TRANSACTION VALIDATION TESTS
# ==========================================

@pytest.mark.parametrize(
    "field,invalid_value",
    [
        ("bill_no", "   "),
        ("product", ""),
        ("sold_qty", 0),
        ("sales_amount", "not-a-number"),
        ("date", "not-a-date"),
    ],
)
def test_writer_rejects_invalid_transactions(
    tmp_path,
    field,
    invalid_value,
):
    original = tmp_path / "original.xlsx"
    output = tmp_path / "updated.xlsx"

    make_workbook(original)

    original_hash = sha256(original)

    invalid_sale = fake_sale("TEST002")

    # Change one field to an invalid value.
    invalid_sale[field] = invalid_value

    incoming = [
        fake_sale("TEST001"),
        invalid_sale,
    ]

    with pytest.raises(ValueError):
        write_synthetic_copy(
            original,
            output,
            incoming,
        )

    # Invalid data must not create an output.
    assert not output.exists()

    # Original workbook must remain unchanged.
    assert sha256(original) == original_hash


# ==========================================
# LESSON 35: FINANCIAL RECONCILIATION TESTS
# ==========================================

def test_writer_reconciles_saved_workbook(tmp_path):

    source = tmp_path / "original.xlsx"
    output = tmp_path / "updated.xlsx"

    make_workbook(source)

    original_hash = sha256(source)

    result = write_synthetic_copy(
        source,
        output,
        [
            fake_sale("TEST001"),
            fake_sale("TEST002"),
        ],
    )

    assert result["reconciled"] is True
    assert result["inserted"] == 1

    from src.financial_validation import calculate_totals
    from src.synthetic_excel_writer import read_existing_records

    wb = load_workbook(
        output,
        read_only=True,
        data_only=False,
    )

    try:
        records = read_existing_records(
            wb["SalesRepItem"]
        )
    finally:
        wb.close()

    totals = calculate_totals(records)

    assert totals["sold_qty"] == 2
    assert totals["cost_amount"] == 2000
    assert totals["sales_amount"] == 3000

    # Source Excel file is unchanged.
    assert sha256(source) == original_hash


def test_reconciliation_failure_blocks_output(
    tmp_path,
    monkeypatch,
):

    source = tmp_path / "original.xlsx"
    output = tmp_path / "updated.xlsx"

    make_workbook(source)

    original_hash = sha256(source)

    def reject_reconciliation(*args, **kwargs):
        raise ValueError(
            "Simulated financial mismatch"
        )

    monkeypatch.setattr(
        "src.synthetic_excel_writer.reconcile_totals",
        reject_reconciliation,
    )

    with pytest.raises(
        ValueError,
        match="Simulated financial mismatch",
    ):
        write_synthetic_copy(
            source,
            output,
            [
                fake_sale("TEST001"),
                fake_sale("TEST002"),
            ],
        )

    assert not output.exists()
    assert sha256(source) == original_hash
