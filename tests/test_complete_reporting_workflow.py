
# ==========================================
# LESSON 50
# COMPLETE REPORTING WORKFLOW TESTS
# ==========================================

import json

import pytest

from openpyxl import load_workbook

from src.audit_logging import calculate_file_hash

from src.audit_verification import (
    verify_workbook_audit,
)

from src.synthetic_excel_writer import (
    write_synthetic_copy,
)

from tests.test_synthetic_excel_writer import (
    make_workbook,
    fake_sale,
    sha256,
)


# ==========================================
# TEST 1: COMPLETE SUCCESSFUL WORKFLOW
# ==========================================

def test_complete_reporting_workflow(tmp_path):

    source = tmp_path / "original.xlsx"
    output = tmp_path / "updated.xlsx"
    audit = tmp_path / "audit.json"

    make_workbook(source)

    original_hash = sha256(source)

    # Existing test workbook contains TEST001.
    incoming = [
        fake_sale("TEST001"),
        fake_sale("TEST002"),
    ]

    result = write_synthetic_copy(
        source_path=source,
        output_path=output,
        incoming_records=incoming,
        audit_path=audit,
        run_id="TEST-COMPLETE-001",
    )

    # Verify writer output.
    assert output.exists()
    assert audit.exists()

    assert result["matched"] == 1
    assert result["inserted"] == 1
    assert result["reconciled"] is True

    # Verify audit and workbook together.
    verification = verify_workbook_audit(
        output,
        audit,
    )

    assert verification["verified"] is True
    assert verification["workbook_verified"] is True
    assert verification["synthetic_marker_verified"] is True

    # Verify audit record.
    record = json.loads(
        audit.read_text(encoding="utf-8")
    )

    assert record["run_id"] == "TEST-COMPLETE-001"
    assert record["existing_count"] == 1
    assert record["inserted_count"] == 1
    assert record["final_count"] == 2

    assert record["output_sha256"] == (
        calculate_file_hash(output)
    )

    # Confirm output contains both transactions.
    workbook = load_workbook(
        output,
        read_only=True,
    )

    try:
        worksheet = workbook["SalesRepItem"]

        assert worksheet["A1"].value == (
            "SYNTHETIC_TEST_ONLY"
        )

        bill_numbers = [
            row[0]
            for row in worksheet.iter_rows(
                min_row=5,
                min_col=26,
                max_col=26,
                values_only=True,
            )
            if row[0] is not None
        ]

        assert bill_numbers.count("TEST001") == 1
        assert bill_numbers.count("TEST002") == 1

    finally:
        workbook.close()

    # Original must remain unchanged.
    assert sha256(source) == original_hash


# ==========================================
# TEST 2: DETECT CHANGED EXCEL AFTER AUDIT
# ==========================================

def test_modified_output_fails_verification(
    tmp_path,
):

    source = tmp_path / "original.xlsx"
    output = tmp_path / "updated.xlsx"
    audit = tmp_path / "audit.json"

    make_workbook(source)

    write_synthetic_copy(
        source_path=source,
        output_path=output,
        incoming_records=[
            fake_sale("TEST001"),
            fake_sale("TEST002"),
        ],
        audit_path=audit,
        run_id="TEST-COMPLETE-002",
    )

    assert output.exists()
    assert audit.exists()

    # Simulate modifying the published report.
    workbook = load_workbook(output)

    try:
        worksheet = workbook["SalesRepItem"]
        worksheet["Z5"] = "MODIFIED_BILL"

        workbook.save(output)

    finally:
        workbook.close()

    # Original audit must reject changed output.
    with pytest.raises(
        ValueError,
        match="checksum does not match",
    ):
        verify_workbook_audit(
            output,
            audit,
        )


# ==========================================
# LESSON 51
# AUDIT VS ACTUAL WORKBOOK COUNT TESTS
# ==========================================


# TEST 3: AUDIT CLAIMS TOO MANY TRANSACTIONS

def test_audit_count_must_match_excel(tmp_path):

    source = tmp_path / "original.xlsx"
    output = tmp_path / "updated.xlsx"
    audit = tmp_path / "audit.json"

    make_workbook(source)

    write_synthetic_copy(
        source_path=source,
        output_path=output,
        incoming_records=[
            fake_sale("TEST001"),
            fake_sale("TEST002"),
        ],
        audit_path=audit,
        run_id="TEST-COUNT-001",
    )

    # Workbook contains 2 transactions.
    # Change the audit to claim 10.
    record = json.loads(
        audit.read_text(encoding="utf-8")
    )

    record["existing_count"] = 8
    record["inserted_count"] = 2
    record["final_count"] = 10

    audit.write_text(
        json.dumps(record),
        encoding="utf-8",
    )

    with pytest.raises(ValueError):
        verify_workbook_audit(output, audit)


# TEST 4: WORKBOOK TRANSACTION DISAPPEARS

def test_missing_excel_transaction_detected(
    tmp_path,
):

    source = tmp_path / "original.xlsx"
    output = tmp_path / "updated.xlsx"
    audit = tmp_path / "audit.json"

    make_workbook(source)

    write_synthetic_copy(
        source_path=source,
        output_path=output,
        incoming_records=[
            fake_sale("TEST001"),
            fake_sale("TEST002"),
        ],
        audit_path=audit,
        run_id="TEST-COUNT-002",
    )

    # Remove the second transaction.
    workbook = load_workbook(output)

    try:
        worksheet = workbook["SalesRepItem"]
        worksheet.delete_rows(6, 1)
        workbook.save(output)
    finally:
        workbook.close()

    # Update the checksum to match the
    # changed file, but leave counts unchanged.
    record = json.loads(
        audit.read_text(encoding="utf-8")
    )

    record["output_sha256"] = (
        calculate_file_hash(output)
    )

    audit.write_text(
        json.dumps(record),
        encoding="utf-8",
    )

    # The file hash now matches, but there
    # should be only 1 transaction instead of 2.
    with pytest.raises(ValueError):
        verify_workbook_audit(output, audit)

# ==========================================
# LESSON 53
# INCOMPLETE TRANSACTION DETECTION
# ==========================================


# TEST 5: DETECT MISSING PRODUCT

def test_missing_product_is_rejected(tmp_path):

    source = tmp_path / "original.xlsx"
    output = tmp_path / "updated.xlsx"
    audit = tmp_path / "audit.json"

    make_workbook(source)

    write_synthetic_copy(
        source_path=source,
        output_path=output,
        incoming_records=[
            fake_sale("TEST001"),
            fake_sale("TEST002"),
        ],
        audit_path=audit,
        run_id="TEST-QUALITY-001",
    )

    workbook = load_workbook(output)

    try:
        worksheet = workbook["SalesRepItem"]

        # AO = Product column.
        # Remove the second transaction's product.
        worksheet["AO6"] = None

        workbook.save(output)

    finally:
        workbook.close()

    # Update checksum to isolate
    # transaction-content validation.
    record = json.loads(
        audit.read_text(encoding="utf-8")
    )

    record["output_sha256"] = (
        calculate_file_hash(output)
    )

    audit.write_text(
        json.dumps(record),
        encoding="utf-8",
    )

    # Transaction count still equals 2,
    # but a required product is missing.
    with pytest.raises(ValueError):
        verify_workbook_audit(output, audit)


# TEST 6: DETECT MISSING SOLD QUANTITY

def test_missing_sold_quantity_is_rejected(
    tmp_path,
):

    source = tmp_path / "original.xlsx"
    output = tmp_path / "updated.xlsx"
    audit = tmp_path / "audit.json"

    make_workbook(source)

    write_synthetic_copy(
        source_path=source,
        output_path=output,
        incoming_records=[
            fake_sale("TEST001"),
            fake_sale("TEST002"),
        ],
        audit_path=audit,
        run_id="TEST-QUALITY-002",
    )

    workbook = load_workbook(output)

    try:
        worksheet = workbook["SalesRepItem"]

        # AT = Sold Quantity column.
        worksheet["AT6"] = None

        workbook.save(output)

    finally:
        workbook.close()

    # Update checksum so the audit hash
    # matches the modified workbook.
    record = json.loads(
        audit.read_text(encoding="utf-8")
    )

    record["output_sha256"] = (
        calculate_file_hash(output)
    )

    audit.write_text(
        json.dumps(record),
        encoding="utf-8",
    )

    # Both Bill Numbers still exist.
    # However, quantity is missing.
    with pytest.raises(ValueError):
        verify_workbook_audit(output, audit)
