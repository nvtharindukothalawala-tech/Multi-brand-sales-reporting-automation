
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
