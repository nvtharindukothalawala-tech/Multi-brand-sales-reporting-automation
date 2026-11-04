
# ==========================================
# LESSON 48
# REAL SYNTHETIC WORKBOOK AUDIT TESTS
# ==========================================

import json

import pytest

from openpyxl import Workbook

from src.audit_logging import (
    calculate_file_hash,
    create_audit_record,
)

from src.audit_verification import (
    verify_workbook_audit,
)


# ==========================================
# HELPER: CREATE SYNTHETIC WORKBOOK
# ==========================================

def make_test_workbook(
    tmp_path,
    sheet_name="SalesRepItem",
    marker="SYNTHETIC_TEST_ONLY",
):

    excel = tmp_path / "synthetic.xlsx"

    wb = Workbook()
    ws = wb.active
    ws.title = sheet_name
    ws["A1"] = marker

    wb.save(excel)
    wb.close()

    return excel


# ==========================================
# HELPER: CREATE MATCHING AUDIT
# ==========================================

def make_test_audit(tmp_path, excel):

    audit = tmp_path / "audit.json"

    record = create_audit_record(
        run_id="TEST-WORKBOOK-001",
        existing_count=0,
        inserted_count=0,
        reconciled=True,
        output_hash=calculate_file_hash(excel),
    )

    audit.write_text(
        json.dumps(record),
        encoding="utf-8",
    )

    return audit


# ==========================================
# TEST 1: VALID SYNTHETIC WORKBOOK
# ==========================================

def test_valid_synthetic_workbook(tmp_path):

    excel = make_test_workbook(tmp_path)
    audit = make_test_audit(tmp_path, excel)

    result = verify_workbook_audit(
        excel,
        audit,
    )

    assert result["verified"] is True
    assert result["run_id"] == "TEST-WORKBOOK-001"


# ==========================================
# TEST 2: MISSING SYNTHETIC MARKER
# ==========================================

def test_missing_synthetic_marker_rejected(
    tmp_path,
):

    excel = make_test_workbook(
        tmp_path,
        marker="NOT_SYNTHETIC",
    )

    audit = make_test_audit(tmp_path, excel)

    with pytest.raises(ValueError):
        verify_workbook_audit(
            excel,
            audit,
        )


# ==========================================
# TEST 3: MISSING SALES SHEET
# ==========================================

def test_missing_sales_sheet_rejected(tmp_path):

    excel = make_test_workbook(
        tmp_path,
        sheet_name="WrongSheet",
    )

    audit = make_test_audit(tmp_path, excel)

    with pytest.raises(ValueError):
        verify_workbook_audit(
            excel,
            audit,
        )


# ==========================================
# TEST 4: INVALID XLSX FILE
# ==========================================

def test_invalid_xlsx_rejected(tmp_path):

    excel = tmp_path / "synthetic.xlsx"

    excel.write_bytes(
        b"THIS_IS_NOT_A_REAL_XLSX_FILE"
    )

    audit = make_test_audit(tmp_path, excel)

    with pytest.raises(ValueError):
        verify_workbook_audit(
            excel,
            audit,
        )
