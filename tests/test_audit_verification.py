
# ==========================================
# LESSON 44 - AUDIT VERIFICATION TESTS
# ==========================================

import json

import pytest

from src.audit_logging import (
    calculate_file_hash,
    create_audit_record,
)

from src.audit_verification import (
    verify_audit_integrity,
)


# ==========================================
# HELPER: CREATE TEST FILES
# ==========================================

def create_test_files(tmp_path):

    excel = tmp_path / "synthetic.xlsx"
    audit = tmp_path / "audit.json"

    # Synthetic content for checksum testing.
    excel.write_bytes(
        b"SYNTHETIC_EXCEL_TEST_CONTENT"
    )

    record = create_audit_record(
        run_id="TEST-VERIFY-001",
        existing_count=10,
        inserted_count=2,
        reconciled=True,
        output_hash=calculate_file_hash(excel),
    )

    audit.write_text(
        json.dumps(record, indent=4),
        encoding="utf-8",
    )

    return excel, audit


# ==========================================
# TEST 1: VALID REPORT AND AUDIT
# ==========================================

def test_valid_audit_integrity(tmp_path):

    excel, audit = create_test_files(tmp_path)

    result = verify_audit_integrity(
        excel,
        audit,
    )

    assert result["verified"] is True
    assert result["run_id"] == "TEST-VERIFY-001"


# ==========================================
# TEST 2: MODIFIED REPORT DETECTED
# ==========================================

def test_modified_report_is_rejected(tmp_path):

    excel, audit = create_test_files(tmp_path)

    # Simulate an unauthorized file change.
    excel.write_bytes(
        b"MODIFIED_SYNTHETIC_CONTENT"
    )

    with pytest.raises(ValueError):
        verify_audit_integrity(
            excel,
            audit,
        )


# ==========================================
# TEST 3: INVALID AUDIT STATUS
# ==========================================

def test_failed_audit_status_is_rejected(tmp_path):

    excel, audit = create_test_files(tmp_path)

    record = json.loads(
        audit.read_text(encoding="utf-8")
    )

    record["status"] = "FAILED"
    record["reconciled"] = False

    audit.write_text(
        json.dumps(record),
        encoding="utf-8",
    )

    with pytest.raises(ValueError):
        verify_audit_integrity(
            excel,
            audit,
        )


# ==========================================
# TEST 4: MISSING AUDIT FILE
# ==========================================

def test_missing_audit_file_is_rejected(tmp_path):

    excel, audit = create_test_files(tmp_path)

    audit.unlink()

    with pytest.raises(FileNotFoundError):
        verify_audit_integrity(
            excel,
            audit,
        )


# ==========================================
# LESSON 46 - ADVANCED AUDIT VALIDATION
# ==========================================


# TEST 5: CORRUPTED JSON

def test_corrupted_audit_json_is_rejected(tmp_path):

    excel, audit = create_test_files(tmp_path)

    audit.write_text(
        "{invalid_json",
        encoding="utf-8",
    )

    with pytest.raises(ValueError):
        verify_audit_integrity(excel, audit)


# TEST 6: MISSING CHECKSUM

def test_missing_checksum_is_rejected(tmp_path):

    excel, audit = create_test_files(tmp_path)

    record = json.loads(
        audit.read_text(encoding="utf-8")
    )

    record.pop("output_sha256")

    audit.write_text(
        json.dumps(record),
        encoding="utf-8",
    )

    with pytest.raises(ValueError):
        verify_audit_integrity(excel, audit)


# TEST 7: INCORRECT REPORT TYPE

def test_wrong_report_type_is_rejected(tmp_path):

    excel, audit = create_test_files(tmp_path)

    record = json.loads(
        audit.read_text(encoding="utf-8")
    )

    record["report_type"] = "PRODUCTION"

    audit.write_text(
        json.dumps(record),
        encoding="utf-8",
    )

    with pytest.raises(ValueError):
        verify_audit_integrity(excel, audit)


# TEST 8: INCONSISTENT TRANSACTION COUNTS

def test_inconsistent_transaction_counts_rejected(
    tmp_path,
):
    excel, audit = create_test_files(tmp_path)

    record = json.loads(
        audit.read_text(encoding="utf-8")
    )

    # Existing 10 + Inserted 2 should equal 12.
    record["final_count"] = 999

    audit.write_text(
        json.dumps(record),
        encoding="utf-8",
    )

    with pytest.raises(ValueError):
        verify_audit_integrity(excel, audit)


# TEST 9: INVALID NEGATIVE COUNTS

def test_negative_transaction_count_rejected(
    tmp_path,
):
    excel, audit = create_test_files(tmp_path)

    record = json.loads(
        audit.read_text(encoding="utf-8")
    )

    record["existing_count"] = -10

    audit.write_text(
        json.dumps(record),
        encoding="utf-8",
    )

    with pytest.raises(ValueError):
        verify_audit_integrity(excel, audit)
