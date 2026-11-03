
# ==========================================
# LESSON 39 - AUDIT LOGGING TESTS
# ==========================================

import json
import hashlib
import pytest

from src.audit_logging import (
    create_audit_record,
    save_audit_log,
    calculate_file_hash,
)


# TEST 1: CREATE AUDIT RECORD

def test_create_audit_record():

    record = create_audit_record(
        run_id="TEST-RUN-001",
        existing_count=100,
        inserted_count=5,
        reconciled=True,
    )

    assert record["run_id"] == "TEST-RUN-001"
    assert record["existing_count"] == 100
    assert record["inserted_count"] == 5
    assert record["final_count"] == 105
    assert record["reconciled"] is True
    assert "timestamp" in record


# TEST 2: SAVE JSON AUDIT LOG

def test_save_audit_log(tmp_path):

    output = tmp_path / "audit.json"

    record = create_audit_record(
        run_id="TEST-RUN-002",
        existing_count=10,
        inserted_count=2,
        reconciled=True,
    )

    save_audit_log(record, output)

    assert output.exists()

    with open(
        output,
        "r",
        encoding="utf-8",
    ) as file:
        saved = json.load(file)

    assert saved["run_id"] == "TEST-RUN-002"
    assert saved["final_count"] == 12


# TEST 3: VERIFY FILE CHECKSUM

def test_calculate_file_hash(tmp_path):

    file_path = tmp_path / "test.txt"

    file_path.write_text(
        "SYNTHETIC_TEST_DATA",
        encoding="utf-8",
    )

    result = calculate_file_hash(file_path)

    expected = hashlib.sha256(
        file_path.read_bytes()
    ).hexdigest()

    assert result == expected
    assert len(result) == 64
