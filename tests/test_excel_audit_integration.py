
# ==========================================
# LESSON 41
# EXCEL AUDIT INTEGRATION TESTS
# ==========================================

import json

import pytest

from src.audit_logging import calculate_file_hash
from src.synthetic_excel_writer import (
    write_synthetic_copy,
)

from tests.test_synthetic_excel_writer import (
    make_workbook,
    fake_sale,
    sha256,
)


# ==========================================
# TEST 1: SUCCESSFUL AUDIT CREATION
# ==========================================

def test_writer_creates_verified_audit(tmp_path):

    source = tmp_path / "original.xlsx"
    output = tmp_path / "updated.xlsx"
    audit = tmp_path / "audit.json"

    make_workbook(source)

    source_hash = sha256(source)

    result = write_synthetic_copy(
        source,
        output,
        [
            fake_sale("TEST001"),
            fake_sale("TEST002"),
        ],
        audit_path=audit,
        run_id="TEST-RUN-001",
    )

    assert output.exists()
    assert audit.exists()

    assert result["reconciled"] is True
    assert result["inserted"] == 1

    with open(
        audit,
        "r",
        encoding="utf-8",
    ) as file:
        record = json.load(file)

    assert record["run_id"] == "TEST-RUN-001"
    assert record["existing_count"] == 1
    assert record["inserted_count"] == 1
    assert record["final_count"] == 2
    assert record["reconciled"] is True
    assert record["status"] == "VERIFIED"

    assert record["output_sha256"] == (
        calculate_file_hash(output)
    )

    # Original workbook must not change.
    assert sha256(source) == source_hash


# ==========================================
# TEST 2: PROTECT EXISTING AUDIT FILE
# ==========================================

def test_existing_audit_blocks_output(tmp_path):

    source = tmp_path / "original.xlsx"
    output = tmp_path / "updated.xlsx"
    audit = tmp_path / "audit.json"

    make_workbook(source)

    source_hash = sha256(source)

    # Simulate an existing audit record.
    audit.write_text(
        '{"status": "EXISTING"}',
        encoding="utf-8",
    )

    existing_audit_hash = sha256(audit)

    with pytest.raises(FileExistsError):

        write_synthetic_copy(
            source,
            output,
            [
                fake_sale("TEST001"),
                fake_sale("TEST002"),
            ],
            audit_path=audit,
            run_id="TEST-RUN-002",
        )

    # Neither existing file is modified.
    assert sha256(source) == source_hash
    assert sha256(audit) == existing_audit_hash

    # No report should be published.
    assert not output.exists()
