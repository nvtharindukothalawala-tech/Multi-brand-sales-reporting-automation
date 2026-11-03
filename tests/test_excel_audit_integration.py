
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


# ==========================================
# LESSON 43 - FAILURE RECOVERY TESTS
# ==========================================

import os


# ==========================================
# TEST 3: AUDIT PUBLICATION FAILURE
# ==========================================

def test_audit_publication_failure_rolls_back(
    tmp_path,
    monkeypatch,
):
    source = tmp_path / "original.xlsx"
    output = tmp_path / "updated.xlsx"
    audit = tmp_path / "audit.json"

    make_workbook(source)

    source_hash = sha256(source)
    real_link = os.link

    def simulated_link(src, dst, *args, **kwargs):
        if os.fspath(dst) == os.fspath(audit):
            raise PermissionError(
                "Simulated audit permission error"
            )

        return real_link(
            src, dst, *args, **kwargs
        )

    monkeypatch.setattr(
        "src.synthetic_excel_writer.os.link",
        simulated_link,
    )

    with pytest.raises(
        PermissionError,
        match="Simulated audit permission error",
    ):
        write_synthetic_copy(
            source,
            output,
            [
                fake_sale("TEST001"),
                fake_sale("TEST002"),
            ],
            audit_path=audit,
            run_id="TEST-ROLLBACK-001",
        )

    # Published Excel should be rolled back.
    assert not output.exists()

    # Failed audit should not be published.
    assert not audit.exists()

    # Original workbook remains unchanged.
    assert sha256(source) == source_hash


# ==========================================
# TEST 4: EXISTING EXCEL OUTPUT PROTECTED
# ==========================================

def test_existing_excel_output_is_preserved(
    tmp_path,
):
    source = tmp_path / "original.xlsx"
    output = tmp_path / "updated.xlsx"
    audit = tmp_path / "audit.json"

    make_workbook(source)

    output.write_bytes(
        b"EXISTING_TEST_OUTPUT"
    )

    original_hash = sha256(source)
    output_hash = sha256(output)

    with pytest.raises(FileExistsError):
        write_synthetic_copy(
            source,
            output,
            [
                fake_sale("TEST001"),
                fake_sale("TEST002"),
            ],
            audit_path=audit,
            run_id="TEST-EXISTING-001",
        )

    assert sha256(source) == original_hash
    assert sha256(output) == output_hash
    assert not audit.exists()


# ==========================================
# TEST 5: INVALID AUDIT CONFIGURATION
# ==========================================

@pytest.mark.parametrize(
    "audit_path,run_id",
    [
        ("audit.json", None),
        (None, "TEST-RUN-001"),
    ],
)
def test_incomplete_audit_configuration_is_rejected(
    tmp_path,
    audit_path,
    run_id,
):
    source = tmp_path / "original.xlsx"
    output = tmp_path / "updated.xlsx"

    make_workbook(source)

    original_hash = sha256(source)

    if audit_path is not None:
        audit_path = tmp_path / audit_path

    with pytest.raises(ValueError):
        write_synthetic_copy(
            source,
            output,
            [
                fake_sale("TEST001"),
                fake_sale("TEST002"),
            ],
            audit_path=audit_path,
            run_id=run_id,
        )

    assert not output.exists()
    assert sha256(source) == original_hash
