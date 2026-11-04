"""Lesson 57: no real business files; all workbook fixtures are synthetic."""

import json

import pytest
from openpyxl import load_workbook

from src.audit_logging import calculate_file_hash
from src.audit_verification import verify_workbook_audit
from src.publication_inspection import (
    IncompletePublicationError,
    inspect_publication,
)
from src.synthetic_excel_writer import write_synthetic_copy
from tests.test_synthetic_excel_writer import fake_sale, make_workbook


def create_pair(tmp_path):
    source = tmp_path / "source.xlsx"
    output = tmp_path / "output.xlsx"
    audit = tmp_path / "audit.json"
    make_workbook(source)
    write_synthetic_copy(
        source, output, [fake_sale("TEST001"), fake_sale("TEST002")],
        audit_path=audit, run_id="LESSON-57",
    )
    return output, audit


def update_audit_checksum(output, audit):
    record = json.loads(audit.read_text(encoding="utf-8"))
    record["output_sha256"] = calculate_file_hash(output)
    audit.write_text(json.dumps(record), encoding="utf-8")


def test_orphan_transaction_row_is_rejected(tmp_path):
    output, audit = create_pair(tmp_path)
    workbook = load_workbook(output)
    try:
        sheet = workbook["SalesRepItem"]
        # A new transaction-like row with no Bill Number.
        sheet["AO7"] = "ORPHAN_PRODUCT"
        sheet["AT7"] = 1
        sheet["AV7"] = 900
        workbook.save(output)
    finally:
        workbook.close()
    update_audit_checksum(output, audit)
    with pytest.raises(ValueError, match="Orphan transaction"):
        verify_workbook_audit(output, audit)


def test_blank_formatted_row_is_allowed(tmp_path):
    output, audit = create_pair(tmp_path)
    workbook = load_workbook(output)
    try:
        sheet = workbook["SalesRepItem"]
        sheet["AO7"].number_format = "0.00"
        workbook.save(output)
    finally:
        workbook.close()
    update_audit_checksum(output, audit)
    assert verify_workbook_audit(output, audit)["verified"] is True


def test_complete_pair_is_verified(tmp_path):
    output, audit = create_pair(tmp_path)
    inspected = inspect_publication(output, audit)
    assert inspected["publication_status"] == "VERIFIED"
    assert inspected["financial_totals_verified"] is True


def test_absent_pair_is_not_published(tmp_path):
    inspected = inspect_publication(tmp_path / "none.xlsx", tmp_path / "none.json")
    assert inspected == {"publication_status": "NOT_PUBLISHED", "verified": False}


def test_excel_only_is_detected_without_deleting_file(tmp_path):
    output, audit = create_pair(tmp_path)
    audit.unlink()
    original_hash = calculate_file_hash(output)
    with pytest.raises(IncompletePublicationError, match="missing audit JSON"):
        inspect_publication(output, audit)
    assert calculate_file_hash(output) == original_hash
    assert output.exists()


def test_audit_only_is_detected_without_deleting_file(tmp_path):
    output, audit = create_pair(tmp_path)
    original_audit = audit.read_bytes()
    output.unlink()
    with pytest.raises(IncompletePublicationError, match="missing Excel workbook"):
        inspect_publication(output, audit)
    assert audit.read_bytes() == original_audit


def test_corrupted_complete_pair_is_rejected(tmp_path):
    output, audit = create_pair(tmp_path)
    output.write_bytes(b"not a workbook")
    with pytest.raises(ValueError, match="checksum"):
        inspect_publication(output, audit)
