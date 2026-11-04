"""Lesson 58: synthetic-only, read-only publication recovery assessment."""

import json

from src.audit_logging import calculate_file_hash
from src.publication_recovery import assess_publication_recovery
from tests.test_publication_inspection import create_pair


def test_complete_pair_is_verified_and_unchanged(tmp_path):
    output, audit = create_pair(tmp_path)
    before_excel = calculate_file_hash(output)
    before_audit = audit.read_bytes()

    result = assess_publication_recovery(output, audit)

    assert result["recovery_status"] == "VERIFIED"
    assert result["verified"] is True
    assert result["financial_totals_verified"] is True
    assert result["safe_to_use"] is True
    assert result["manual_review_required"] is False
    assert calculate_file_hash(output) == before_excel
    assert audit.read_bytes() == before_audit


def test_both_missing_is_not_published(tmp_path):
    output = tmp_path / "absent.xlsx"
    audit = tmp_path / "absent.json"
    result = assess_publication_recovery(output, audit)
    assert result["recovery_status"] == "NOT_PUBLISHED"
    assert result["verified"] is False
    assert result["safe_to_use"] is False
    assert not output.exists() and not audit.exists()


def test_only_excel_requires_manual_review_without_deletion(tmp_path):
    output, audit = create_pair(tmp_path)
    audit.unlink()
    original_hash = calculate_file_hash(output)
    result = assess_publication_recovery(output, audit)
    assert result["recovery_status"] == "INCOMPLETE"
    assert result["manual_review_required"] is True
    assert result["safe_to_use"] is False
    assert calculate_file_hash(output) == original_hash
    assert not audit.exists()


def test_only_audit_requires_manual_review_without_deletion(tmp_path):
    output, audit = create_pair(tmp_path)
    output.unlink()
    original_audit = audit.read_bytes()
    result = assess_publication_recovery(output, audit)
    assert result["recovery_status"] == "INCOMPLETE"
    assert result["manual_review_required"] is True
    assert result["safe_to_use"] is False
    assert audit.read_bytes() == original_audit
    assert not output.exists()


def test_corrupted_workbook_requires_review_without_modification(tmp_path):
    output, audit = create_pair(tmp_path)
    output.write_bytes(b"damaged workbook")
    previous_excel = output.read_bytes()
    previous_audit = audit.read_bytes()
    result = assess_publication_recovery(output, audit)
    assert result["recovery_status"] == "REVIEW_REQUIRED"
    assert result["verified"] is False
    assert output.read_bytes() == previous_excel
    assert audit.read_bytes() == previous_audit


def test_corrupted_audit_requires_review_without_modification(tmp_path):
    output, audit = create_pair(tmp_path)
    audit.write_text("{bad json", encoding="utf-8")
    previous_excel = output.read_bytes()
    previous_audit = audit.read_bytes()
    result = assess_publication_recovery(output, audit)
    assert result["recovery_status"] == "REVIEW_REQUIRED"
    assert output.read_bytes() == previous_excel
    assert audit.read_bytes() == previous_audit


def test_legacy_audit_is_flagged_for_financial_review(tmp_path):
    output, audit = create_pair(tmp_path)
    record = json.loads(audit.read_text(encoding="utf-8"))
    record.pop("audit_schema_version", None)
    record.pop("financial_totals", None)
    audit.write_text(json.dumps(record), encoding="utf-8")
    previous_excel = output.read_bytes()
    previous_audit = audit.read_bytes()
    result = assess_publication_recovery(output, audit)
    assert result["recovery_status"] == "VERIFIED"
    assert result["financial_totals_verified"] is False
    assert result["manual_review_required"] is True
    assert result["safe_to_use"] is False
    assert output.read_bytes() == previous_excel
    assert audit.read_bytes() == previous_audit
