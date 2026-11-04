"""Lesson 59: standalone synthetic publication recovery reports."""

import json
from datetime import datetime

import pytest

from src.recovery_reporting import create_recovery_report
from tests.test_publication_inspection import create_pair


def test_verified_pair_report_contains_real_assessment(tmp_path):
    excel, audit = create_pair(tmp_path)
    report_path = tmp_path / "recovery.json"
    before_excel, before_audit = excel.read_bytes(), audit.read_bytes()

    report = create_recovery_report(excel, audit, report_path, run_id="L59-001")

    assert json.loads(report_path.read_text(encoding="utf-8")) == report
    assert report["report_type"] == "SYNTHETIC_PUBLICATION_RECOVERY_ASSESSMENT"
    assert report["report_schema_version"] == 1
    assert report["run_id"] == "L59-001"
    assert report["recovery_status"] == "VERIFIED"
    assert report["safe_to_use"] is True
    assert report["financial_totals_verified"] is True
    assert report["manual_review_required"] is False
    assert datetime.fromisoformat(report["created_at_utc"]).tzinfo is not None
    assert excel.read_bytes() == before_excel
    assert audit.read_bytes() == before_audit


def test_absent_pair_creates_not_published_report(tmp_path):
    excel, audit, destination = (
        tmp_path / "absent.xlsx", tmp_path / "absent.json", tmp_path / "recovery.json"
    )
    result = create_recovery_report(excel, audit, destination)
    assert result["recovery_status"] == "NOT_PUBLISHED"
    assert result["safe_to_use"] is False
    assert not excel.exists() and not audit.exists()


def test_excel_only_is_reported_without_modification(tmp_path):
    excel, audit = create_pair(tmp_path)
    audit.unlink()
    old_excel = excel.read_bytes()
    result = create_recovery_report(excel, audit, tmp_path / "recovery.json")
    assert result["recovery_status"] == "INCOMPLETE"
    assert result["manual_review_required"] is True
    assert excel.read_bytes() == old_excel
    assert not audit.exists()


def test_audit_only_is_reported_without_modification(tmp_path):
    excel, audit = create_pair(tmp_path)
    excel.unlink()
    old_audit = audit.read_bytes()
    result = create_recovery_report(excel, audit, tmp_path / "recovery.json")
    assert result["recovery_status"] == "INCOMPLETE"
    assert result["safe_to_use"] is False
    assert audit.read_bytes() == old_audit
    assert not excel.exists()


def test_corrupted_report_pair_is_flagged_for_review(tmp_path):
    excel, audit = create_pair(tmp_path)
    excel.write_bytes(b"damaged")
    old_excel, old_audit = excel.read_bytes(), audit.read_bytes()
    result = create_recovery_report(excel, audit, tmp_path / "recovery.json")
    assert result["recovery_status"] == "REVIEW_REQUIRED"
    assert result["safe_to_use"] is False
    assert excel.read_bytes() == old_excel and audit.read_bytes() == old_audit


def test_legacy_audit_is_not_marked_safe(tmp_path):
    excel, audit = create_pair(tmp_path)
    data = json.loads(audit.read_text(encoding="utf-8"))
    data.pop("audit_schema_version", None)
    data.pop("financial_totals", None)
    audit.write_text(json.dumps(data), encoding="utf-8")
    result = create_recovery_report(excel, audit, tmp_path / "recovery.json")
    assert result["recovery_status"] == "VERIFIED"
    assert result["legacy_audit"] is True
    assert result["financial_totals_verified"] is False
    assert result["safe_to_use"] is False
    assert result["manual_review_required"] is True


def test_existing_report_cannot_be_overwritten(tmp_path):
    destination = tmp_path / "recovery.json"
    destination.write_text("KEEP THIS", encoding="utf-8")
    with pytest.raises(FileExistsError):
        create_recovery_report(tmp_path / "missing.xlsx", tmp_path / "missing.json", destination)
    assert destination.read_text(encoding="utf-8") == "KEEP THIS"


@pytest.mark.parametrize("which", ["excel", "audit"])
def test_report_path_cannot_equal_input_path(tmp_path, which):
    excel = tmp_path / "source.xlsx"
    audit = tmp_path / "audit.json"
    destination = excel if which == "excel" else audit
    with pytest.raises(ValueError):
        create_recovery_report(excel, audit, destination)
    assert not excel.exists() and not audit.exists()


def test_no_temporary_files_remain_after_success(tmp_path):
    destination = tmp_path / "recovery.json"
    create_recovery_report(tmp_path / "missing.xlsx", tmp_path / "missing.json", destination)
    assert [path.name for path in tmp_path.iterdir()] == ["recovery.json"]
