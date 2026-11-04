"""Lesson 56: safety checks using synthetic-only workbook fixtures."""

import json
from decimal import Decimal

import pytest
from openpyxl import load_workbook

from src.audit_logging import calculate_file_hash
from src.audit_verification import verify_audit_integrity, verify_workbook_audit
from src.synthetic_excel_writer import write_synthetic_copy
from tests.test_synthetic_excel_writer import fake_sale, make_workbook


def create_pair(tmp_path):
    source = tmp_path / "source.xlsx"
    output = tmp_path / "output.xlsx"
    audit = tmp_path / "audit.json"
    make_workbook(source)
    original_hash = calculate_file_hash(source)
    write_synthetic_copy(
        source,
        output,
        [fake_sale("TEST001"), fake_sale("TEST002")],
        audit_path=audit,
        run_id="LESSON-56",
    )
    return source, output, audit, original_hash


def test_version2_audit_requires_all_financial_totals(tmp_path):
    _, output, audit, _ = create_pair(tmp_path)
    data = json.loads(audit.read_text(encoding="utf-8"))
    assert data["audit_schema_version"] == 2
    data.pop("financial_totals")
    audit.write_text(json.dumps(data), encoding="utf-8")
    with pytest.raises(ValueError, match="missing financial totals"):
        verify_workbook_audit(output, audit)


@pytest.mark.parametrize("bad_value", [None, "NaN", "Infinity", "not-numeric", -1])
def test_version2_audit_rejects_invalid_sales_total(tmp_path, bad_value):
    _, output, audit, _ = create_pair(tmp_path)
    data = json.loads(audit.read_text(encoding="utf-8"))
    data["financial_totals"]["sales_amount"] = bad_value
    audit.write_text(json.dumps(data), encoding="utf-8")
    with pytest.raises(ValueError):
        verify_workbook_audit(output, audit)


def test_legacy_audit_explicitly_reports_unchecked_totals(tmp_path):
    _, output, audit, _ = create_pair(tmp_path)
    data = json.loads(audit.read_text(encoding="utf-8"))
    data.pop("audit_schema_version")
    data.pop("financial_totals")
    audit.write_text(json.dumps(data), encoding="utf-8")
    result = verify_workbook_audit(output, audit)
    assert result["legacy_audit"] is True
    assert result["financial_totals_checked"] is False
    assert result["financial_totals_verified"] is False
    assert result["audit_schema_version"] == 1


def test_writer_preserves_source_workbook(tmp_path):
    source, output, audit, original_hash = create_pair(tmp_path)
    assert source.exists()
    assert calculate_file_hash(source) == original_hash
    result = verify_workbook_audit(output, audit)
    assert result["financial_totals_verified"] is True
    assert result["legacy_audit"] is False


def test_workbook_rejects_nonfinite_sale_amount(tmp_path):
    _, output, audit, _ = create_pair(tmp_path)
    workbook = load_workbook(output)
    try:
        workbook["SalesRepItem"]["AV6"] = "Infinity"
        workbook.save(output)
    finally:
        workbook.close()
    data = json.loads(audit.read_text(encoding="utf-8"))
    data["output_sha256"] = calculate_file_hash(output)
    audit.write_text(json.dumps(data), encoding="utf-8")
    with pytest.raises(ValueError):
        verify_workbook_audit(output, audit)
