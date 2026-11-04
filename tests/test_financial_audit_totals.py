"""Lesson 55: tests using synthetic workbook fixtures only."""

import json

import pytest
from openpyxl import load_workbook

from src.audit_logging import calculate_file_hash
from src.audit_verification import verify_workbook_audit
from src.synthetic_excel_writer import write_synthetic_copy
from tests.test_synthetic_excel_writer import make_workbook, fake_sale


def _create_verified_pair(tmp_path, run_id):
    source = tmp_path / "original.xlsx"
    output = tmp_path / "updated.xlsx"
    audit = tmp_path / "audit.json"
    make_workbook(source)
    write_synthetic_copy(
        source, output, [fake_sale("TEST001"), fake_sale("TEST002")],
        audit_path=audit, run_id=run_id,
    )
    return source, output, audit


def test_new_audit_contains_verified_financial_totals(tmp_path):
    source, output, audit = _create_verified_pair(tmp_path, "TEST-TOTAL-001")
    audit_data = json.loads(audit.read_text(encoding="utf-8"))
    assert set(audit_data["financial_totals"]) == {
        "sold_qty", "cost_amount", "sales_amount"
    }
    result = verify_workbook_audit(output, audit)
    assert result["verified"] is True
    assert result["financial_totals_checked"] is True


def test_changed_sale_amount_with_updated_hash_is_rejected(tmp_path):
    source, output, audit = _create_verified_pair(tmp_path, "TEST-TOTAL-002")
    wb = load_workbook(output)
    try:
        ws = wb["SalesRepItem"]
        ws["AV6"] = ws["AV6"].value + 100
        wb.save(output)
    finally:
        wb.close()

    audit_data = json.loads(audit.read_text(encoding="utf-8"))
    audit_data["output_sha256"] = calculate_file_hash(output)
    audit.write_text(json.dumps(audit_data), encoding="utf-8")

    with pytest.raises(ValueError, match="financial total mismatch"):
        verify_workbook_audit(output, audit)
