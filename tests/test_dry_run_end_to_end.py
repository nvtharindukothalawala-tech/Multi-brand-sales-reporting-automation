
import hashlib
import json
import shutil
import subprocess
import sys
from pathlib import Path

import pandas as pd
from openpyxl import Workbook


PROJECT_ROOT = Path(__file__).resolve().parent.parent


# ==========================================
# 1. CREATE A FAKE SALES TRANSACTION
# ==========================================

def fake_sale(bill_no, sale_date="2026-10-10"):
    return {
        "bill_no": bill_no,
        "product": "TEST_SHIRT",
        "sold_qty": 1,
        "cost_price": 1000,
        "sales_price": 1500,
        "cost_amount": 1000,
        "sales_amount": 1500,
        "date": sale_date,
    }


# ==========================================
# 2. SET UP AN ISOLATED TEST PROJECT
# ==========================================

def create_test_project(tmp_path, approval="REVIEW",
                        sale_date="2026-10-10"):

    src_dir = tmp_path / "src"
    src_dir.mkdir()

    # Copy only the Python modules needed
    # to run the actual planner.
    for filename in [
        "__init__.py",
        "dry_run_update_planner.py",
        "transaction_matching.py",
        "planner_validation.py",
    ]:
        shutil.copy2(
            PROJECT_ROOT / "src" / filename,
            src_dir / filename,
        )

    template_dir = tmp_path / "data" / "templates"
    template_dir.mkdir(parents=True)

    sales_dir = (
        tmp_path / "output" / "business_data"
        / "be_dapper"
    )
    sales_dir.mkdir(parents=True)

    config_dir = tmp_path / "config"
    config_dir.mkdir()

    # Create fake incoming sales
    incoming = [
        fake_sale("TEST001", sale_date),
        fake_sale("TEST002", sale_date),
    ]

    pd.DataFrame(incoming).to_csv(
        sales_dir / "sales.csv",
        index=False,
    )

    # Create fake existing Excel transactions.
    # Only TEST001 already exists.
    wb = Workbook()
    ws = wb.active
    ws.title = "SalesRepItem"

    record = incoming[0]

    column_map = {
        "bill_no": "Z",
        "product": "AO",
        "cost_price": "AP",
        "sales_price": "AQ",
        "sold_qty": "AT",
        "cost_amount": "AU",
        "sales_amount": "AV",
        "date": "AZ",
    }

    for field, column in column_map.items():
        ws[f"{column}5"] = record[field]

    workbook_path = (
        template_dir
        / "BeDapper_ChannelWise_Report.xlsx"
    )

    wb.save(workbook_path)
    wb.close()

    # Create fake reporting-scope configuration
    pd.DataFrame([
        {
            "location": "TEST_STORE",
            "proposed_reporting_business": "TEST_BRAND",
            "status": approval,
        }
    ]).to_csv(
        config_dir / "reporting_scope_review.csv",
        index=False,
    )

    return workbook_path


# ==========================================
# 3. HELPER: RUN THE ACTUAL PLANNER
# ==========================================

def run_planner(tmp_path):
    return subprocess.run(
        [
            sys.executable,
            "-m",
            "src.dry_run_update_planner",
        ],
        cwd=tmp_path,
        capture_output=True,
        text=True,
        check=False,
    )


def file_hash(file_path):
    return hashlib.sha256(
        file_path.read_bytes()
    ).hexdigest()


# ==========================================
# 4. TEST: PENDING APPROVAL
# ==========================================

def test_pending_approval_end_to_end(tmp_path):

    workbook_path = create_test_project(
        tmp_path,
        approval="REVIEW",
    )

    original_hash = file_hash(workbook_path)

    result = run_planner(tmp_path)

    assert result.returncode == 0, result.stderr

    audit_path = (
        tmp_path / "output" / "audit"
        / "be_dapper_dry_run.json"
    )

    assert audit_path.is_file()

    audit = json.loads(
        audit_path.read_text(encoding="utf-8")
    )

    assert audit["incoming_sales_lines"] == 2
    assert audit["existing_workbook_lines"] == 1
    assert audit["matching_lines"] == 1
    assert audit["new_candidate_lines"] == 1
    assert audit["existing_only_lines"] == 0

    assert audit["decision"] == (
        "BLOCKED_PENDING_APPROVAL"
    )

    assert audit["update_authorized"] is False

    # Confirm the original Excel file is unchanged
    assert file_hash(workbook_path) == original_hash


# ==========================================
# 5. TEST: ALL APPROVED STILL REQUIRES REVIEW
# ==========================================

def test_approved_scope_requires_review(tmp_path):

    workbook_path = create_test_project(
        tmp_path,
        approval="APPROVED",
    )

    original_hash = file_hash(workbook_path)

    result = run_planner(tmp_path)

    assert result.returncode == 0, result.stderr

    audit_path = (
        tmp_path / "output" / "audit"
        / "be_dapper_dry_run.json"
    )

    audit = json.loads(
        audit_path.read_text(encoding="utf-8")
    )

    assert audit["decision"] == "REVIEW_REQUIRED"
    assert audit["update_authorized"] is False

    assert file_hash(workbook_path) == original_hash


# ==========================================
# 6. TEST: WRONG MONTH IS REJECTED
# ==========================================

def test_wrong_month_rejected_end_to_end(tmp_path):

    workbook_path = create_test_project(
        tmp_path,
        approval="APPROVED",
        sale_date="2026-11-10",
    )

    original_hash = file_hash(workbook_path)

    result = run_planner(tmp_path)

    # Wrong month should cause the planner to stop.
    assert result.returncode != 0

    assert (
        "outside the reporting month"
        in result.stderr
    )

    audit_path = (
        tmp_path / "output" / "audit"
        / "be_dapper_dry_run.json"
    )

    # No successful audit should be produced.
    assert not audit_path.exists()

    # Excel workbook must remain unchanged.
    assert file_hash(workbook_path) == original_hash
