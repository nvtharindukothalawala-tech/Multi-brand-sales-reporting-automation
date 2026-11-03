
from pathlib import Path
import json

import pandas as pd
from openpyxl import load_workbook

from src.transaction_matching import compare_transactions
from src.planner_validation import (
    validate_reporting_month,
    plan_update_decision,
)


# ==========================================
# 1. PROJECT CONFIGURATION
# ==========================================

BASE_DIR = Path(__file__).resolve().parent.parent

TEMPLATE_FILE = (
    BASE_DIR
    / "data"
    / "templates"
    / "BeDapper_ChannelWise_Report.xlsx"
)

SALES_FILE = (
    BASE_DIR
    / "output"
    / "business_data"
    / "be_dapper"
    / "sales.csv"
)

SCOPE_FILE = (
    BASE_DIR
    / "config"
    / "reporting_scope_review.csv"
)

AUDIT_DIR = (
    BASE_DIR
    / "output"
    / "audit"
)

REPORTING_MONTH = "2026-10"


# ==========================================
# 2. VALIDATE INPUT FILES
# ==========================================

required_files = [
    TEMPLATE_FILE,
    SALES_FILE,
    SCOPE_FILE,
]

for file_path in required_files:

    if not file_path.is_file():
        raise FileNotFoundError(
            f"Required file missing: {file_path}"
        )

    if file_path.stat().st_size == 0:
        raise ValueError(
            f"Required file is empty: {file_path}"
        )

print("\nInput file validation: PASSED")


# ==========================================
# 3. READ INCOMING SALES DATA
# ==========================================

incoming = pd.read_csv(
    SALES_FILE,
    dtype={
        "bill_no": "string",
        "product": "string",
    },
)

required_sales_columns = {
    "bill_no",
    "product",
    "sold_qty",
    "cost_price",
    "sales_price",
    "cost_amount",
    "sales_amount",
    "date",
}

missing_sales_columns = (
    required_sales_columns - set(incoming.columns)
)

if missing_sales_columns:
    raise ValueError(
        "Incoming sales CSV is missing columns: "
        f"{sorted(missing_sales_columns)}"
    )

if incoming.empty:
    raise ValueError(
        "Incoming sales CSV contains no transactions."
    )

print("\nIncoming sales data loaded successfully.")
print("Incoming Sales Lines:", len(incoming))


# ==========================================
# 4. VALIDATE REPORTING MONTH
# ==========================================

validate_reporting_month(
    dates=incoming["date"],
    expected_month=REPORTING_MONTH,
)

print("\nReporting month validation: PASSED")


# ==========================================
# 5. READ EXISTING EXCEL TRANSACTIONS
# ==========================================

workbook = load_workbook(
    TEMPLATE_FILE,
    read_only=True,
    data_only=False,
)

existing_records = []

try:

    if "SalesRepItem" not in workbook.sheetnames:
        raise ValueError(
            "Required sheet 'SalesRepItem' "
            "not found in Excel workbook."
        )

    ws = workbook["SalesRepItem"]

    for row in ws.iter_rows(
        min_row=5,
        values_only=True,
    ):

        # Column Z is the bill number.
        # Skip rows without a bill number.
        if len(row) < 52 or row[25] is None:
            continue

        existing_records.append(
            {
                "bill_no": row[25],       # Z
                "product": row[40],       # AO
                "cost_price": row[41],    # AP
                "sales_price": row[42],   # AQ
                "sold_qty": row[45],      # AT
                "cost_amount": row[46],   # AU
                "sales_amount": row[47],  # AV
                "date": row[51],          # AZ
            }
        )

finally:
    workbook.close()

print("\nExisting Excel transactions loaded.")
print("Existing Workbook Lines:", len(existing_records))


# ==========================================
# 6. COMPARE TRANSACTIONS
# ==========================================

incoming_records = incoming.to_dict("records")

comparison = compare_transactions(
    incoming_records,
    existing_records,
)

matched_count = comparison["matched"]
new_count = comparison["new"]
existing_only_count = comparison["existing_only"]

# Check that all transaction occurrences
# are accounted for.

if matched_count + new_count != len(incoming_records):
    raise ValueError(
        "Incoming transaction reconciliation failed."
    )

if (
    matched_count + existing_only_count
    != len(existing_records)
):
    raise ValueError(
        "Existing workbook reconciliation failed."
    )

print("\nTransaction comparison validated.")
print("Matching Lines:", matched_count)
print("New Candidate Lines:", new_count)
print("Existing-Only Lines:", existing_only_count)


# ==========================================
# 7. READ BUSINESS APPROVAL CONFIGURATION
# ==========================================

scope_review = pd.read_csv(
    SCOPE_FILE,
    dtype="string",
    keep_default_na=False,
)

required_scope_columns = {
    "location",
    "proposed_reporting_business",
    "status",
}

missing_scope_columns = (
    required_scope_columns - set(scope_review.columns)
)

if missing_scope_columns:
    raise ValueError(
        "Reporting-scope CSV is missing columns: "
        f"{sorted(missing_scope_columns)}"
    )

scope_review["status"] = (
    scope_review["status"]
    .str.strip()
    .str.upper()
)

# Check that at least one reporting-scope
# rule is present.
if scope_review.empty:
    print("\nWarning: No reporting-scope rules found.")


# ==========================================
# 8. BUSINESS APPROVAL VALIDATION
# ==========================================

pending_count = int(
    scope_review["status"].ne("APPROVED").sum()
)

decision = plan_update_decision(
    dates=incoming["date"],
    expected_month=REPORTING_MONTH,
    statuses=scope_review["status"].tolist(),
)

# This script is always READ-ONLY.
# Approval does not permit Excel modification.
update_authorized = False

print("\nBusiness approval validation completed.")
print("Pending Scope Rules:", pending_count)
print("Decision:", decision)
print("Update Authorized:", update_authorized)


# ==========================================
# 9. GENERATE DRY-RUN AUDIT REPORT
# ==========================================

AUDIT_DIR.mkdir(
    parents=True,
    exist_ok=True,
)

audit_report = {
    "business": "Be Dapper",
    "reporting_month": REPORTING_MONTH,
    "mode": "DRY_RUN",

    "incoming_sales_lines": len(incoming),
    "existing_workbook_lines": len(existing_records),

    "matching_lines": matched_count,
    "new_candidate_lines": new_count,
    "existing_only_lines": existing_only_count,

    "pending_scope_rules": pending_count,

    "decision": decision,
    "update_authorized": update_authorized,
    "original_workbook_modified": False,

    "note": (
        "Transaction comparison uses a composite "
        "matching key, not a unique transaction-line ID. "
        "No Excel workbook writing was performed. "
        "Reporting-scope approval is still required."
    ),
}

output_file = (
    AUDIT_DIR
    / "be_dapper_dry_run.json"
)

with open(
    output_file,
    "w",
    encoding="utf-8",
) as file:

    json.dump(
        audit_report,
        file,
        indent=4,
    )

print("\nAudit report generated successfully.")


# ==========================================
# 10. DISPLAY FINAL DRY-RUN RESULTS
# ==========================================

print("\n====================================")
print("       BE DAPPER DRY-RUN REPORT")
print("====================================")

print("\n===== REPORTING PERIOD =====")
print("Reporting Month:", REPORTING_MONTH)

print("\n===== TRANSACTION COMPARISON =====")
print("Incoming Sales Lines:", len(incoming))
print("Existing Workbook Lines:", len(existing_records))
print("Matching Lines:", matched_count)
print("New Candidate Lines:", new_count)
print("Existing-Only Lines:", existing_only_count)

print("\n===== BUSINESS APPROVAL =====")
print("Pending Scope Rules:", pending_count)
print("Decision:", decision)
print("Update Authorized:", update_authorized)

print("\n===== AUDIT OUTPUT =====")
print("Audit Report:", output_file)
print("Original Workbook Modified: No")

print("\n====================================")
print(" DRY-RUN COMPLETED SUCCESSFULLY")
print("====================================")
