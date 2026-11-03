
from pathlib import Path
from collections import Counter
from decimal import Decimal, InvalidOperation
from datetime import date, datetime
import json

import pandas as pd
from openpyxl import load_workbook
from src.transaction_matching import compare_transactions

# ==========================================
# 1. PROJECT CONFIGURATION
# ==========================================

BASE_DIR = Path(__file__).resolve().parent.parent

TEMPLATE_FILE = (
    BASE_DIR / "data" / "templates"
    / "BeDapper_ChannelWise_Report.xlsx"
)

SALES_FILE = (
    BASE_DIR / "output" / "business_data"
    / "be_dapper" / "sales.csv"
)

SCOPE_FILE = (
    BASE_DIR / "config"
    / "reporting_scope_review.csv"
)

AUDIT_DIR = BASE_DIR / "output" / "audit"

REPORTING_MONTH = "2026-10"


# ==========================================
# 2. DATA NORMALIZATION FUNCTIONS
# ==========================================

def clean_text(value):
    if pd.isna(value):
        return ""

    return str(value).strip().casefold()


def clean_number(value):
    if pd.isna(value):
        return None

    try:
        return Decimal(str(value)).quantize(
            Decimal("0.01")
        )
    except (InvalidOperation, ValueError):
        return None


def clean_date(value):
    if value is None or pd.isna(value):
        return None

    if isinstance(value, (datetime, date)):
        return value.strftime("%Y-%m-%d")

    parsed = pd.to_datetime(
        str(value).strip(),
        errors="coerce"
    )

    if pd.isna(parsed):
        return None

    return parsed.strftime("%Y-%m-%d")


# ==========================================
# 3. CREATE TRANSACTION MATCHING KEY
# ==========================================

def make_key(record):
    return (
        clean_text(record["bill_no"]),
        clean_text(record["product"]),
        clean_number(record["sold_qty"]),
        clean_number(record["cost_price"]),
        clean_number(record["sales_price"]),
        clean_number(record["cost_amount"]),
        clean_number(record["sales_amount"]),
        clean_date(record["date"]),
    )


# ==========================================
# 4. VALIDATE INPUT FILES
# ==========================================

for file_path in [
    TEMPLATE_FILE,
    SALES_FILE,
    SCOPE_FILE
]:
    if not file_path.is_file():
        raise FileNotFoundError(
            f"Required file missing: {file_path}"
        )


# ==========================================
# 5. READ INCOMING SALES DATA
# ==========================================

incoming = pd.read_csv(
    SALES_FILE,
    dtype={
        "bill_no": "string",
        "product": "string"
    }
)

required_columns = {
    "bill_no", "product", "sold_qty",
    "cost_price", "sales_price",
    "cost_amount", "sales_amount", "date"
}

if not required_columns.issubset(incoming.columns):
    raise ValueError(
        "Incoming sales CSV is missing required columns."
    )

# Check the reporting month
source_dates = pd.to_datetime(
    incoming["date"],
    format="%Y-%m-%d %H:%M:%S",
    errors="coerce"
)

invalid_source_dates = int(source_dates.isna().sum())

wrong_month_rows = int(
    (
        source_dates.notna()
        & source_dates.dt.strftime("%Y-%m").ne(
            REPORTING_MONTH
        )
    ).sum()
)

if invalid_source_dates or wrong_month_rows:
    raise ValueError(
        "Invalid sales dates or records outside "
        "the reporting month. Review the data."
    )


# ==========================================
# 6. READ EXISTING EXCEL TRANSACTIONS
# ==========================================

workbook = load_workbook(
    TEMPLATE_FILE,
    read_only=True,
    data_only=False
)

existing_records = []

try:
    ws = workbook["SalesRepItem"]

    for row in ws.iter_rows(
        min_row=5,
        values_only=True
    ):

        if len(row) < 52 or row[25] is None:
            continue

        existing_records.append({
            "bill_no": row[25],       # Z
            "product": row[40],       # AO
            "cost_price": row[41],    # AP
            "sales_price": row[42],   # AQ
            "sold_qty": row[45],      # AT
            "cost_amount": row[46],   # AU
            "sales_amount": row[47],  # AV
            "date": row[51]           # AZ
        })

finally:
    workbook.close()



# ==========================================
# 7. COMPARE TRANSACTIONS
# ==========================================

# Convert incoming DataFrame into records
incoming_records = incoming.to_dict("records")

# Use the reusable, tested matching function
comparison = compare_transactions(
    incoming_records,
    existing_records
)

matched_count = comparison["matched"]
new_count = comparison["new"]
existing_only_count = comparison["existing_only"]

# Validate transaction counts
assert (
    matched_count + new_count
    == len(incoming_records)
), "Incoming transaction count mismatch"

assert (
    matched_count + existing_only_count
    == len(existing_records)
), "Existing workbook count mismatch"

print("\nTransaction comparison validated.")



# ==========================================
# 8. CHECK BUSINESS APPROVAL STATUS
# ==========================================

scope_review = pd.read_csv(
    SCOPE_FILE,
    dtype="string"
)

required_scope_columns = {
    "location",
    "proposed_reporting_business",
    "status"
}

if not required_scope_columns.issubset(
    scope_review.columns
):
    raise ValueError(
        "Reporting-scope configuration is incomplete."
    )

scope_review["status"] = (
    scope_review["status"]
    .str.strip()
    .str.upper()
)

# Every required rule must be approved.
all_approved = (
    len(scope_review) > 0
    and scope_review["status"].eq("APPROVED").all()
)

pending_count = int(
    scope_review["status"].ne(
        "APPROVED"
    ).sum()
)

# Current planner never authorizes Excel writes.
update_authorized = False

if not all_approved:
    decision = "BLOCKED_PENDING_APPROVAL"
else:
    decision = "REVIEW_REQUIRED"


# ==========================================
# 9. GENERATE DRY-RUN AUDIT REPORT
# ==========================================

AUDIT_DIR.mkdir(
    parents=True,
    exist_ok=True
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
        "Matches are based on a composite "
        "transaction key, not a unique line ID. "
        "No Excel writing was performed."
    )
}

output_file = (
    AUDIT_DIR / "be_dapper_dry_run.json"
)

with open(
    output_file,
    "w",
    encoding="utf-8"
) as file:
    json.dump(
        audit_report,
        file,
        indent=4
    )


# ==========================================
# 10. DISPLAY FINAL RESULTS
# ==========================================

print("\n===== BE DAPPER DRY-RUN =====")

print("Reporting Month:", REPORTING_MONTH)
print("Incoming Sales Lines:", len(incoming))
print("Existing Workbook Lines:", len(existing_records))
print("Matching Lines:", matched_count)
print("New Candidate Lines:", new_count)
print("Existing-Only Lines:", existing_only_count)

print("\n===== BUSINESS APPROVAL =====")

print("Pending Scope Rules:", pending_count)
print("Decision:", decision)
print("Update Authorized:", update_authorized)

print("\n===== OUTPUT =====")

print("Audit Report:", output_file)
print("Original Workbook Modified: No")

print("\nDry-run planner completed successfully!")
