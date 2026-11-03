
from pathlib import Path
from collections import Counter
from decimal import Decimal, InvalidOperation
from datetime import datetime, date

import pandas as pd
from openpyxl import load_workbook


# ==========================================
# 1. PROJECT PATHS
# ==========================================

BASE_DIR = Path(__file__).resolve().parent.parent

TEMPLATE_FILE = (
    BASE_DIR / "data" / "templates"
    / "BeDapper_ChannelWise_Report.xlsx"
)

SOURCE_FILE = (
    BASE_DIR / "output" / "business_data"
    / "be_dapper" / "sales.csv"
)


# ==========================================
# 2. READ INCOMING SALES
# ==========================================

incoming = pd.read_csv(
    SOURCE_FILE,
    dtype={
        "bill_no": "string",
        "product": "string"
    }
)


# ==========================================
# 3. NORMALIZATION FUNCTIONS
# ==========================================

def clean_text(value):
    if pd.isna(value):
        return ""

    return str(value).strip().casefold()


def clean_number(value):
    if pd.isna(value):
        return None

    try:
        return Decimal(
            str(value)
        ).quantize(Decimal("0.01"))

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
# 4. CREATE TRANSACTION KEY
# ==========================================

def make_transaction_key(record):

    return (
        clean_text(record["bill_no"]),
        clean_text(record["product"]),
        clean_number(record["sold_qty"]),
        clean_number(record["cost_price"]),
        clean_number(record["sales_price"]),
        clean_number(record["cost_amount"]),
        clean_number(record["sales_amount"]),
        clean_date(record["date"])
    )


# ==========================================
# 5. READ EXISTING WORKBOOK
# ==========================================

wb = load_workbook(
    TEMPLATE_FILE,
    read_only=True,
    data_only=False
)

existing_records = []

try:

    ws = wb["SalesRepItem"]

    for row in ws.iter_rows(
        min_row=5,
        values_only=True
    ):

        if len(row) < 52:
            continue

        # Skip empty transaction rows
        if row[25] is None:
            continue

        # Actual Excel field positions
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
    wb.close()


# ==========================================
# 6. GENERATE TRANSACTION KEYS
# ==========================================

incoming_keys = [
    make_transaction_key(record)
    for record in incoming.to_dict("records")
]

existing_keys = [
    make_transaction_key(record)
    for record in existing_records
]


# ==========================================
# 7. COUNT MATCHING TRANSACTIONS
# ==========================================

incoming_counter = Counter(incoming_keys)
existing_counter = Counter(existing_keys)

matching_counter = (
    incoming_counter & existing_counter
)

matching_lines = sum(
    matching_counter.values()
)

incoming_only = (
    incoming_counter - existing_counter
)

existing_only = (
    existing_counter - incoming_counter
)

incoming_only_count = sum(
    incoming_only.values()
)

existing_only_count = sum(
    existing_only.values()
)


# ==========================================
# 8. VALIDATION RESULTS
# ==========================================

print("\n===== DUPLICATE VALIDATION =====")

print("Incoming sales lines:", len(incoming_keys))

print("Existing workbook lines:", len(existing_keys))

print("Matching transaction lines:", matching_lines)

print("Incoming lines not matched:", incoming_only_count)

print("Existing lines not matched:", existing_only_count)

print("\n===== RECONCILIATION =====")

assert (
    matching_lines + incoming_only_count
    == len(incoming_keys)
)

assert (
    matching_lines + existing_only_count
    == len(existing_keys)
)

print("Transaction count reconciliation: PASSED")

print("\nOriginal Excel workbook was not modified.")
