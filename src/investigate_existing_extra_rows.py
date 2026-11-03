
from pathlib import Path
from collections import Counter, defaultdict
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
# 2. NORMALIZATION FUNCTIONS
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

    result = pd.to_datetime(
        str(value).strip(),
        errors="coerce"
    )

    if pd.isna(result):
        return None

    return result.strftime("%Y-%m-%d")


def make_key(record):
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
# 3. LOAD NEW CSV
# ==========================================

incoming = pd.read_csv(
    SOURCE_FILE,
    dtype={"bill_no": "string"}
)

incoming_keys = Counter(
    make_key(record)
    for record in incoming.to_dict("records")
)

# ==========================================
# 4. LOAD EXISTING WORKBOOK
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

        if len(row) < 52 or row[25] is None:
            continue

        existing_records.append({
            "bill_no": row[25],        # Z
            "location": row[26],       # AA
            "product": row[40],        # AO
            "cost_price": row[41],     # AP
            "sales_price": row[42],    # AQ
            "sold_qty": row[45],       # AT
            "cost_amount": row[46],    # AU
            "sales_amount": row[47],   # AV
            "date": row[51]            # AZ
        })

finally:
    wb.close()

# ==========================================
# 5. IDENTIFY EXISTING-ONLY LINES
# ==========================================

remaining_incoming = incoming_keys.copy()
extra_records = []

for record in existing_records:

    key = make_key(record)

    if remaining_incoming[key] > 0:
        remaining_incoming[key] -= 1
    else:
        extra_records.append(record)

# ==========================================
# 6. SUMMARIZE BY LOCATION
# ==========================================

location_summary = defaultdict(
    lambda: {
        "lines": 0,
        "sales_amount": Decimal("0.00")
    }
)

for record in extra_records:

    location = str(
        record["location"]
    ).strip()

    location_summary[location]["lines"] += 1

    amount = clean_number(
        record["sales_amount"]
    )

    if amount is not None:
        location_summary[location][
            "sales_amount"
        ] += amount

print("\n===== EXISTING-ONLY RECORDS =====")

print("Total existing workbook lines:", len(existing_records))
print("Incoming sales lines:", len(incoming))
print("Additional existing lines:", len(extra_records))

print("\n===== ADDITIONAL LINES BY LOCATION =====")

for location, values in sorted(
    location_summary.items(),
    key=lambda item: item[1]["lines"],
    reverse=True
):
    print(
        f"{location}: "
        f"{values['lines']} lines, "
        f"Sales Amount = {values['sales_amount']}"
    )

print("\nInvestigation completed!")
print("Original workbook was not modified.")
