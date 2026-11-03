
from pathlib import Path
from collections import defaultdict
from numbers import Real

import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

# ==========================================
# 1. DEFINE FILE PATHS
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

# ==========================================
# 2. READ NEW SALES DATA
# ==========================================

sales = pd.read_csv(
    SALES_FILE,
    dtype={"bill_no": "string"}
)

sales["bill_no"] = (
    sales["bill_no"].str.strip()
)

# Group incoming records by bill number
sales_by_bill = defaultdict(list)

for record in sales.to_dict("records"):
    sales_by_bill[str(record["bill_no"])].append(record)

# ==========================================
# 3. LOAD EXISTING WORKBOOK
# ==========================================

wb = load_workbook(
    TEMPLATE_FILE,
    read_only=True,
    data_only=False
)

# ==========================================
# 4. HELPER FUNCTIONS
# ==========================================

def clean_text(value):
    if pd.isna(value):
        return ""
    return str(value).strip().casefold()


def numbers_match(first, second):
    try:
        if pd.isna(first) or pd.isna(second):
            return False

        if isinstance(first, bool) or isinstance(second, bool):
            return False

        return abs(float(first) - float(second)) <= 0.01
    except (ValueError, TypeError):
        return False


# ==========================================
# 5. IDENTIFY POSSIBLE FIELD COLUMNS
# ==========================================

source_fields = [
    "product",
    "sold_qty",
    "cost_price",
    "sales_price",
    "cost_amount",
    "sales_amount",
    "discount",
    "delivery_amount",
]

numeric_fields = set(source_fields) - {"product"}

# Results are counts of matching source values
match_counts = {
    field: defaultdict(int)
    for field in source_fields
}

overlapping_workbook_rows = 0

try:
    ws = wb["SalesRepItem"]

    # Read rows after repeated headers
    for row in ws.iter_rows(
        min_row=5,
        values_only=True
    ):

        if len(row) < 52:
            continue

        # Column Z: existing Bill No
        bill_no = str(row[25]).strip() if row[25] is not None else ""

        if bill_no not in sales_by_bill:
            continue

        overlapping_workbook_rows += 1

        # Each source row with the same bill is a
        # candidate. Counts are diagnostic only.
        source_records = sales_by_bill[bill_no]

        for column_index in range(26, 52):

            existing_value = row[column_index]

            if existing_value is None:
                continue

            column_letter = get_column_letter(
                column_index + 1
            )

            for field in source_fields:

                matches = any(
                    numbers_match(existing_value, record[field])
                    if field in numeric_fields
                    else clean_text(existing_value) ==
                         clean_text(record[field])
                    for record in source_records
                )

                if matches:
                    match_counts[field][column_letter] += 1

finally:
    wb.close()

# ==========================================
# 6. DISPLAY RESULTS
# ==========================================

print("\n===== FIELD ALIGNMENT ANALYSIS =====")

print(
    "Existing workbook rows with overlapping bills:",
    overlapping_workbook_rows
)

for field in source_fields:

    print(f"\nSOURCE FIELD: {field}")

    ranked = sorted(
        match_counts[field].items(),
        key=lambda item: item[1],
        reverse=True
    )

    for column, count in ranked[:5]:
        print(
            f"  Candidate column {column}: "
            f"{count} matching workbook rows"
        )

print("\nAnalysis completed successfully!")
print("Original workbook was not modified.")
