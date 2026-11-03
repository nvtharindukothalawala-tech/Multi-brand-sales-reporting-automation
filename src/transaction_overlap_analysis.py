
from pathlib import Path
from collections import Counter
from datetime import datetime, date
from numbers import Real

import pandas as pd

from openpyxl import load_workbook
from openpyxl.utils.datetime import from_excel


# ==========================================
# 1. DEFINE PROJECT PATHS
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


# ==========================================
# 2. CHECK INPUT FILES
# ==========================================

if not TEMPLATE_FILE.is_file():
    raise FileNotFoundError(
        f"Excel template not found: {TEMPLATE_FILE}"
    )

if not SALES_FILE.is_file():
    raise FileNotFoundError(
        f"Sales CSV not found: {SALES_FILE}"
    )


# ==========================================
# 3. READ NEW SALES DATA
# ==========================================

new_sales = pd.read_csv(
    SALES_FILE,
    dtype={"bill_no": "string"}
)

print("\n===== NEW SALES DATA =====")
print("Incoming sales lines:", len(new_sales))


# ==========================================
# 4. CONVERT EXCEL DATES
# ==========================================

def convert_excel_date(value, excel_epoch):

    # Missing values
    if pd.isna(value):
        return pd.NaT

    # Python datetime or date values
    if isinstance(value, (datetime, date)):
        return pd.Timestamp(value)

    # Excel serial dates (integers/floats)
    if isinstance(value, Real) and not isinstance(
        value, bool
    ):
        try:
            converted = from_excel(
                value,
                epoch=excel_epoch
            )

            return pd.Timestamp(converted)

        except (ValueError, TypeError, OverflowError):
            return pd.NaT

    # Text dates
    if isinstance(value, str):

        value = value.strip()

        if not value:
            return pd.NaT

        return pd.to_datetime(
            value,
            errors="coerce"
        )

    return pd.NaT


# ==========================================
# 5. READ EXISTING WORKBOOK
# ==========================================

wb = load_workbook(
    TEMPLATE_FILE,
    read_only=True,
    data_only=False
)

existing_records = []
date_types = Counter()

try:

    excel_epoch = wb.epoch

    ws = wb["SalesRepItem"]

    # Sales data starts at row 4
    for row in ws.iter_rows(
        min_row=4,
        values_only=True
    ):

        # Z  = index 25 (Bill No)
        # AO = index 40 (Product)
        # AT = index 45 (Sold Qty)
        # AV = index 47 (Sales Amount)
        # AX = index 49 (Date)

        if len(row) < 50:
            continue

        bill_no = row[25]
        product = row[40]
        sold_qty = row[45]
        sales_amount = row[47]
        sale_date = row[49]

        if bill_no is None:
            continue

        date_types[
            type(sale_date).__name__
        ] += 1

        converted_date = convert_excel_date(
            sale_date,
            excel_epoch
        )

        existing_records.append({
            "bill_no": bill_no,
            "product": product,
            "sold_qty": sold_qty,
            "sales_amount": sales_amount,
            "date": sale_date,
            "parsed_date": converted_date
        })

finally:
    wb.close()


# ==========================================
# 6. CREATE EXISTING DATAFRAME
# ==========================================

existing_df = pd.DataFrame(
    existing_records,
    columns=[
        "bill_no",
        "product",
        "sold_qty",
        "sales_amount",
        "date",
        "parsed_date"
    ]
)


# ==========================================
# 7. DISPLAY DATE STORAGE TYPES
# ==========================================

print("\n===== DATE STORAGE TYPES =====")

for date_type, count in date_types.items():
    print(f"{date_type}: {count}")

print("\nExcel Date Epoch:", excel_epoch)


# ==========================================
# 8. STANDARDIZE DATES
# ==========================================

existing_df["parsed_date"] = pd.to_datetime(
    existing_df["parsed_date"],
    errors="coerce"
).dt.normalize()

new_sales["parsed_date"] = pd.to_datetime(
    new_sales["date"],
    format="%Y-%m-%d %H:%M:%S",
    errors="coerce"
).dt.normalize()


# ==========================================
# 9. DATE VALIDATION
# ==========================================

print("\n===== DATE VALIDATION =====")

print(
    "Existing invalid dates:",
    existing_df["parsed_date"].isna().sum()
)

print(
    "New invalid dates:",
    new_sales["parsed_date"].isna().sum()
)

print(
    "Existing earliest date:",
    existing_df["parsed_date"].min()
)

print(
    "Existing latest date:",
    existing_df["parsed_date"].max()
)

print(
    "New earliest date:",
    new_sales["parsed_date"].min()
)

print(
    "New latest date:",
    new_sales["parsed_date"].max()
)


# ==========================================
# 10. COUNT OCTOBER 2026 RECORDS
# ==========================================

october_mask = (
    existing_df["parsed_date"].dt.year.eq(2026)
    &
    existing_df["parsed_date"].dt.month.eq(10)
)

october_sales_count = int(
    october_mask.sum()
)

print("\n===== OCTOBER 2026 CHECK =====")

print(
    "Existing October 2026 sales lines:",
    october_sales_count
)


# ==========================================
# 11. NORMALIZE MATCHING COLUMNS
# ==========================================

for df in [existing_df, new_sales]:

    df["bill_no"] = (
        df["bill_no"]
        .astype("string")
        .str.strip()
    )

    df["product"] = (
        df["product"]
        .astype("string")
        .str.strip()
    )

    df["sold_qty"] = pd.to_numeric(
        df["sold_qty"],
        errors="coerce"
    )

    df["sales_amount"] = pd.to_numeric(
        df["sales_amount"],
        errors="coerce"
    ).round(2)


# ==========================================
# 12. DEFINE TRANSACTION MATCHING FIELDS
# ==========================================

match_fields = [
    "bill_no",
    "product",
    "sold_qty",
    "sales_amount",
    "parsed_date"
]


# ==========================================
# 13. COUNT EXISTING TRANSACTIONS
# ==========================================

existing_counts = (
    existing_df
    .groupby(
        match_fields,
        dropna=False
    )
    .size()
    .reset_index(
        name="existing_count"
    )
)


# ==========================================
# 14. COUNT NEW TRANSACTIONS
# ==========================================

new_counts = (
    new_sales
    .groupby(
        match_fields,
        dropna=False
    )
    .size()
    .reset_index(
        name="new_count"
    )
)


# ==========================================
# 15. COMPARE TRANSACTION LINES
# ==========================================

comparison = new_counts.merge(
    existing_counts,
    on=match_fields,
    how="left"
)

comparison["existing_count"] = (
    comparison["existing_count"]
    .fillna(0)
    .astype(int)
)

comparison["matching_lines"] = comparison[
    ["new_count", "existing_count"]
].min(axis=1)

matched = int(
    comparison["matching_lines"].sum()
)

unmatched = len(new_sales) - matched


# ==========================================
# 16. DISPLAY FINAL RESULTS
# ==========================================

print("\n===== TRANSACTION OVERLAP =====")

print(
    "Incoming sales lines:",
    len(new_sales)
)

print(
    "Existing sales lines:",
    len(existing_df)
)

print(
    "Matching incoming lines:",
    matched
)

print(
    "Incoming lines without an exact match:",
    unmatched
)

print("\n===== FINAL SUMMARY =====")

print(
    "Total existing workbook records:",
    len(existing_df)
)

print(
    "Existing October 2026 records:",
    october_sales_count
)

print(
    "Matching transaction lines:",
    matched
)

print(
    "Unmatched transaction lines:",
    unmatched
)

print("\nInvestigation completed successfully!")
print("Original Excel workbook was not modified.")
