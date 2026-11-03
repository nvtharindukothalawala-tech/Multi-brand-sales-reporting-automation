
from pathlib import Path
from collections import Counter
from datetime import datetime, date
import pandas as pd
from openpyxl import load_workbook

# ==========================================
# 1. PROJECT PATHS
# ==========================================

BASE_DIR = Path(__file__).resolve().parent.parent

TEMPLATE_FILE = (
    BASE_DIR
    / "data"
    / "templates"
    / "BeDapper_ChannelWise_Report.xlsx"
)

BUSINESS_DIR = (
    BASE_DIR
    / "output"
    / "business_data"
    / "be_dapper"
)

# ==========================================
# 2. READ BUSINESS DATA
# ==========================================

sales = pd.read_csv(
    BUSINESS_DIR / "sales.csv",
    dtype={"bill_no": "string"}
)

returns = pd.read_csv(
    BUSINESS_DIR / "returns.csv",
    dtype={
        "bill_no": "string",
        "return_code": "string"
    }
)

# ==========================================
# 3. OPEN WORKBOOK IN READ-ONLY MODE
# ==========================================

wb = load_workbook(
    TEMPLATE_FILE,
    read_only=True,
    data_only=False
)

try:

    sales_ws = wb["SalesRepItem"]
    returns_ws = wb["Ord.Rtn"]

    # ======================================
    # 4. READ EXISTING SALES IDENTIFIERS
    # ======================================

    existing_sales_bills = []

    existing_sales_rows = 0
    october_sales_rows = 0

    for row in sales_ws.iter_rows(
        min_row=4,
        values_only=True
    ):

        # Column Z = index 25
        # Column AX = index 49

        bill_no = row[25] if len(row) > 25 else None
        sale_date = row[49] if len(row) > 49 else None

        if bill_no is None:
            continue

        existing_sales_rows += 1

        existing_sales_bills.append(
            str(bill_no).strip()
        )

        # Check dates already stored as datetime
        if isinstance(sale_date, (datetime, date)):

            if (
                sale_date.year == 2026
                and sale_date.month == 10
            ):
                october_sales_rows += 1

    # ======================================
    # 5. READ EXISTING RETURNS IDENTIFIERS
    # ======================================

    existing_return_codes = []

    existing_return_rows = 0

    for row in returns_ws.iter_rows(
        min_row=2,
        values_only=True
    ):

        # Column F = index 5
        return_code = row[5] if len(row) > 5 else None

        if return_code is None:
            continue

        existing_return_rows += 1

        existing_return_codes.append(
            str(return_code).strip()
        )

    # ======================================
    # 6. COMPARE AGAINST NEW DATA
    # ======================================

    source_sales_bills = set(
        sales["bill_no"].dropna().str.strip()
    )

    source_return_codes = set(
        returns["return_code"].dropna().str.strip()
    )

    sales_bill_overlap = (
        source_sales_bills
        & set(existing_sales_bills)
    )

    return_code_overlap = (
        source_return_codes
        & set(existing_return_codes)
    )

    # ======================================
    # 7. DISPLAY RESULTS
    # ======================================

    print("\n===== BE DAPPER SALES =====")

    print("New sales lines:", len(sales))
    print("Existing sales lines:", existing_sales_rows)
    print("October 2026 existing sales lines:", october_sales_rows)

    print(
        "Overlapping bill numbers:",
        len(sales_bill_overlap)
    )

    print("\n===== BE DAPPER RETURNS =====")

    print("New return lines:", len(returns))
    print("Existing return lines:", existing_return_rows)

    print(
        "Overlapping return codes:",
        len(return_code_overlap)
    )

    # ======================================
    # 8. INSPECT REPEATED IDENTIFIERS
    # ======================================

    bill_counts = Counter(existing_sales_bills)

    repeated_bills = sum(
        count > 1 for count in bill_counts.values()
    )

    print("\nRepeated existing bill numbers:", repeated_bills)

finally:
    wb.close()

print("\nExisting transaction check completed!")
