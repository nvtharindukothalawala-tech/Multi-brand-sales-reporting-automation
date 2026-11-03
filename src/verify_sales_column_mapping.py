
from pathlib import Path
from openpyxl import load_workbook
import pandas as pd

# ==========================================
# 1. PROJECT PATHS
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
# 2. READ CLEANED SALES DATA
# ==========================================

sales = pd.read_csv(SALES_FILE)

print("\n===== SOURCE DATA TOTALS =====")

source_columns = [
    "sold_qty",
    "cost_amount",
    "sales_amount",
    "delivery_amount"
]

for col in source_columns:

    values = pd.to_numeric(
        sales[col], errors="coerce"
    )

    print(
        f"{col}: "
        f"Count={values.count()}, "
        f"Sum={values.sum():.2f}, "
        f"Min={values.min():.2f}, "
        f"Max={values.max():.2f}"
    )

# ==========================================
# 3. READ EXISTING EXCEL VALUES
# ==========================================

wb = load_workbook(
    TEMPLATE_FILE,
    read_only=True,
    data_only=False
)

try:

    ws = wb["SalesRepItem"]

    # Investigate columns around the shifted data
    columns = {
        "AT": 46,
        "AU": 47,
        "AV": 48,
        "AW": 49,
        "AX": 50,
        "AY": 51,
        "AZ": 52
    }

    print("\n===== EXISTING EXCEL COLUMN TOTALS =====")

    for letter, col_number in columns.items():

        values = []

        for row in ws.iter_rows(
            min_row=5,
            min_col=col_number,
            max_col=col_number,
            values_only=True
        ):

            values.append(row[0])

        numeric = pd.to_numeric(
            pd.Series(values),
            errors="coerce"
        )

        if numeric.count() > 0:

            print(
                f"{letter}: "
                f"Numeric Count={numeric.count()}, "
                f"Sum={numeric.sum():.2f}, "
                f"Min={numeric.min():.2f}, "
                f"Max={numeric.max():.2f}"
            )

        else:
            print(
                f"{letter}: Text/Non-numeric column"
            )

finally:
    wb.close()

print("\nColumn mapping verification completed!")
