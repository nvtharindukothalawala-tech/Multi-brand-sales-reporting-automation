
from pathlib import Path
import pandas as pd

# 1. Project folder
BASE_DIR = Path(__file__).resolve().parent.parent

# 2. Excel file paths
sales_file = BASE_DIR / "data/raw/sales_reports_bill_wise_detail_xlsx_11_1_2026.xlsx"

returns_file = BASE_DIR / "data/raw/return_reports_customer_bill_wise_2026-11-01_11_05.xlsx"

# 3. Read the Excel files
sales_df = pd.read_excel(sales_file)
returns_df = pd.read_excel(returns_file)

# 4. Display column names
print("\n===== SALES COLUMNS =====")

for column in sales_df.columns:
    print(column)

print("\n===== RETURNS COLUMNS =====")

for column in returns_df.columns:
    print(column)

# 5. Identify column data types
print("\n===== SALES DATA TYPES =====")
print(sales_df.dtypes)

print("\n===== RETURNS DATA TYPES =====")
print(returns_df.dtypes)

# 6. Check missing values
print("\n===== SALES MISSING VALUES =====")
print(sales_df.isnull().sum())

print("\n===== RETURNS MISSING VALUES =====")
print(returns_df.isnull().sum())
