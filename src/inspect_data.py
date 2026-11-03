
from pathlib import Path
import pandas as pd

# Find the project directory
BASE_DIR = Path(__file__).resolve().parent.parent

# File locations
sales_file = BASE_DIR / "data/raw/sales_reports_bill_wise_detail_xlsx_11_1_2026.xlsx"

returns_file = BASE_DIR / "data/raw/return_reports_customer_bill_wise_2026-11-01_11_05.xlsx"

# Read both Excel files
sales_df = pd.read_excel(sales_file, dtype=str)
returns_df = pd.read_excel(returns_file, dtype=str)

# Display basic information
print("===== SALES DATA =====")
print("Total Rows:", sales_df.shape[0])
print("Total Columns:", sales_df.shape[1])

print("\n===== RETURNS DATA =====")
print("Total Rows:", returns_df.shape[0])
print("Total Columns:", returns_df.shape[1])
