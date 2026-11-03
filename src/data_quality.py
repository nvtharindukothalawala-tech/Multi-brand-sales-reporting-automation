
from pathlib import Path
import pandas as pd

# 1. Project folder
BASE_DIR = Path(__file__).resolve().parent.parent

# 2. File locations
sales_file = BASE_DIR / "data/raw/sales_reports_bill_wise_detail_xlsx_11_1_2026.xlsx"
returns_file = BASE_DIR / "data/raw/return_reports_customer_bill_wise_2026-11-01_11_05.xlsx"

# 3. Read data
# Protect bill numbers and contact numbers as text
sales_df = pd.read_excel(
    sales_file,
    dtype={
        "Bill No": "string",
        "Contact No1": "string",
        "Contact No2": "string"
    }
)

returns_df = pd.read_excel(
    returns_file,
    dtype={
        "Bill No": "string",
        "Return Code": "string"
    }
)

# 4. Define a reusable data-quality function
def data_quality_report(df, dataset_name):

    print(f"\n===== {dataset_name} =====")

    total_rows = len(df)

    print("Total Rows:", total_rows)
    print("Total Columns:", len(df.columns))

    # Identify missing values
    missing = df.isna().sum()

    # Calculate missing percentages
    missing_percentage = (
        missing / total_rows * 100
    ).round(2)

    # Create report table
    report = pd.DataFrame({
        "Column": df.columns,
        "Data Type": df.dtypes.astype(str).values,
        "Missing Count": missing.values,
        "Missing Percentage": missing_percentage.values
    })

    print("\nMissing Value Report:")
    print(report.to_string(index=False))

    # Check exact duplicate rows
    duplicate_rows = df.duplicated().sum()

    print("\nExact Duplicate Rows:", duplicate_rows)

    return report


# 5. Run the checks
sales_report = data_quality_report(
    sales_df, "SALES DATA"
)

returns_report = data_quality_report(
    returns_df, "RETURNS DATA"
)

# 6. Save the reports
output_folder = BASE_DIR / "output"
output_folder.mkdir(exist_ok=True)

sales_report.to_csv(
    output_folder / "sales_quality_report.csv",
    index=False
)

returns_report.to_csv(
    output_folder / "returns_quality_report.csv",
    index=False
)

print("\nData quality reports generated successfully!")
