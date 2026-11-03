
from pathlib import Path
import pandas as pd

# 1. Define folders
BASE_DIR = Path(__file__).resolve().parent.parent
RAW_DIR = BASE_DIR / "data" / "raw"
PROCESSED_DIR = BASE_DIR / "data" / "processed"

PROCESSED_DIR.mkdir(parents=True, exist_ok=True)

# 2. Define input files
sales_file = RAW_DIR / "sales_reports_bill_wise_detail_xlsx_11_1_2026.xlsx"

returns_file = RAW_DIR / "return_reports_customer_bill_wise_2026-11-01_11_05.xlsx"

# 3. Read source files
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

# 4. Standardize column names
def clean_column_names(df):
    df = df.copy()

    df.columns = (
        df.columns
        .str.strip()
        .str.lower()
        .str.replace(" ", "_", regex=False)
    )

    return df


sales_clean = clean_column_names(sales_df)
returns_clean = clean_column_names(returns_df)

# 5. Clean text fields without changing original data
def clean_text_columns(df):
    df = df.copy()

    for column in df.select_dtypes(
        include=["object", "string"]
    ).columns:

        df[column] = df[column].astype("string").str.strip()

    return df


sales_clean = clean_text_columns(sales_clean)
returns_clean = clean_text_columns(returns_clean)

# 6. Convert dates
# dayfirst=True is provisional; validate the raw format
# before using these dates in financial reporting.

# Convert sales dates using the correct format
sales_clean["date"] = pd.to_datetime(
    sales_clean["date"],
    format="%Y-%m-%d %H:%M:%S",
    errors="coerce"
)


returns_clean["return_date"] = pd.to_datetime(
    returns_clean["return_date"],
    errors="coerce"
)

# 7. Check invalid dates
sales_invalid_dates = sales_clean["date"].isna().sum()
returns_invalid_dates = returns_clean["return_date"].isna().sum()

print("Invalid sales dates:", sales_invalid_dates)
print("Invalid return dates:", returns_invalid_dates)

# 8. Only export if date conversion passed
if sales_invalid_dates > 0 or returns_invalid_dates > 0:
    raise ValueError(
        "Invalid dates found. Review the source date format."
    )

# 9. Save cleaned files
sales_clean.to_csv(
    PROCESSED_DIR / "sales_clean.csv",
    index=False
)

returns_clean.to_csv(
    PROCESSED_DIR / "returns_clean.csv",
    index=False
)

print("\nData cleaning completed successfully!")
print("Sales rows:", len(sales_clean))
print("Returns rows:", len(returns_clean))


# Verify the reporting period
print("\n===== DATE VALIDATION =====")

print("Sales Start Date:", sales_clean["date"].min())
print("Sales End Date:", sales_clean["date"].max())

print("Returns Start Date:", returns_clean["return_date"].min())
print("Returns End Date:", returns_clean["return_date"].max())
