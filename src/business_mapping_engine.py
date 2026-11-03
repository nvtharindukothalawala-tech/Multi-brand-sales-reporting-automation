
from pathlib import Path
import sys
import pandas as pd

# 1. Define project folders
BASE_DIR = Path(__file__).resolve().parent.parent

DATA_DIR = BASE_DIR / "data" / "processed"
OUTPUT_DIR = BASE_DIR / "output" / "business_data"

OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

# 2. Import mapping configuration
sys.path.insert(0, str(BASE_DIR))

from config.business_mapping import (
    SALES_BUSINESS_MAP,
    RETURNS_BUSINESS_MAP,
    BUSINESSES,
)

# 3. Read cleaned datasets
sales_df = pd.read_csv(
    DATA_DIR / "sales_clean.csv",
    dtype={"bill_no": "string"},
)

returns_df = pd.read_csv(
    DATA_DIR / "returns_clean.csv",
    dtype={
        "bill_no": "string",
        "return_code": "string",
    },
)

# 4. Apply business mapping
sales_df["business"] = (
    sales_df["location"]
    .map(SALES_BUSINESS_MAP)
    .fillna("UNMAPPED")
)

returns_df["business"] = (
    returns_df["warehouse"]
    .map(RETURNS_BUSINESS_MAP)
    .fillna("UNMAPPED")
)

# 5. Export files for each business
for business in BUSINESSES:

    folder_name = business.lower().replace(" ", "_")
    business_folder = OUTPUT_DIR / folder_name

    business_folder.mkdir(
        parents=True,
        exist_ok=True
    )

    business_sales = sales_df[
        sales_df["business"] == business
    ]

    business_returns = returns_df[
        returns_df["business"] == business
    ]

    business_sales.to_csv(
        business_folder / "sales.csv",
        index=False
    )

    business_returns.to_csv(
        business_folder / "returns.csv",
        index=False
    )

    print(f"\nBusiness: {business}")
    print("Sales lines:", len(business_sales))
    print("Return lines:", len(business_returns))

# 6. Export unmapped records
unmapped_sales = sales_df[
    sales_df["business"] == "UNMAPPED"
]

unmapped_returns = returns_df[
    returns_df["business"] == "UNMAPPED"
]

unmapped_sales.to_csv(
    OUTPUT_DIR / "unmapped_sales.csv",
    index=False
)

unmapped_returns.to_csv(
    OUTPUT_DIR / "unmapped_returns.csv",
    index=False
)

# 7. Validate total rows
assert len(sales_df) == 6530
assert len(returns_df) == 304

print("\n===== RECONCILIATION =====")
print("Total sales:", len(sales_df))
print("Mapped sales:", len(sales_df) - len(unmapped_sales))
print("Unmapped sales:", len(unmapped_sales))

print("Total returns:", len(returns_df))
print("Mapped returns:", len(returns_df) - len(unmapped_returns))
print("Unmapped returns:", len(unmapped_returns))

print("\nBusiness mapping completed successfully!")
