
from pathlib import Path
import pandas as pd

# 1. Define project folders
BASE_DIR = Path(__file__).resolve().parent.parent
DATA_DIR = BASE_DIR / "data" / "processed"

# 2. Read cleaned datasets
sales_df = pd.read_csv(
    DATA_DIR / "sales_clean.csv",
    dtype={"bill_no": "string"}
)

returns_df = pd.read_csv(
    DATA_DIR / "returns_clean.csv",
    dtype={
        "bill_no": "string",
        "return_code": "string"
    }
)

# 3. Count sales locations
print("\n===== SALES LOCATIONS =====")

sales_locations = (
    sales_df["location"]
    .value_counts(dropna=False)
)

print(sales_locations.to_string())

# 4. Count returns warehouses
print("\n===== RETURNS WAREHOUSES =====")

return_warehouses = (
    returns_df["warehouse"]
    .value_counts(dropna=False)
)

print(return_warehouses.to_string())

# 5. Count returns outlets
print("\n===== RETURNS OUTLETS =====")

return_outlets = (
    returns_df["outlet"]
    .value_counts(dropna=False)
)

print(return_outlets.to_string())

# 6. Save mapping discovery files
output_dir = BASE_DIR / "output"
output_dir.mkdir(exist_ok=True)

sales_locations.rename_axis(
    "location"
).reset_index(
    name="transaction_lines"
).to_csv(
    output_dir / "sales_location_summary.csv",
    index=False
)

return_warehouses.rename_axis(
    "warehouse"
).reset_index(
    name="return_lines"
).to_csv(
    output_dir / "returns_warehouse_summary.csv",
    index=False
)

print("\nBusiness mapping analysis completed!")
