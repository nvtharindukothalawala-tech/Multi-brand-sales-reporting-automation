
from pathlib import Path
import pandas as pd

# 1. Project folders
BASE_DIR = Path(__file__).resolve().parent.parent

DATA_DIR = BASE_DIR / "data" / "processed"
BUSINESS_DIR = BASE_DIR / "output" / "business_data"

# 2. Read the cleaned datasets
sales = pd.read_csv(DATA_DIR / "sales_clean.csv")
returns = pd.read_csv(DATA_DIR / "returns_clean.csv")

# 3. Business folder names
businesses = {
    "Be Desi": "be_desi",
    "Protecto": "protecto",
    "Giggles": "giggles",
    "Safety Sure": "safety_sure",
    "Be Dapper": "be_dapper"
}

# 4. Calculate original financial totals
original_sales = sales["sales_amount"].sum()
original_returns = returns["sale_amount"].sum()

original_sold_qty = sales["sold_qty"].sum()
original_return_qty = returns["return_qty"].sum()

print("===== ORIGINAL DATA TOTALS =====")
print("Sales Amount:", round(original_sales, 2))
print("Return Amount:", round(original_returns, 2))

# 5. Calculate each business total
results = []

for business, folder in businesses.items():

    business_sales = pd.read_csv(
        BUSINESS_DIR / folder / "sales.csv"
    )

    business_returns = pd.read_csv(
        BUSINESS_DIR / folder / "returns.csv"
    )

    results.append({
        "business": business,
        "sales_lines": len(business_sales),
        "return_lines": len(business_returns),
        "sales_amount": business_sales["sales_amount"].sum(),
        "return_amount": business_returns["sale_amount"].sum(),
        "sold_qty": business_sales["sold_qty"].sum(),
        "return_qty": business_returns["return_qty"].sum()
    })

summary = pd.DataFrame(results)

# 6. Read unmapped transactions
unmapped_sales = pd.read_csv(
    BUSINESS_DIR / "unmapped_sales.csv"
)

unmapped_returns = pd.read_csv(
    BUSINESS_DIR / "unmapped_returns.csv"
)

# 7. Recalculate totals
reconciled_sales = (
    summary["sales_amount"].sum()
    + unmapped_sales["sales_amount"].sum()
)

reconciled_returns = (
    summary["return_amount"].sum()
    + unmapped_returns["sale_amount"].sum()
)

reconciled_sold_qty = (
    summary["sold_qty"].sum()
    + unmapped_sales["sold_qty"].sum()
)

reconciled_return_qty = (
    summary["return_qty"].sum()
    + unmapped_returns["return_qty"].sum()
)

# 8. Check differences
sales_difference = original_sales - reconciled_sales
returns_difference = original_returns - reconciled_returns

print("\n===== RECONCILIATION =====")
print("Sales Difference:", round(sales_difference, 2))
print("Returns Difference:", round(returns_difference, 2))
print("Sold Qty Difference:", original_sold_qty - reconciled_sold_qty)
print("Return Qty Difference:", original_return_qty - reconciled_return_qty)

# 9. Financial validation
assert abs(sales_difference) < 0.01, "Sales totals mismatch"
assert abs(returns_difference) < 0.01, "Return totals mismatch"

assert original_sold_qty == reconciled_sold_qty
assert original_return_qty == reconciled_return_qty

# 10. Save the report
summary.to_csv(
    BASE_DIR / "output" / "financial_summary.csv",
    index=False
)

print("\nFinancial reconciliation PASSED!")
print(summary.to_string(index=False))
