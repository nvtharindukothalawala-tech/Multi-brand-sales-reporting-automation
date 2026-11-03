
from pathlib import Path
import pandas as pd

BASE_DIR = Path(__file__).resolve().parent.parent

review_file = (
    BASE_DIR
    / "config"
    / "reporting_scope_review.csv"
)

review = pd.read_csv(
    review_file,
    dtype="string"
)

required_columns = {
    "location",
    "proposed_reporting_business",
    "status"
}

if not required_columns.issubset(review.columns):
    raise ValueError("Missing required mapping columns")

print("\n===== REPORTING SCOPE REVIEW =====")

print("Total proposed mappings:", len(review))

print("\nMapping statuses:")
print(review["status"].value_counts())

approved = review[
    review["status"] == "APPROVED"
]

pending = review[
    review["status"] != "APPROVED"
]

print("\nApproved mappings:", len(approved))
print("Pending mappings:", len(pending))

if len(pending) > 0:
    print("\nWARNING: Reporting mappings require approval.")

print("\nMapping review completed!")
