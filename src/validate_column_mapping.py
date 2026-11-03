
from pathlib import Path
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
import pandas as pd

# =====================================
# 1. DEFINE PROJECT PATHS
# =====================================

BASE_DIR = Path(__file__).resolve().parent.parent

DATA_DIR = BASE_DIR / "data" / "processed"

TEMPLATE_FILE = (
    BASE_DIR
    / "data"
    / "templates"
    / "BeDapper_ChannelWise_Report.xlsx"
)

# =====================================
# 2. LOAD CLEANED SOURCE DATA
# =====================================

sales = pd.read_csv(
    DATA_DIR / "sales_clean.csv",
    nrows=0
)

returns = pd.read_csv(
    DATA_DIR / "returns_clean.csv",
    nrows=0
)

# nrows=0 loads column names only,
# not actual customer transactions.

# =====================================
# 3. OPEN EXCEL TEMPLATE
# =====================================

workbook = load_workbook(
    TEMPLATE_FILE,
    read_only=True,
    data_only=False
)

# =====================================
# 4. NORMALIZE COLUMN NAMES
# =====================================

def normalize_name(name):

    return (
        str(name)
        .strip()
        .lower()
        .replace(" ", "_")
    )

# =====================================
# 5. COMPARE SOURCE AND TEMPLATE
# =====================================

def compare_columns(
    source_columns,
    sheet_name,
    header_row,
    start_column,
    end_column
):

    sheet = workbook[sheet_name]

    print("\n" + "=" * 55)
    print("SHEET:", sheet_name)
    print("=" * 55)

    source_set = set(source_columns)
    matched_sources = set()

    print("\nCOLUMN MAPPING RESULTS:")

    for column_number in range(
        start_column,
        end_column + 1
    ):

        cell = sheet.cell(
            row=header_row,
            column=column_number
        )

        if cell.value is None:
            continue

        excel_name = normalize_name(cell.value)
        column_letter = get_column_letter(
            column_number
        )

        if excel_name in source_set:

            print(
                f"{column_letter}: "
                f"{excel_name} -> MATCHED"
            )

            matched_sources.add(excel_name)

        else:

            print(
                f"{column_letter}: "
                f"{excel_name} -> REVIEW REQUIRED"
            )

    print("\nSOURCE COLUMNS NOT MATCHED:")

    for column in sorted(
        source_set - matched_sources
    ):
        print("-", column)


# =====================================
# 6. RUN VALIDATION
# =====================================

try:

    compare_columns(
        source_columns=sales.columns,
        sheet_name="SalesRepItem",
        header_row=3,
        start_column=26,  # Z
        end_column=53     # BA
    )

    compare_columns(
        source_columns=returns.columns,
        sheet_name="Ord.Rtn",
        header_row=1,
        start_column=5,   # E
        end_column=24     # X
    )

finally:
    workbook.close()

print("\nColumn mapping validation completed!")
