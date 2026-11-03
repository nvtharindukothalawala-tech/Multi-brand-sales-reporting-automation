
from pathlib import Path
from collections import Counter

import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

BASE_DIR = Path(__file__).resolve().parent.parent

TEMPLATE_FILE = (
    BASE_DIR / "data" / "templates"
    / "BeDapper_ChannelWise_Report.xlsx"
)

SOURCE_FILE = (
    BASE_DIR / "output" / "business_data"
    / "be_dapper" / "sales.csv"
)

# Read only CSV column names
source_columns = pd.read_csv(
    SOURCE_FILE,
    nrows=0
).columns.tolist()

# Open workbook without modifying it
wb = load_workbook(
    TEMPLATE_FILE,
    read_only=True,
    data_only=False
)

try:
    ws = wb["SalesRepItem"]

    print("\n===== SOURCE CSV COLUMNS =====")

    for index, column in enumerate(
        source_columns, start=1
    ):
        print(index, column)

    print("\n===== EXCEL HEADER COMPARISON =====")

    for col_num in range(26, 54):

        letter = get_column_letter(col_num)

        row3 = ws.cell(
            row=3, column=col_num
        ).value

        row4 = ws.cell(
            row=4, column=col_num
        ).value

        print(
            f"{letter}: "
            f"Row 3 = {str(row3)[:35]} | "
            f"Row 4 = {str(row4)[:35]}"
        )

    print("\n===== DATA TYPE ANALYSIS =====")

    # Inspect first 100 data rows
    # Only report types, not transaction values
    for col_num in range(26, 54):

        letter = get_column_letter(col_num)

        types = Counter()

        for row in ws.iter_rows(
            min_row=5,
            max_row=min(ws.max_row, 104),
            min_col=col_num,
            max_col=col_num
        ):

            value = row[0].value

            if value is not None:
                types[type(value).__name__] += 1

        print(letter, dict(types))

finally:
    wb.close()

print("\nColumn alignment inspection completed!")
