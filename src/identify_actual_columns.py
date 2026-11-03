
from pathlib import Path
from collections import Counter
from datetime import datetime, date

import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils.datetime import from_excel

# ==========================================
# 1. PROJECT PATH
# ==========================================

BASE_DIR = Path(__file__).resolve().parent.parent

FILE_PATH = (
    BASE_DIR / "data" / "templates"
    / "BeDapper_ChannelWise_Report.xlsx"
)

# ==========================================
# 2. OPEN WORKBOOK
# ==========================================

wb = load_workbook(
    FILE_PATH,
    read_only=True,
    data_only=False
)

# ==========================================
# 3. SAFE DATE PARSING
# ==========================================

def parse_possible_date(value):

    if value is None or isinstance(value, bool):
        return pd.NaT

    if isinstance(value, (datetime, date)):
        return pd.Timestamp(value)

    # Numeric dates are only candidates if
    # they fall in a plausible Excel date range.
    if isinstance(value, (int, float)):

        if 40000 <= value <= 55000:
            try:
                return pd.Timestamp(
                    from_excel(value, epoch=wb.epoch)
                )
            except (ValueError, TypeError, OverflowError):
                return pd.NaT

        return pd.NaT

    if isinstance(value, str):

        value = value.strip()

        if not value:
            return pd.NaT

        # Check text dates using expected formats
        formats = [
            "%Y-%m-%d %H:%M:%S",
            "%Y-%m-%d",
            "%d/%m/%Y",
            "%d/%m/%Y %H:%M:%S",
        ]

        for fmt in formats:

            result = pd.to_datetime(
                value,
                format=fmt,
                errors="coerce"
            )

            if not pd.isna(result):
                return result

    return pd.NaT


# ==========================================
# 4. INSPECT CANDIDATE COLUMNS
# ==========================================

try:

    ws = wb["SalesRepItem"]

    columns = {
        "AW": 49,
        "AX": 50,
        "AY": 51,
        "AZ": 52,
    }

    print("\n===== COLUMN INVESTIGATION =====")

    for column_name, column_number in columns.items():

        types = Counter()
        date_count = 0
        october_count = 0
        text_count = 0

        # Start at row 5 because row 4
        # contains repeated headers.
        for row in ws.iter_rows(
            min_row=5,
            min_col=column_number,
            max_col=column_number,
            values_only=True
        ):

            value = row[0]

            if value is None:
                continue

            types[type(value).__name__] += 1

            if isinstance(value, str):
                text_count += 1

            parsed = parse_possible_date(value)

            if not pd.isna(parsed):

                if 2020 <= parsed.year <= 2030:
                    date_count += 1

                if (
                    parsed.year == 2026
                    and parsed.month == 10
                ):
                    october_count += 1

        print("\nColumn:", column_name)
        print("Data types:", dict(types))
        print("Text values:", text_count)
        print("Plausible dates:", date_count)
        print("October 2026 dates:", october_count)

finally:
    wb.close()

print("\nColumn investigation completed!")
