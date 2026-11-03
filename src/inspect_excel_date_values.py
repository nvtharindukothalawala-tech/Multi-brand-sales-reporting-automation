
from pathlib import Path
from collections import Counter
from openpyxl import load_workbook

# 1. Project path
BASE_DIR = Path(__file__).resolve().parent.parent

FILE_PATH = (
    BASE_DIR
    / "data"
    / "templates"
    / "BeDapper_ChannelWise_Report.xlsx"
)

# 2. Open the workbook
wb = load_workbook(
    FILE_PATH,
    read_only=True,
    data_only=False
)

try:
    ws = wb["SalesRepItem"]

    print("===== EXCEL DATE INVESTIGATION =====")
    print("Workbook epoch:", wb.epoch)
    print("AX header:", ws["AX3"].value)

    value_types = Counter()
    number_formats = Counter()
    numeric_values = []
    samples = []

    # Inspect the existing transaction date column
    for row in ws.iter_rows(
        min_row=4,
        min_col=50,
        max_col=50
    ):
        cell = row[0]
        value = cell.value

        if value is None:
            continue

        value_types[type(value).__name__] += 1
        number_formats[cell.number_format] += 1

        if isinstance(value, (int, float)) and not isinstance(value, bool):
            numeric_values.append(value)

        # Only sample the first 12 values
        if len(samples) < 12:
            samples.append((
                cell.row,
                repr(value),
                type(value).__name__,
                cell.number_format
            ))

    print("\n===== DATA TYPES =====")
    for value_type, count in value_types.items():
        print(value_type, ":", count)

    print("\n===== NUMBER FORMATS =====")
    for number_format, count in number_formats.items():
        print(repr(number_format), ":", count)

    print("\n===== NUMERIC VALUE RANGE =====")
    if numeric_values:
        print("Minimum:", min(numeric_values))
        print("Maximum:", max(numeric_values))

    print("\n===== FIRST 12 DATE CELLS =====")
    for row_num, value, data_type, number_format in samples:
        print(
            f"Row {row_num}: "
            f"Value={value}, "
            f"Type={data_type}, "
            f"Format={number_format}"
        )

finally:
    wb.close()

print("\nDate investigation completed!")
