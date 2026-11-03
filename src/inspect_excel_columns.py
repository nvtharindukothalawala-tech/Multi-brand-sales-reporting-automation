
from pathlib import Path
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

# ==========================================
# 1. DEFINE PROJECT PATHS
# ==========================================

BASE_DIR = Path(__file__).resolve().parent.parent

file_path = (
    BASE_DIR
    / "data"
    / "templates"
    / "BeDapper_ChannelWise_Report.xlsx"
)

# ==========================================
# 2. OPEN WORKBOOK SAFELY
# ==========================================

workbook = load_workbook(
    filename=file_path,
    read_only=True,
    data_only=False
)

# ==========================================
# 3. DEFINE WORKSHEETS
# ==========================================

sheet_names = [
    "SalesRepItem",
    "Ord.Rtn"
]

try:

    # ======================================
    # 4. INSPECT EACH WORKSHEET
    # ======================================

    for sheet_name in sheet_names:

        sheet = workbook[sheet_name]

        print("\n" + "=" * 60)
        print("SHEET:", sheet_name)
        print("=" * 60)

        # Header row location
        if sheet_name == "SalesRepItem":
            header_row = 3
        else:
            header_row = 1

        print("\nCOLUMN HEADERS:")

        # Read only the header row
        header_cells = sheet.iter_rows(
            min_row=header_row,
            max_row=header_row
        )

        for row in header_cells:

            # enumerate provides column positions
            for column_number, cell in enumerate(
                row,
                start=1
            ):

                if cell.value is not None:

                    column_letter = get_column_letter(
                        column_number
                    )

                    print(
                        f"{column_letter}: "
                        f"{str(cell.value)[:80]}"
                    )

        # ==================================
        # 5. INSPECT FORMULA LOCATIONS
        # ==================================

        print("\nFORMULA LOCATIONS (FIRST 10 ROWS):")

        formula_count = 0

        for row_number, row in enumerate(
            sheet.iter_rows(
                min_row=1,
                max_row=min(sheet.max_row or 0, 10)
            ),
            start=1
        ):

            for column_number, cell in enumerate(
                row,
                start=1
            ):

                if cell.data_type == "f":

                    formula_count += 1

                    column_letter = get_column_letter(
                        column_number
                    )

                    coordinate = (
                        f"{column_letter}{row_number}"
                    )

                    print(
                        f"Formula found at {coordinate}"
                    )

        print(
            "\nTotal formulas found:",
            formula_count
        )

finally:

    # Always close workbook
    workbook.close()

print("\nColumn inspection completed successfully!")
