
from pathlib import Path
from openpyxl import load_workbook
from zipfile import is_zipfile, BadZipFile

# ==========================================
# 1. DEFINE PROJECT FOLDERS
# ==========================================

BASE_DIR = Path(__file__).resolve().parent.parent

TEMPLATE_DIR = BASE_DIR / "data" / "templates"


# ==========================================
# 2. DEFINE FIVE BUSINESS FILES
# ==========================================

business_files = {
    "Be Dapper": "BeDapper_ChannelWise_Report.xlsx",
    "Be Desi": "BeDesi_ChannelWise_Report.xlsx",
    "Giggles": "Giggles_ChannelWise_Report.xlsx",
    "Protecto": "Protecto_ChannelWise_Report.xlsx",
    "Safety Sure": "SafetySure_ChannelWise_Report.xlsx",
}


# ==========================================
# 3. CHECK PROJECT PATHS
# ==========================================

print("\n===== FILE PATH CHECK =====")

print("Project Folder:", BASE_DIR)
print("Templates Folder:", TEMPLATE_DIR)

print("\nDoes templates folder exist?")
print(TEMPLATE_DIR.is_dir())


# Stop if the templates folder does not exist
if not TEMPLATE_DIR.is_dir():
    raise FileNotFoundError(
        "Templates folder not found. "
        "Create data/templates inside your project."
    )


print("\nExcel files found:")

for file in sorted(TEMPLATE_DIR.glob("*.xlsx")):
    print(file.name)


# ==========================================
# 4. INSPECT ALL FIVE BUSINESS WORKBOOKS
# ==========================================

print("\n===== EXCEL TEMPLATE INSPECTION =====")

valid_files = 0
invalid_files = 0

for business, filename in business_files.items():

    file_path = TEMPLATE_DIR / filename

    print("\n" + "=" * 50)
    print("BUSINESS:", business)
    print("FILE:", filename)

    # --------------------------------------
    # CHECK 1: FILE EXISTS
    # --------------------------------------

    if not file_path.is_file():
        print("ERROR: Excel file not found!")
        invalid_files += 1
        continue

    # --------------------------------------
    # CHECK 2: FILE SIZE
    # --------------------------------------

    file_size = file_path.stat().st_size

    print("File Size:", file_size, "bytes")

    if file_size == 0:
        print("ERROR: Excel file is empty (0 bytes).")
        invalid_files += 1
        continue

    # --------------------------------------
    # CHECK 3: VALID ZIP STRUCTURE
    # --------------------------------------

    if not is_zipfile(file_path):
        print("ERROR: Invalid XLSX ZIP structure.")
        invalid_files += 1
        continue

    # --------------------------------------
    # CHECK 4: OPEN EXCEL WORKBOOK
    # --------------------------------------

    try:

        # Read-only mode protects the original
        workbook = load_workbook(
            filename=file_path,
            read_only=True,
            data_only=False,
            keep_links=False
        )

        print("\nSTATUS: Workbook opened successfully!")
        print("Total Worksheets:", len(workbook.worksheets))

        print("\nWORKSHEETS:")

        # ----------------------------------
        # 5. INSPECT EACH WORKSHEET
        # ----------------------------------

        for sheet in workbook.worksheets:

            print("\n" + "-" * 40)

            print("Sheet Name:", sheet.title)
            print("Total Rows:", sheet.max_row)
            print("Total Columns:", sheet.max_column)

            print(
                "\nFirst 20 rows - "
                "populated cell counts:"
            )

            max_inspect_row = min(sheet.max_row or 0, 20)

            if max_inspect_row > 0:

                for row_number, row in enumerate(
                    sheet.iter_rows(
                        min_row=1,
                        max_row=max_inspect_row
                    ),
                    start=1
                ):

                    populated = sum(
                        cell.value is not None
                        for cell in row
                    )

                    if populated > 0:

                        print(
                            f"Row {row_number}: "
                            f"{populated} populated cells"
                        )

        valid_files += 1

    # --------------------------------------
    # HANDLE INVALID WORKBOOK ERRORS
    # --------------------------------------

    except (BadZipFile, OSError, ValueError, KeyError) as error:

        print("\nERROR: Unable to read workbook.")
        print("Details:", error)

        invalid_files += 1

    finally:

        # Close workbook if it was opened
        if "workbook" in locals():
            workbook.close()
            del workbook


# ==========================================
# 6. FINAL VALIDATION SUMMARY
# ==========================================

print("\n" + "=" * 50)
print("FINAL EXCEL VALIDATION SUMMARY")
print("=" * 50)

print("Total Business Files:", len(business_files))
print("Valid Files:", valid_files)
print("Invalid Files:", invalid_files)

if invalid_files == 0:

    print("\nSUCCESS!")
    print("All five business workbooks are valid.")

else:

    print("\nATTENTION!")
    print(
        invalid_files,
        "business workbook(s) need correction."
    )

print("\nExcel template inspection completed!")
