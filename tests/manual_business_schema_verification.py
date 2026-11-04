"""Read-only manual verification of real workbook *headers* against schema.

Run from project root:
    .\\.venv\\Scripts\\python.exe tests/manual_business_schema_verification.py

Prints only business names and header-validation results, no transaction data.
Does not save, overwrite, or change any workbook.
"""
from pathlib import Path
import sys

# Allow direct execution as tests/manual_business_schema_verification.py
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from openpyxl import load_workbook
from src.business_schema import BUSINESS_SCHEMAS

BASE = Path(r"C:\Users\Tharindu\Desktop\Ass")
FILES = {
    "Be Dapper": "BeDapper_ChannelWise_Report.xlsx",
    "Be Desi": "BeDesi_ChannelWise_Report.xlsx",
    "Giggles": "Giggles_ChannelWise_Report.xlsx",
    "Protecto": "Protecto_ChannelWise_Report.xlsx",
    "Safety Sure": "SafetySure_ChannelWise_Report.xlsx",
}

# Schema uses 'product' for both Product and Product Description.
LABELS = {
    "sales": {
        "bill_no": "Bill No", "location": "Location", "product": "Product",
        "sold_qty": "Sold Qty", "sales_amount": "Sales Amount",
        "date": "Date", "note": "Note", "exchange_ref": "Exchange Ref",
    },
    "returns": {
        "return_code": "Return Code", "bill_no": "Bill No",
        "product": "Product Description", "return_qty": "Return Qty",
        "sale_amount": "Sale Amount", "return_date": "Return Date",
    },
}


def clean(value):
    return " ".join(str(value or "").strip().casefold().split())


def main():
    overall_pass = True
    for business, filename in FILES.items():
        path = BASE / filename
        print(f"\n=== {business} ===")
        if not path.is_file():
            print("REVIEW: workbook not found")
            overall_pass = False
            continue
        try:
            workbook = load_workbook(path, read_only=True, data_only=False)
        except Exception as exc:
            print(f"REVIEW: cannot open workbook ({type(exc).__name__})")
            overall_pass = False
            continue
        try:
            for section in ("sales", "returns"):
                schema = BUSINESS_SCHEMAS[business][section]
                sheet = schema["sheet"]
                if sheet not in workbook.sheetnames:
                    print(f"REVIEW: missing sheet {sheet}")
                    overall_pass = False
                    continue
                ws = workbook[sheet]
                mismatches = []
                for field, label in LABELS[section].items():
                    letter = schema["columns"][field]
                    coordinate = f"{letter}{schema['header_row']}"
                    if clean(ws[coordinate].value) != clean(label):
                        mismatches.append((field, coordinate))
                if mismatches:
                    print(f"{section.upper()}: REVIEW ({len(mismatches)} header mismatch(es))")
                    for field, coordinate in mismatches:
                        print(f"  {field}: expected label at {coordinate}")
                    overall_pass = False
                else:
                    print(f"{section.upper()}: PASS ({len(LABELS[section])} headers)")
        finally:
            workbook.close()
    print("\nNo original workbooks were modified.")
    print("OVERALL:", "PASS" if overall_pass else "REVIEW REQUIRED")
    raise SystemExit(0 if overall_pass else 1)


if __name__ == "__main__":
    main()
