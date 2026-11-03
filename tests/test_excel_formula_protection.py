
import pytest

from openpyxl import Workbook, load_workbook
from openpyxl.styles import PatternFill, Font

from src.synthetic_excel_writer import (
    TEST_MARKER,
    write_synthetic_copy,
)


# ==========================================
# 1. CREATE FAKE TRANSACTION
# ==========================================

def fake_sale(bill_no):
    return {
        "bill_no": bill_no,
        "product": "TEST_SHIRT",
        "sold_qty": 1,
        "cost_price": 1000,
        "sales_price": 1500,
        "cost_amount": 1000,
        "sales_amount": 1500,
        "date": "2026-10-10",
    }


# ==========================================
# 2. CREATE FORMATTED TEST WORKBOOK
# ==========================================

def create_formatted_workbook(file_path):

    wb = Workbook()
    ws = wb.active
    ws.title = "SalesRepItem"

    # Required synthetic workbook marker
    ws["A1"] = TEST_MARKER

    # Existing calculated formula
    ws["B5"] = "=AT5*AQ5"

    # Existing cell formatting
    ws["B5"].font = Font(
        bold=True,
        color="FFFFFF"
    )

    ws["B5"].fill = PatternFill(
        fill_type="solid",
        fgColor="4472C4"
    )

    ws["AV5"].number_format = '#,##0.00'

    record = fake_sale("TEST001")

    columns = {
        "bill_no": "Z",
        "product": "AO",
        "cost_price": "AP",
        "sales_price": "AQ",
        "sold_qty": "AT",
        "cost_amount": "AU",
        "sales_amount": "AV",
        "date": "AZ",
    }

    for field, column in columns.items():
        ws[f"{column}5"] = record[field]

    wb.save(file_path)
    wb.close()


# ==========================================
# TEST 1: EXISTING FORMULA IS PRESERVED
# ==========================================

def test_existing_formula_is_preserved(tmp_path):

    original = tmp_path / "original.xlsx"
    output = tmp_path / "updated.xlsx"

    create_formatted_workbook(original)

    write_synthetic_copy(
        original,
        output,
        [fake_sale("TEST001"), fake_sale("TEST002")]
    )

    wb = load_workbook(output)

    try:
        ws = wb["SalesRepItem"]

        assert ws["B5"].value == "=AT5*AQ5"
        assert ws["B5"].font.bold is True
        assert ws["AV5"].number_format == '#,##0.00'

    finally:
        wb.close()


# ==========================================
# TEST 2: NEW ROW SHOULD GET FORMULA
# ==========================================

def test_new_row_formula_is_translated(tmp_path):

    original = tmp_path / "original.xlsx"
    output = tmp_path / "updated.xlsx"

    create_formatted_workbook(original)

    write_synthetic_copy(
        original,
        output,
        [fake_sale("TEST001"), fake_sale("TEST002")]
    )

    wb = load_workbook(output)

    try:
        ws = wb["SalesRepItem"]

        # New row should have adjusted references.
        assert ws["B6"].value == "=AT6*AQ6"

    finally:
        wb.close()


# ==========================================
# TEST 3: NEW ROW SHOULD GET FORMATTING
# ==========================================

def test_new_row_formatting_is_copied(tmp_path):

    original = tmp_path / "original.xlsx"
    output = tmp_path / "updated.xlsx"

    create_formatted_workbook(original)

    write_synthetic_copy(
        original,
        output,
        [fake_sale("TEST001"), fake_sale("TEST002")]
    )

    wb = load_workbook(output)

    try:
        ws = wb["SalesRepItem"]

        assert ws["B6"].font.bold is True
        assert (
            ws["B6"].fill.fgColor.rgb
            == ws["B5"].fill.fgColor.rgb
        )
        assert (
            ws["AV6"].number_format
            == ws["AV5"].number_format
        )

    finally:
        wb.close()


# ==========================================
# TEST 4: ABSOLUTE AND MIXED REFERENCES
# ==========================================

def test_absolute_and_mixed_references(tmp_path):

    original = tmp_path / "original.xlsx"
    output = tmp_path / "updated.xlsx"

    create_formatted_workbook(original)

    # Add a more complicated formula.
    wb = load_workbook(original)

    try:
        ws = wb["SalesRepItem"]

        ws["B5"] = "=AT5*$AQ$2+AQ$1+$AT5"

        wb.save(original)
    finally:
        wb.close()

    write_synthetic_copy(
        original,
        output,
        [
            fake_sale("TEST001"),
            fake_sale("TEST002"),
        ],
    )

    wb = load_workbook(output)

    try:
        ws = wb["SalesRepItem"]

        assert ws["B6"].value == (
            "=AT6*$AQ$2+AQ$1+$AT6"
        )

    finally:
        wb.close()


# ==========================================
# TEST 5: MULTIPLE NEW ROWS
# ==========================================

def test_formulas_translate_across_rows(tmp_path):

    original = tmp_path / "original.xlsx"
    output = tmp_path / "updated.xlsx"

    create_formatted_workbook(original)

    write_synthetic_copy(
        original,
        output,
        [
            fake_sale("TEST001"),
            fake_sale("TEST002"),
            fake_sale("TEST003"),
            fake_sale("TEST004"),
        ],
    )

    wb = load_workbook(output)

    try:
        ws = wb["SalesRepItem"]

        assert ws["B5"].value == "=AT5*AQ5"
        assert ws["B6"].value == "=AT6*AQ6"
        assert ws["B7"].value == "=AT7*AQ7"
        assert ws["B8"].value == "=AT8*AQ8"

        assert ws["Z8"].value == "TEST004"

    finally:
        wb.close()


# ==========================================
# TEST 6: PRESERVE OTHER WORKSHEETS
# ==========================================

def test_other_worksheets_are_preserved(tmp_path):

    original = tmp_path / "original.xlsx"
    output = tmp_path / "updated.xlsx"

    create_formatted_workbook(original)

    wb = load_workbook(original)

    try:
        summary = wb.create_sheet("Monthly Total")

        summary["A1"] = "SYNTHETIC SUMMARY"
        summary["B2"] = "=SUM(1,2,3)"

        wb.save(original)
    finally:
        wb.close()

    write_synthetic_copy(
        original,
        output,
        [
            fake_sale("TEST001"),
            fake_sale("TEST002"),
        ],
    )

    wb = load_workbook(output)

    try:
        assert wb.sheetnames == [
            "SalesRepItem",
            "Monthly Total",
        ]

        summary = wb["Monthly Total"]

        assert summary["A1"].value == (
            "SYNTHETIC SUMMARY"
        )

        assert summary["B2"].value == (
            "=SUM(1,2,3)"
        )

    finally:
        wb.close()
