"""Synthetic-only structural tests for the five business schemas.

No company files are loaded or modified.
"""
import re
from collections import Counter

import pytest
from openpyxl import Workbook
from openpyxl.utils.cell import column_index_from_string

from src.business_schema import BUSINESS_SCHEMAS


EXPECTED_BUSINESSES = {
    "Be Dapper", "Be Desi", "Giggles", "Protecto", "Safety Sure"
}
SALES_FIELDS = {
    "bill_no", "location", "product", "sold_qty", "sales_amount",
    "date", "note", "exchange_ref",
}
RETURN_FIELDS = {
    "return_code", "bill_no", "product", "return_qty",
    "sale_amount", "return_date",
}
EXPECTED_START_COLUMNS = {
    "Be Dapper": ("Z", "G"),
    "Be Desi": ("AB", "F"),
    "Giggles": ("W", "F"),
    "Protecto": ("Y", "F"),
    "Safety Sure": ("Y", "F"),
}


def _check_schema(schema, expected_sheet, expected_fields):
    assert schema["sheet"] == expected_sheet
    assert isinstance(schema["header_row"], int)
    assert isinstance(schema["first_data_row"], int)
    assert schema["header_row"] >= 1
    assert schema["first_data_row"] == schema["header_row"] + 1

    columns = schema["columns"]
    assert expected_fields <= set(columns), (
        f"Missing fields: {expected_fields - set(columns)}"
    )
    selected = [columns[field] for field in expected_fields]
    assert len(selected) == len(set(selected)), "Duplicate destination columns"
    for letter in selected:
        assert isinstance(letter, str)
        assert re.fullmatch(r"[A-Z]{1,3}", letter)
        assert 1 <= column_index_from_string(letter) <= 16384


def test_exactly_five_businesses_configured():
    assert set(BUSINESS_SCHEMAS) == EXPECTED_BUSINESSES


@pytest.mark.parametrize("business", sorted(EXPECTED_BUSINESSES))
def test_sales_schema_structure(business):
    _check_schema(
        BUSINESS_SCHEMAS[business]["sales"],
        "SalesRepItem",
        SALES_FIELDS,
    )


@pytest.mark.parametrize("business", sorted(EXPECTED_BUSINESSES))
def test_returns_schema_structure(business):
    _check_schema(
        BUSINESS_SCHEMAS[business]["returns"],
        "Ord.Rtn",
        RETURN_FIELDS,
    )


def test_known_header_layouts_are_distinct():
    for business, (bill_column, return_code_column) in EXPECTED_START_COLUMNS.items():
        assert BUSINESS_SCHEMAS[business]["sales"]["columns"]["bill_no"] == bill_column
        assert BUSINESS_SCHEMAS[business]["returns"]["columns"]["return_code"] == return_code_column


def test_notes_and_exchange_refs_use_separate_columns():
    for business in EXPECTED_BUSINESSES:
        columns = BUSINESS_SCHEMAS[business]["sales"]["columns"]
        assert columns["note"] != columns["exchange_ref"]


def test_no_real_workbook_is_required_for_schema_validation(tmp_path):
    """Build a synthetic workbook in memory; never read business files."""
    for business in EXPECTED_BUSINESSES:
        workbook = Workbook()
        worksheet = workbook.active
        worksheet.title = "SalesRepItem"
        schema = BUSINESS_SCHEMAS[business]["sales"]
        for field, letter in schema["columns"].items():
            worksheet[f"{letter}{schema['header_row']}"] = field
        for field, letter in schema["columns"].items():
            assert worksheet[f"{letter}{schema['header_row']}"].value == field
        workbook.close()
