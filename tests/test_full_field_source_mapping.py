from copy import deepcopy
from datetime import date, datetime
from decimal import Decimal
from pathlib import Path

import pytest
from openpyxl import Workbook

from src.business_schema import BUSINESS_SCHEMAS
from src.full_field_source_mapping import (
    SALES_FIELDS, SALES_HEADERS, FIRST_COL, normalize_sales_record,
    sales_cell_values, sales_column_mapping, read_synthetic_sales_export,
)


def example_sale():
    r = {field: None for field in SALES_FIELDS}
    r.update({
        'bill_no': 'SYN-BILL-001', 'location': 'Fictional Shop',
        'customer': 'Synthetic Customer', 'address': 'Fictional Road',
        'created_by': 'Test User', 'city': 'Sample City',
        'sales_channel': 'Retail', 'note': 'Synthetic note',
        'exchange_ref': 'TEST-EX-01', 'contact_no1': '0770000000',
        'contact_no2': '0700000000', 'sales_person': 'Test Agent',
        'invoice_type': 'Cash', 'reference': 'REF-01',
        'delivery_amount': Decimal('50.00'), 'product': 'TEST PRODUCT',
        'cost_price': '400.00', 'sales_price': '600.00',
        'discount': '20.00', 'exchange': '0', 'sold_qty': 2,
        'cost_amount': '800.00', 'sales_amount': '1180.00',
        'paid_by': 'Cash', 'date': '2026-10-15',
    })
    return r


def make_export(path, sales, *, headers=SALES_HEADERS, serials=None):
    wb = Workbook()
    ws = wb.active
    ws.append(headers)
    for row in sales:
        ws.append([float(value) if isinstance(value, Decimal) else value
                   for value in (list(row.values()) + [serials])])
    wb.save(path)
    wb.close()


@pytest.mark.parametrize('business', list(FIRST_COL))
def test_all_25_columns_match_known_positions(business):
    columns = sales_column_mapping(business)
    assert len(columns) == 25
    assert len(set(columns.values())) == 25
    assert columns['bill_no'] == FIRST_COL[business]
    for key, letter in BUSINESS_SCHEMAS[business]['sales']['columns'].items():
        assert columns[key] == letter


@pytest.mark.parametrize('business', list(FIRST_COL))
def test_all_fields_map_without_changing_original(business):
    original = example_sale()
    snapshot = deepcopy(original)
    mapped = sales_cell_values(original, business, '2026-10')
    assert original == snapshot
    assert len(mapped) == 25
    assert mapped[sales_column_mapping(business)['note']] == 'Synthetic note'
    assert mapped[sales_column_mapping(business)['exchange_ref']] == 'TEST-EX-01'
    assert mapped[sales_column_mapping(business)['sales_amount']] == Decimal('1180.00')
    assert mapped[sales_column_mapping(business)['date']] == date(2026, 10, 15)


def test_does_not_accept_unknown_business():
    with pytest.raises(ValueError, match='Unknown business'):
        sales_column_mapping('Unknown Shop')


def test_does_not_accept_missing_sales_fields():
    rec = example_sale()
    del rec['paid_by']
    with pytest.raises(ValueError, match='Missing sales fields'):
        normalize_sales_record(rec, '2026-10')


def test_rejects_outside_reporting_month():
    with pytest.raises(ValueError, match='outside reporting month'):
        normalize_sales_record(example_sale(), '2026-11')


def test_rejects_ambiguous_date_string():
    rec = example_sale()
    rec['date'] = '15/10/2026'
    with pytest.raises(ValueError, match='ISO'):
        normalize_sales_record(rec, '2026-10')


def test_rejects_nonfinite_amount():
    rec = example_sale()
    rec['sales_amount'] = float('nan')
    with pytest.raises(ValueError, match='Non-finite'):
        normalize_sales_record(rec, '2026-10')


@pytest.mark.parametrize('quantity', [0, -1, 1.5, 'bad'])
def test_invalid_quantity_blocked(quantity):
    rec = example_sale()
    rec['sold_qty'] = quantity
    with pytest.raises(ValueError):
        normalize_sales_record(rec, '2026-10')


def test_rejects_excel_formula_like_text():
    rec = example_sale()
    rec['note'] = '=HYPERLINK("http://evil", "Click")'
    with pytest.raises(ValueError, match='Formula-like'):
        normalize_sales_record(rec, '2026-10')


def test_preserves_phone_leading_zero():
    result = normalize_sales_record(example_sale(), '2026-10')
    assert result['contact_no1'] == '0770000000'


def test_accepts_date_and_datetime():
    rec = example_sale()
    rec['date'] = datetime(2026, 10, 15, 16, 40)
    assert normalize_sales_record(rec, '2026-10')['date'] == date(2026, 10, 15)


def test_reads_synthetic_export_in_read_only_mode(tmp_path):
    src = tmp_path / 'synthetic_sales.xlsx'
    make_export(src, [example_sale()])
    before = src.read_bytes()
    loaded = read_synthetic_sales_export(src, '2026-10')
    assert len(loaded) == 1
    assert len(loaded[0]) == 25
    assert loaded[0]['sales_amount'] == Decimal('1180')
    assert src.read_bytes() == before


def test_rejects_wrong_export_header(tmp_path):
    src = tmp_path / 'bad_header.xlsx'
    headers = list(SALES_HEADERS)
    headers[7] = 'Exchange Ref'
    make_export(src, [example_sale()], headers=headers)
    with pytest.raises(ValueError, match='headers'):
        read_synthetic_sales_export(src, '2026-10')


def test_export_is_never_written_on_rejected_record(tmp_path):
    src = tmp_path / 'synthetic_sales.xlsx'
    rec = example_sale()
    rec['date'] = '2026-11-01'
    make_export(src, [rec])
    before = src.read_bytes()
    with pytest.raises(ValueError, match='outside reporting month'):
        read_synthetic_sales_export(src, '2026-10')
    assert src.read_bytes() == before


def test_does_not_implicitly_open_private_workbooks(tmp_path):
    rec = example_sale()
    result = sales_cell_values(rec, 'Be Dapper', '2026-10')
    assert len(result) == 25
    assert not list(tmp_path.iterdir())


def test_rejects_wrong_month_format():
    with pytest.raises(ValueError, match='YYYY-MM'):
        normalize_sales_record(example_sale(), 'October 2026')


def test_rejects_records_after_empty_row(tmp_path):
    src = tmp_path / 'synthetic_sales.xlsx'
    wb = Workbook()
    ws = wb.active
    ws.append(SALES_HEADERS)
    ws.append([None] * 26)
    ws.append(list(example_sale().values()) + [None])
    wb.save(src)
    wb.close()
    with pytest.raises(ValueError, match='after blank row'):
        read_synthetic_sales_export(src, '2026-10')
