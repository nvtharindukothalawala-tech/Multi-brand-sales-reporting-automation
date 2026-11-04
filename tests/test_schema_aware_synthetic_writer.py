"""Synthetic-only unit tests. Never reads or saves any business workbook."""
from datetime import date, datetime
from pathlib import Path

import pytest
from openpyxl import Workbook, load_workbook

from src.business_schema import BUSINESS_SCHEMAS
from src.schema_aware_synthetic_writer import write_synthetic_business_workbook

SALE = {'bill_no': 'TEST-001', 'product': 'Synthetic Shirt', 'sold_qty': 2,
        'sales_amount': 1250.50, 'date': '2026-10-12'}
RETURN = {'return_code': 'RET-01', 'bill_no': 'TEST-001',
          'product': 'Synthetic Shirt', 'return_qty': 1,
          'sale_amount': 625.25, 'return_date': '2026-10-15'}


def template(path, business, *, marker=True, kind='sales', existing=None, formula=False):
    config = BUSINESS_SCHEMAS[business]
    wb = Workbook()
    sales_ws = wb.active
    sales_ws.title = 'SalesRepItem'
    returns_ws = wb.create_sheet('Ord.Rtn')
    if marker:
        sales_ws['A1'] = 'SYNTHETIC_TEST_ONLY'
    header_names = {
        'sales': {'bill_no': 'Bill No', 'product': 'Product', 'sold_qty': 'Sold Qty',
                  'sales_amount': 'Sales Amount', 'date': 'Date'},
        'returns': {'return_code': 'Return Code', 'bill_no': 'Bill No',
                    'product': 'Product Description', 'return_qty': 'Return Qty',
                    'sale_amount': 'Sale Amount', 'return_date': 'Return Date'},
    }
    for section, ws in [('sales', sales_ws), ('returns', returns_ws)]:
        sub = config[section]
        for field, label in header_names[section].items():
            ws[f'{sub["columns"][field]}{sub["header_row"]}'] = label
    if existing:
        sub = config[kind]
        ws = wb[sub['sheet']]
        row = sub['first_data_row']
        for key, value in existing.items():
            if key in sub['columns']:
                ws[f'{sub["columns"][key]}{row}'] = date.fromisoformat(value) if key in ('date', 'return_date') else value
        if formula:
            ws[f'B{row}'] = f'=C{row}+1'
    wb.save(path)


@pytest.mark.parametrize('business', list(BUSINESS_SCHEMAS))
@pytest.mark.parametrize('kind,record', [('sales', SALE), ('returns', RETURN)])
def test_all_five_synthetic_layouts(tmp_path, business, kind, record):
    source, target = tmp_path/'input.xlsx', tmp_path/'output.xlsx'
    template(source, business)
    result = write_synthetic_business_workbook(source, target, business, kind, [record])
    assert result == {'existing': 0, 'requested': 1, 'added': 1}
    sub = BUSINESS_SCHEMAS[business][kind]
    with_target = load_workbook(target)
    ws = with_target[sub['sheet']]
    row = sub['first_data_row']
    for key, val in record.items():
        actual = ws[f'{sub["columns"][key]}{row}'].value
        assert (actual.date() if isinstance(actual, datetime) else actual) == (date.fromisoformat(val) if key in ('date', 'return_date') else val)
    assert with_target['SalesRepItem']['A1'].value == 'SYNTHETIC_TEST_ONLY'
    # Input must remain unmodified.
    source_book = load_workbook(source)
    assert source_book[sub['sheet']][f'{sub["columns"]["bill_no"]}{row}'].value is None
    with_target.close()
    source_book.close()


def test_existing_rows_are_not_repeated(tmp_path):
    source, target = tmp_path/'input.xlsx', tmp_path/'output.xlsx'
    template(source, 'Giggles', existing=SALE)
    result = write_synthetic_business_workbook(source, target, 'Giggles', 'sales', [SALE])
    assert result['added'] == 0
    assert result['existing'] == 1


def test_duplicate_occurrences_are_preserved(tmp_path):
    source, target = tmp_path/'input.xlsx', tmp_path/'output.xlsx'
    template(source, 'Be Desi', existing=SALE)
    result = write_synthetic_business_workbook(source, target, 'Be Desi', 'sales', [SALE, SALE, SALE])
    assert result['added'] == 2


def test_unmarked_workbook_blocked_and_no_output(tmp_path):
    source, target = tmp_path/'input.xlsx', tmp_path/'output.xlsx'
    template(source, 'Protecto', marker=False)
    with pytest.raises(PermissionError, match='SYNTHETIC_TEST_ONLY'):
        write_synthetic_business_workbook(source, target, 'Protecto', 'sales', [SALE])
    assert not target.exists()


def test_existing_output_never_overwritten(tmp_path):
    source, target = tmp_path/'input.xlsx', tmp_path/'output.xlsx'
    template(source, 'Safety Sure')
    target.write_bytes(b'original')
    with pytest.raises(ValueError, match='overwriting'):
        write_synthetic_business_workbook(source, target, 'Safety Sure', 'sales', [SALE])
    assert target.read_bytes() == b'original'


def test_header_mismatch_blocks_output(tmp_path):
    source, target = tmp_path/'input.xlsx', tmp_path/'output.xlsx'
    template(source, 'Be Dapper')
    wb = load_workbook(source)
    ws = wb['SalesRepItem']
    ws['AO4'] = 'Wrong Product Header'
    wb.save(source)
    wb.close()
    with pytest.raises(ValueError, match='Header mismatch'):
        write_synthetic_business_workbook(source, target, 'Be Dapper', 'sales', [SALE])
    assert not target.exists()


def test_formula_preserved_and_translated(tmp_path):
    source, target = tmp_path/'input.xlsx', tmp_path/'output.xlsx'
    template(source, 'Be Dapper', existing=SALE, formula=True)
    record = dict(SALE, bill_no='TEST-002')
    write_synthetic_business_workbook(source, target, 'Be Dapper', 'sales', [record])
    wb = load_workbook(target)
    start = BUSINESS_SCHEMAS['Be Dapper']['sales']['first_data_row']
    assert wb['SalesRepItem'][f'B{start+1}'].value == f'=C{start+1}+1'
    wb.close()


def test_wrong_date_and_nonfinite_amount_blocked(tmp_path):
    source, target = tmp_path/'input.xlsx', tmp_path/'output.xlsx'
    template(source, 'Giggles')
    for bad in [dict(SALE, date='12/10/2026'), dict(SALE, sales_amount=float('nan'))]:
        with pytest.raises(ValueError):
            write_synthetic_business_workbook(source, target, 'Giggles', 'sales', [bad])
    assert not target.exists()


def test_noncontiguous_rows_blocked(tmp_path):
    source, target = tmp_path/'input.xlsx', tmp_path/'output.xlsx'
    template(source, 'Giggles', existing=SALE)
    wb = load_workbook(source)
    ws = wb['SalesRepItem']
    first = BUSINESS_SCHEMAS['Giggles']['sales']['first_data_row']
    ws[f'A{first+2}'] = 'Unexpected footer'
    wb.save(source)
    wb.close()
    with pytest.raises(ValueError, match='Content after blank row'):
        write_synthetic_business_workbook(source, target, 'Giggles', 'sales', [SALE])
    assert not target.exists()


def test_formula_injection_is_blocked(tmp_path):
    source, target = tmp_path/'input.xlsx', tmp_path/'output.xlsx'
    template(source, 'Giggles')
    with pytest.raises(ValueError, match='Formula-like'):
        write_synthetic_business_workbook(source, target, 'Giggles', 'sales', [dict(SALE, product='=HYPERLINK("http://bad")')])
    assert not target.exists()
