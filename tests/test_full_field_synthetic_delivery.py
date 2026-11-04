"""Synthetic integration tests; never use real company files."""
from __future__ import annotations

from copy import deepcopy
from pathlib import Path
import json

import pytest
from openpyxl import load_workbook

from src.business_schema import BUSINESS_SCHEMAS
from src.full_field_source_mapping import SALES_FIELDS, sales_column_mapping
from src.full_field_synthetic_writer import write_full_field_synthetic_sales
from src.full_field_synthetic_batch import run_full_field_synthetic_batch
from scripts.run_final_synthetic_demo import (
    make_synthetic_template, synthetic_sale, synthetic_return,
)


@pytest.mark.parametrize('business', list(BUSINESS_SCHEMAS))
def test_all_25_fields_saved_in_correct_columns(tmp_path, business):
    src, out = tmp_path/'synthetic.xlsx', tmp_path/'report.xlsx'
    make_synthetic_template(src, business)
    record = synthetic_sale(business, 1)
    result = write_full_field_synthetic_sales(src, out, business, [record], '2026-10')
    assert result['added'] == 1
    wb = load_workbook(out, read_only=True)
    try:
        ws = wb['SalesRepItem']
        row = BUSINESS_SCHEMAS[business]['sales']['first_data_row']
        mapping = sales_column_mapping(business)
        for field in SALES_FIELDS:
            cell = ws[f'{mapping[field]}{row}']
            assert cell.value is not None or record[field] is None, (business, field)
        assert ws[f'{mapping["customer"]}{row}'].value == 'Example Customer 1'
        assert ws[f'{mapping["sales_amount"]}{row}'].value == 300
        assert ws[f'{mapping["note"]}{row}'].value == 'TRAINING DATA'
        assert wb['Monthly Total']['A1'].value == 'SYNTHETIC TRAINING REPORT'
    finally:
        wb.close()
    assert src.exists()


@pytest.mark.parametrize('business', list(BUSINESS_SCHEMAS))
def test_repeat_full_record_is_idempotent(tmp_path, business):
    src = tmp_path/'src.xlsx'
    first = tmp_path/'first.xlsx'
    second = tmp_path/'second.xlsx'
    make_synthetic_template(src, business)
    record = synthetic_sale(business, 1)
    write_full_field_synthetic_sales(src, first, business, [record], '2026-10')
    result = write_full_field_synthetic_sales(first, second, business, [record], '2026-10')
    assert result['added'] == 0
    assert result['existing'] == 1


def test_conflicting_full_field_is_blocked(tmp_path):
    business = 'Be Dapper'
    src, initial, blocked = tmp_path/'src.xlsx', tmp_path/'first.xlsx', tmp_path/'blocked.xlsx'
    make_synthetic_template(src, business)
    record = synthetic_sale(business, 1)
    write_full_field_synthetic_sales(src, initial, business, [record], '2026-10')
    changed = deepcopy(record)
    changed['customer'] = 'Different Example Customer'
    with pytest.raises(ValueError, match='Full-field mismatch'):
        write_full_field_synthetic_sales(initial, blocked, business, [changed], '2026-10')
    assert not blocked.exists()


def test_repeated_identical_rows_preserve_occurrence_counts(tmp_path):
    business = 'Giggles'
    src, out = tmp_path/'src.xlsx', tmp_path/'out.xlsx'
    make_synthetic_template(src, business)
    record = synthetic_sale(business, 1)
    result = write_full_field_synthetic_sales(src, out, business, [record, record], '2026-10')
    assert result['added'] == 2
    verify = tmp_path/'verify.xlsx'
    assert write_full_field_synthetic_sales(out, verify, business, [record, record], '2026-10')['added'] == 0


def test_unmarked_workbook_is_blocked(tmp_path):
    business = 'Protecto'
    src, out = tmp_path/'src.xlsx', tmp_path/'out.xlsx'
    make_synthetic_template(src, business)
    from openpyxl import load_workbook
    wb = load_workbook(src)
    wb['SalesRepItem']['A1'] = None
    wb.save(src)
    wb.close()
    with pytest.raises(PermissionError):
        write_full_field_synthetic_sales(src, out, business, [synthetic_sale(business, 1)], '2026-10')
    assert not out.exists()


def test_existing_output_never_overwritten(tmp_path):
    business = 'Be Desi'
    src, out = tmp_path/'src.xlsx', tmp_path/'out.xlsx'
    make_synthetic_template(src, business)
    out.write_bytes(b'DO NOT CHANGE')
    with pytest.raises(ValueError):
        write_full_field_synthetic_sales(src, out, business, [synthetic_sale(business, 1)], '2026-10')
    assert out.read_bytes() == b'DO NOT CHANGE'


def test_wrong_month_is_blocked(tmp_path):
    src, out = tmp_path/'src.xlsx', tmp_path/'out.xlsx'
    make_synthetic_template(src, 'Be Dapper')
    with pytest.raises(ValueError):
        write_full_field_synthetic_sales(src, out, 'Be Dapper',
                                         [synthetic_sale('Be Dapper', 1)], '2026-09')
    assert not out.exists()


def test_complete_five_business_batch_and_totals(tmp_path):
    jobs = {}
    out = tmp_path/'out'
    out.mkdir()
    for index, business in enumerate(BUSINESS_SCHEMAS, 1):
        src = tmp_path/f'src{index}.xlsx'
        make_synthetic_template(src, business)
        jobs[business] = {'source': src, 'sales': [synthetic_sale(business, index)],
                          'returns': [synthetic_return(index)]}
    result = run_full_field_synthetic_batch(jobs, out, '2026-10')
    assert result['totals']['sales_records'] == 5
    assert result['totals']['return_records'] == 5
    assert result['totals']['requested_net_amount'] == '750.0' or float(result['totals']['requested_net_amount']) == 750
    assert len(list(out.glob('*.xlsx'))) == 5
    assert json.loads((out/'batch_summary.json').read_text()) == result


def test_batch_rejects_one_unmarked_source_without_publishing(tmp_path):
    jobs = {}
    out = tmp_path/'out'
    out.mkdir()
    for index, business in enumerate(BUSINESS_SCHEMAS, 1):
        src = tmp_path/f'src{index}.xlsx'
        make_synthetic_template(src, business)
        jobs[business] = {'source': src, 'sales': [synthetic_sale(business, index)],
                          'returns': [synthetic_return(index)]}
    target = jobs['Safety Sure']['source']
    wb = load_workbook(target)
    wb['SalesRepItem']['A1'] = 'NOT_SYNTHETIC'
    wb.save(target)
    wb.close()
    with pytest.raises(PermissionError):
        run_full_field_synthetic_batch(jobs, out, '2026-10')
    assert list(out.iterdir()) == []
