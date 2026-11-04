"""Tests run on fictional workbooks created in pytest tmp_path only."""
from copy import deepcopy
from decimal import Decimal
from pathlib import Path
from hashlib import sha256
import json

import pytest
from openpyxl import Workbook, load_workbook

from src.business_schema import BUSINESS_SCHEMAS
from src.synthetic_monthly_batch import run_synthetic_monthly_batch

LABELS = {
    'sales': {'bill_no': 'Bill No', 'product': 'Product', 'sold_qty': 'Sold Qty',
              'sales_amount': 'Sales Amount', 'date': 'Date'},
    'returns': {'return_code': 'Return Code', 'bill_no': 'Bill No',
                'product': 'Product Description', 'return_qty': 'Return Qty',
                'sale_amount': 'Sale Amount', 'return_date': 'Return Date'},
}


def fake_template(path, business, marked=True):
    wb = Workbook()
    wb.active.title = 'SalesRepItem'
    wb.create_sheet('Ord.Rtn')
    if marked:
        wb['SalesRepItem']['A1'] = 'SYNTHETIC_TEST_ONLY'
    for kind in ('sales', 'returns'):
        schema = BUSINESS_SCHEMAS[business][kind]
        for field, label in LABELS[kind].items():
            wb[schema['sheet']][f"{schema['columns'][field]}{schema['header_row']}"] = label
    wb.save(path)
    wb.close()


def fake_jobs(path):
    jobs = {}
    for i, business in enumerate(sorted(BUSINESS_SCHEMAS)):
        source = path / f'synthetic_input_{i}.xlsx'
        fake_template(source, business)
        jobs[business] = {
            'source': source,
            'sales': [
                {'bill_no': f'SYN-{i}-A', 'product': 'Fictional Shirt', 'sold_qty': 2,
                 'sales_amount': '150.25', 'date': '2026-10-10'},
                {'bill_no': f'SYN-{i}-A', 'product': 'Fictional Shirt', 'sold_qty': 2,
                 'sales_amount': '150.25', 'date': '2026-10-10'},
            ],
            'returns': [{'return_code': f'RTN-{i}', 'bill_no': f'SYN-{i}-A',
                         'product': 'Fictional Shirt', 'return_qty': 1,
                         'sale_amount': '75.25', 'return_date': '2026-10-15'}],
        }
    return jobs


def output_files(path):
    return sorted(p for p in path.iterdir() if p.is_file())


def test_batch_five_business_financial_reconciliation(tmp_path):
    jobs = fake_jobs(tmp_path)
    signatures = {k: sha256(v['source'].read_bytes()).hexdigest() for k, v in jobs.items()}
    result = run_synthetic_monthly_batch(jobs, tmp_path, '2026-10')
    assert result['mode'] == 'SYNTHETIC_TEST_ONLY'
    assert result['totals']['sales_records'] == 10
    assert result['totals']['return_records'] == 5
    assert Decimal(result['totals']['sales_amount']) == Decimal('1502.50')
    assert Decimal(result['totals']['return_amount']) == Decimal('376.25')
    assert Decimal(result['totals']['requested_net_amount']) == Decimal('1126.25')
    assert len(result['businesses']) == 5
    assert json.loads((tmp_path / 'batch_summary.json').read_text()) == result
    for business, item in jobs.items():
        assert sha256(item['source'].read_bytes()).hexdigest() == signatures[business]
        rec = result['businesses'][business]
        assert rec['sales']['added'] == 2
        assert rec['returns']['added'] == 1
        output = tmp_path / rec['file']
        assert output.exists()
        wb = load_workbook(output, read_only=True)
        try:
            assert wb['SalesRepItem']['A1'].value == 'SYNTHETIC_TEST_ONLY'
            s = BUSINESS_SCHEMAS[business]['sales']
            r = BUSINESS_SCHEMAS[business]['returns']
            assert wb[s['sheet']][f"{s['columns']['bill_no']}{s['first_data_row']}"].value == item['sales'][0]['bill_no']
            assert wb[s['sheet']][f"{s['columns']['bill_no']}{s['first_data_row']+1}"].value == item['sales'][1]['bill_no']
            assert wb[r['sheet']][f"{r['columns']['return_code']}{r['first_data_row']}"].value == item['returns'][0]['return_code']
        finally:
            wb.close()


def test_preflight_wrong_month_no_outputs(tmp_path):
    jobs = fake_jobs(tmp_path)
    jobs['Be Desi']['sales'][0]['date'] = '2026-11-01'
    with pytest.raises(ValueError, match='outside reporting month'):
        run_synthetic_monthly_batch(jobs, tmp_path, '2026-10')
    assert not (tmp_path / 'batch_summary.json').exists()
    assert len(output_files(tmp_path)) == 5


def test_preflight_unmarked_template_no_outputs(tmp_path):
    jobs = fake_jobs(tmp_path)
    source = jobs['Safety Sure']['source']
    fake_template(source, 'Safety Sure', marked=False)
    with pytest.raises(PermissionError, match='unmarked'):
        run_synthetic_monthly_batch(jobs, tmp_path, '2026-10')
    assert len(output_files(tmp_path)) == 5


def test_preflight_output_collision_preserved(tmp_path):
    jobs = fake_jobs(tmp_path)
    existing = tmp_path / 'be_dapper_synthetic.xlsx'
    existing.write_bytes(b'KEEP THIS FILE')
    with pytest.raises(FileExistsError):
        run_synthetic_monthly_batch(jobs, tmp_path, '2026-10')
    assert existing.read_bytes() == b'KEEP THIS FILE'
    assert not (tmp_path / 'batch_summary.json').exists()


def test_staging_failure_does_not_publish_anything(tmp_path):
    jobs = fake_jobs(tmp_path)
    # Header mismatch bypasses basic preflight and is caught in the writer stage.
    source = jobs['Safety Sure']['source']
    wb = load_workbook(source)
    schema = BUSINESS_SCHEMAS['Safety Sure']['returns']
    wb[schema['sheet']][f"{schema['columns']['product']}{schema['header_row']}"] = 'WRONG'
    wb.save(source)
    wb.close()
    with pytest.raises(ValueError, match='Header mismatch'):
        run_synthetic_monthly_batch(jobs, tmp_path, '2026-10')
    assert len(output_files(tmp_path)) == 5


def test_existing_summary_is_never_overwritten(tmp_path):
    jobs = fake_jobs(tmp_path)
    summary = tmp_path / 'batch_summary.json'
    summary.write_text('OLD')
    with pytest.raises(FileExistsError):
        run_synthetic_monthly_batch(jobs, tmp_path, '2026-10')
    assert summary.read_text() == 'OLD'


def test_requires_all_five_businesses(tmp_path):
    jobs = fake_jobs(tmp_path)
    jobs.pop('Giggles')
    with pytest.raises(ValueError, match='exactly the five'):
        run_synthetic_monthly_batch(jobs, tmp_path, '2026-10')


@pytest.mark.parametrize('invalid_month', ['2026-13', '2026-1', 'October 2026'])
def test_invalid_month_is_blocked(tmp_path, invalid_month):
    jobs = fake_jobs(tmp_path)
    with pytest.raises(ValueError, match='[Mm]onth|reporting month'):
        run_synthetic_monthly_batch(jobs, tmp_path, invalid_month)


def test_duplicate_template_path_blocked(tmp_path):
    jobs = fake_jobs(tmp_path)
    jobs['Be Desi']['source'] = jobs['Be Dapper']['source']
    with pytest.raises(ValueError, match='independent synthetic source'):
        run_synthetic_monthly_batch(jobs, tmp_path, '2026-10')


def test_no_output_directory_is_created(tmp_path):
    jobs = fake_jobs(tmp_path)
    missing = tmp_path / 'nonexistent'
    with pytest.raises(ValueError, match='already exist'):
        run_synthetic_monthly_batch(jobs, missing, '2026-10')
    assert not missing.exists()
