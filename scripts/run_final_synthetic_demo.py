"""One-command fictional five-business demo. Never reads real business files.

Run from repository root:
    .\\.venv\\Scripts\\python.exe -m scripts.run_final_synthetic_demo
Generated workbook outputs stay in output/synthetic_final_demo/ (do not commit).
"""
from __future__ import annotations

from datetime import date
from pathlib import Path
from tempfile import TemporaryDirectory
import json

from openpyxl import Workbook
from src.business_schema import BUSINESS_SCHEMAS
from src.full_field_source_mapping import SALES_FIELDS, SALES_HEADERS, sales_column_mapping
from src.full_field_synthetic_batch import run_full_field_synthetic_batch
from src.schema_aware_synthetic_writer import HEADERS, MARKER

MONTH = '2026-10'


def synthetic_sale(business, index):
    """All persons and transactions are invented test values."""
    row = dict.fromkeys(SALES_FIELDS)
    row.update({
        'bill_no': f'DEMO-{index:03}', 'location': f'Synthetic {business}',
        'customer': f'Example Customer {index}', 'address': 'Fictional Avenue',
        'created_by': 'DEMO', 'city': 'Demo City', 'sales_channel': 'Demo Retail',
        'note': 'TRAINING DATA', 'exchange_ref': None,
        'contact_no1': '0000000000', 'contact_no2': None,
        'sales_person': 'Demo Seller', 'invoice_type': 'Retail',
        'reference': f'EXAMPLE-{index}', 'delivery_amount': '20.00',
        'product': f'Fictional Product {index}', 'cost_price': '100.00',
        'sales_price': '150.00', 'discount': '0.00', 'exchange': '0.00',
        'sold_qty': 2, 'cost_amount': '200.00', 'sales_amount': '300.00',
        'paid_by': 'Demo Cash', 'date': date(2026, 10, 15),
    })
    return row


def synthetic_return(index):
    return {'return_code': f'RET-DEMO-{index:03}', 'bill_no': f'DEMO-{index:03}',
            'product': f'Fictional Product {index}', 'return_qty': 1,
            'sale_amount': '150.00', 'return_date': date(2026, 10, 16)}


def make_synthetic_template(destination, business):
    """Only creates NEW fictional fixtures. Refuses to replace an existing path."""
    if destination.exists():
        raise FileExistsError(destination)
    wb = Workbook()
    sales = wb.active
    sales.title = 'SalesRepItem'
    sales['A1'] = MARKER
    # 25 full fields in their actual verified row-4 positions.
    sales_schema = BUSINESS_SCHEMAS[business]['sales']
    sales_map = sales_column_mapping(business)
    for field, header in zip(SALES_FIELDS, SALES_HEADERS):
        sales[f'{sales_map[field]}{sales_schema["header_row"]}'] = header
    returns = wb.create_sheet('Ord.Rtn')
    return_schema = BUSINESS_SCHEMAS[business]['returns']
    for field, header in HEADERS['returns'].items():
        returns[f'{return_schema["columns"][field]}{return_schema["header_row"]}'] = header
    wb.create_sheet('Monthly Total')['A1'] = 'SYNTHETIC TRAINING REPORT'
    wb.save(destination)
    wb.close()


def run_demo(output_dir=Path('output') / 'synthetic_final_demo'):
    output_dir = Path(output_dir)
    if output_dir.exists() and any(output_dir.iterdir()):
        raise FileExistsError(f'Refusing to overwrite demo outputs: {output_dir}')
    output_dir.mkdir(parents=True, exist_ok=True)
    with TemporaryDirectory(prefix='synthetic_demo_sources_') as tmp:
        jobs = {}
        for index, business in enumerate(BUSINESS_SCHEMAS, 1):
            src = Path(tmp) / f'{index}_independent_synthetic.xlsx'
            make_synthetic_template(src, business)
            jobs[business] = {'source': src, 'sales': [synthetic_sale(business, index)],
                              'returns': [synthetic_return(index)]}
        summary = run_full_field_synthetic_batch(jobs, output_dir, MONTH)
    return summary


if __name__ == '__main__':
    result = run_demo()
    print('FINAL SYNTHETIC DEMO PASSED')
    print('Reports: 5 | Sales:', result['totals']['sales_records'],
          '| Returns:', result['totals']['return_records'])
    print('Net:', result['totals']['requested_net_amount'])
    print('Output: output/synthetic_final_demo')
