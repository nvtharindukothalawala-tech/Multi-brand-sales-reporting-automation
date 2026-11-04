"""Five-business, full-25-field sales plus verified-core returns batch.

SYNTHETIC ONLY. Stages all workbooks, verifies, then exclusive-publishes.
Returns are limited to the six verified fields; real files are never modified.
"""
from __future__ import annotations

from datetime import date
from decimal import Decimal
from hashlib import sha256
from pathlib import Path
from tempfile import TemporaryDirectory
import json

from src.business_schema import BUSINESS_SCHEMAS
from src.full_field_source_mapping import normalize_sales_record
from src.full_field_synthetic_writer import _exclusive_copy, write_full_field_synthetic_sales
from src.schema_aware_synthetic_writer import write_synthetic_business_workbook
from src.synthetic_monthly_batch import _check_jobs, _filename, _totals


def run_full_field_synthetic_batch(jobs, output_dir, month):
    """Publish exactly five synthetic workbooks + batch_summary.json on success.

    Inputs follow existing batch coordinator's jobs structure. If normal
    validation/write fails, no output is published. A process crash during
    publication is not fully atomic and requires manual output inspection.
    """
    destination, checked = _check_jobs(jobs, output_dir, month)
    # Validate full-field data for all brands before staging starts.
    normalized = {}
    for business, job in checked.items():
        normalized[business] = {
            **job,
            'sales': [normalize_sales_record(row, month) for row in job['sales']],
        }
    stage = {}
    summary = {'mode': 'SYNTHETIC_TEST_ONLY', 'month': month,
               'sales_fields_verified': 25, 'returns_fields_verified': 6,
               'businesses': {}}
    with TemporaryDirectory(prefix='.full_field_batch_', dir=destination) as tmp:
        staging = Path(tmp)
        for business, job in normalized.items():
            sales_path = staging / f'{business.replace(" ", "_")}_sales.xlsx'
            final_path = staging / _filename(business)
            result_sales = write_full_field_synthetic_sales(
                job['source'], sales_path, business, job['sales'], month
            )
            result_returns = write_synthetic_business_workbook(
                sales_path, final_path, business, 'returns', job['returns']
            )
            sales_totals = _totals(job['sales'], 'sold_qty', 'sales_amount')
            return_totals = _totals(job['returns'], 'return_qty', 'sale_amount')
            summary['businesses'][business] = {
                'file': _filename(business),
                'sales': {**sales_totals, 'added': result_sales['added'],
                          'already_present': result_sales['requested'] - result_sales['added']},
                'returns': {**return_totals, 'added': result_returns['added'],
                            'already_present': result_returns['requested'] - result_returns['added']},
                'requested_net_amount': str(Decimal(sales_totals['amount']) - Decimal(return_totals['amount'])),
                'sha256': sha256(final_path.read_bytes()).hexdigest(),
            }
            stage[business] = final_path
        values = list(summary['businesses'].values())
        total_sales = sum((Decimal(i['sales']['amount']) for i in values), Decimal('0'))
        total_returns = sum((Decimal(i['returns']['amount']) for i in values), Decimal('0'))
        summary['totals'] = {
            'sales_records': sum(i['sales']['records'] for i in values),
            'return_records': sum(i['returns']['records'] for i in values),
            'sales_quantity': sum(i['sales']['quantity'] for i in values),
            'return_quantity': sum(i['returns']['quantity'] for i in values),
            'sales_amount': str(total_sales), 'return_amount': str(total_returns),
            'requested_net_amount': str(total_sales - total_returns),
        }
        summary_source = staging / 'batch_summary.json'
        summary_source.write_text(json.dumps(summary, indent=2) + '\n', encoding='utf-8')
        created = []
        try:
            for business in sorted(stage):
                target = destination / _filename(business)
                # This helper exclusively creates and cleans partial files.
                _exclusive_copy(stage[business], target)
                created.append(target)
            target_summary = destination / 'batch_summary.json'
            _exclusive_copy(summary_source, target_summary)
            created.append(target_summary)
        except BaseException:
            for path in reversed(created):
                path.unlink(missing_ok=True)
            raise
    return summary
