"""Transactional five-business coordinator for independently created synthetic workbooks.

Never use this module on real company workbooks or marked copies of real data.
No existing output is overwritten; a failed run removes only files it created.
"""
from __future__ import annotations

from datetime import date
from decimal import Decimal
from pathlib import Path
from tempfile import TemporaryDirectory
import json
import shutil

from openpyxl import load_workbook

from src.business_schema import BUSINESS_SCHEMAS
from src.schema_aware_synthetic_writer import (
    MARKER, _date, _identity, write_synthetic_business_workbook,
)


def _filename(business: str) -> str:
    return business.lower().replace(' ', '_') + '_synthetic.xlsx'


def _check_jobs(jobs, output_dir, month):
    if not isinstance(jobs, dict) or set(jobs) != set(BUSINESS_SCHEMAS):
        raise ValueError('Batch must contain exactly the five configured businesses')
    if not isinstance(month, str) or len(month) != 7 or month[4] != '-':
        raise ValueError('Month must use YYYY-MM')
    try:
        parsed_month = date.fromisoformat(month + '-01')
    except ValueError as exc:
        raise ValueError('Invalid reporting month') from exc
    if parsed_month.strftime('%Y-%m') != month:
        raise ValueError('Month must use YYYY-MM')
    destination = Path(output_dir)
    if not destination.is_dir():
        raise ValueError('Output directory must already exist')
    if (destination / 'batch_summary.json').exists():
        raise FileExistsError('Batch summary already exists: overwriting forbidden')
    sources = []
    normalized = {}
    for business in sorted(BUSINESS_SCHEMAS):
        job = jobs[business]
        if not isinstance(job, dict) or set(job) != {'source', 'sales', 'returns'}:
            raise ValueError(f'{business}: expected source, sales and returns')
        source = Path(job['source'])
        if not source.is_file() or source.suffix.lower() != '.xlsx':
            raise ValueError(f'{business}: source must be an existing .xlsx workbook')
        resolved = source.resolve()
        if resolved in sources:
            raise ValueError('Each business must have its own independent synthetic source')
        sources.append(resolved)
        target = (destination / _filename(business)).resolve()
        if target.exists():
            raise FileExistsError(f'{business}: output already exists')
        if target in sources or target == resolved:
            raise ValueError('An output path overlaps a source workbook')
        if not isinstance(job['sales'], (list, tuple)) or not isinstance(job['returns'], (list, tuple)):
            raise ValueError(f'{business}: transactions must be lists or tuples')
        wb = load_workbook(source, read_only=True, data_only=False)
        try:
            if 'SalesRepItem' not in wb or wb['SalesRepItem']['A1'].value != MARKER:
                raise PermissionError(f'{business}: unmarked workbook is blocked')
            if any(BUSINESS_SCHEMAS[business][kind]['sheet'] not in wb for kind in ('sales', 'returns')):
                raise ValueError(f'{business}: required sheet missing')
        finally:
            wb.close()
        for kind in ('sales', 'returns'):
            for record in job[kind]:
                _identity(record, kind)
                date_key = 'date' if kind == 'sales' else 'return_date'
                if _date(record[date_key])[:7] != month:
                    raise ValueError(f'{business}: {kind} transaction outside reporting month')
        normalized[business] = {'source': source, 'sales': list(job['sales']), 'returns': list(job['returns'])}
    targets = {(destination / _filename(business)).resolve() for business in BUSINESS_SCHEMAS}
    if any(source in targets for source in sources):
        raise ValueError('Source workbook cannot be another business output')
    return destination, normalized


def _totals(transactions, qty_key, amount_key):
    return {
        'records': len(transactions),
        'quantity': sum(int(Decimal(str(r[qty_key]))) for r in transactions),
        'amount': str(sum((Decimal(str(r[amount_key])) for r in transactions), Decimal('0'))),
    }


def _publish_new_file(source: Path, destination: Path):
    """Exclusive-create so concurrent users cannot overwrite existing reports."""
    with destination.open('xb') as out:
        with source.open('rb') as inp:
            shutil.copyfileobj(inp, out)


def run_synthetic_monthly_batch(jobs, output_dir, month):
    """Run all five brands; publish five xlsx files and one JSON only on success.

    jobs[business] = {'source': Path, 'sales': list[dict], 'returns': list[dict]}.
    The financial totals reflect *requested* transactions (not historical rows).
    The return value is the same data published in batch_summary.json.
    """
    destination, jobs = _check_jobs(jobs, output_dir, month)
    staged = {}
    summary = {'mode': 'SYNTHETIC_TEST_ONLY', 'month': month, 'businesses': {}}
    with TemporaryDirectory(prefix='.synthetic_batch_', dir=destination) as temp:
        stage_dir = Path(temp)
        # Stage and verify every workbook BEFORE publishing anything.
        for business, job in jobs.items():
            staged_sales = stage_dir / f'{business.replace(" ", "_")}_sales.xlsx'
            staged_final = stage_dir / _filename(business)
            sale_result = write_synthetic_business_workbook(
                job['source'], staged_sales, business, 'sales', job['sales']
            )
            return_result = write_synthetic_business_workbook(
                staged_sales, staged_final, business, 'returns', job['returns']
            )
            sales = _totals(job['sales'], 'sold_qty', 'sales_amount')
            returns = _totals(job['returns'], 'return_qty', 'sale_amount')
            summary['businesses'][business] = {
                'file': _filename(business),
                'sales': {**sales, 'added': sale_result['added'], 'already_present': sale_result['requested'] - sale_result['added']},
                'returns': {**returns, 'added': return_result['added'], 'already_present': return_result['requested'] - return_result['added']},
                'requested_net_amount': str(Decimal(sales['amount']) - Decimal(returns['amount'])),
            }
            staged[business] = staged_final
        summary['totals'] = {
            'sales_records': sum(s['sales']['records'] for s in summary['businesses'].values()),
            'return_records': sum(s['returns']['records'] for s in summary['businesses'].values()),
            'sales_amount': str(sum((Decimal(s['sales']['amount']) for s in summary['businesses'].values()), Decimal('0'))),
            'return_amount': str(sum((Decimal(s['returns']['amount']) for s in summary['businesses'].values()), Decimal('0'))),
        }
        summary['totals']['requested_net_amount'] = str(
            Decimal(summary['totals']['sales_amount']) - Decimal(summary['totals']['return_amount'])
        )
        staged_summary = stage_dir / 'batch_summary.json'
        staged_summary.write_text(json.dumps(summary, indent=2, ensure_ascii=False) + '\n', encoding='utf-8')

        created = []
        try:
            for business in sorted(staged):
                output = destination / _filename(business)
                # Register only after exclusive creation. If writing fails midway,
                # it is still OUR newly-created file, so it may be rolled back.
                with output.open('xb') as target:
                    created.append(output)
                    with staged[business].open('rb') as source:
                        shutil.copyfileobj(source, target)
            summary_path = destination / 'batch_summary.json'
            with summary_path.open('xb') as target:
                created.append(summary_path)
                with staged_summary.open('rb') as source:
                    shutil.copyfileobj(source, target)
        except BaseException:
            for path in reversed(created):
                path.unlink(missing_ok=True)
            raise
    return summary
