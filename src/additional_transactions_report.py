"""Read-only, confidential, downloadable listing of workbook-only transactions.

Does not edit source exports or ChannelWise Excel workbooks.
"""
from __future__ import annotations

import csv
from collections import Counter
from datetime import date, datetime
from decimal import Decimal
from hashlib import sha256
from io import BytesIO, StringIO
from pathlib import Path
from tempfile import TemporaryDirectory
from zipfile import ZipFile, ZIP_DEFLATED

from src import real_business_dry_run as audit_module
from src.reporting_ui_service import input_fingerprint, validate_month, _check_xlsx, _template_dir

FIELDS = (
    'Workbook business', 'Record type', 'Original worksheet', 'Scope classification',
    'Location / warehouse', 'Bill No', 'Return Code', 'Product', 'Quantity',
    'Sale amount (LKR)', 'Transaction date', 'Unit sales price', 'Cost amount',
    'Discount', 'Sales channel', 'Paid by',
)


def _find_additional(source: list[dict], workbook: list[dict], kind: str) -> list[dict]:
    """Multiset subtraction: preserve occurrence counts, including repeated bill numbers."""
    if kind not in ('sales', 'returns'):
        raise ValueError('Unknown transaction type')
    identity = audit_module._sale_identity if kind == 'sales' else audit_module._return_identity
    remaining = Counter(identity(row) for row in source)
    extras = []
    for row in workbook:
        key = identity(row)
        if remaining[key]:
            remaining[key] -= 1
        else:
            extras.append(row)
    return extras


def _safe(value):
    if value is None:
        return ''
    if isinstance(value, datetime):
        value = value.isoformat(sep=' ', timespec='seconds')
    elif isinstance(value, date):
        value = value.isoformat()
    if isinstance(value, Decimal):
        value = str(value)
    if isinstance(value, str) and value.lstrip().startswith(('=', '+', '-', '@', '\t', '\r')):
        value = "'" + value  # Prevent spreadsheet formula execution from imported text
    return value


def _csv(data: list[dict], headers: tuple[str, ...]) -> bytes:
    f = StringIO(newline='')
    w = csv.DictWriter(f, fieldnames=headers)
    w.writeheader()
    for row in data:
        w.writerow({name: _safe(row.get(name)) for name in headers})
    return ('\ufeff' + f.getvalue()).encode('utf-8')


def create_additional_transactions_report_zip(
    sales_bytes: bytes,
    returns_bytes: bytes,
    month: str,
    template_folder: str | Path,
    previously_reviewed: dict,
) -> bytes:
    """Build a private transaction-level report from reviewed, unchanged input files."""
    validate_month(month)
    _check_xlsx(sales_bytes, 'Sales')
    _check_xlsx(returns_bytes, 'Returns')
    root = _template_dir(template_folder)
    fingerprint = input_fingerprint(sales_bytes, returns_bytes, month, root)
    if fingerprint != previously_reviewed.get('ui_input_fingerprint'):
        raise ValueError('Inputs changed since last validation. Validate again.')
    with TemporaryDirectory(prefix='channelwise_additional_') as temp:
        base = Path(temp)
        sales_path = base / 'sales_uploaded.xlsx'
        returns_path = base / 'returns_uploaded.xlsx'
        sales_path.write_bytes(sales_bytes)
        returns_path.write_bytes(returns_bytes)
        validated = audit_module.audit(sales_path, returns_path, root, month)
        if validated.get('input_sha256') != previously_reviewed.get('input_sha256'):
            raise ValueError('Source workbooks changed since last validation. Validate again.')
        srows = audit_module._read_source(sales_path, 'sales', month)
        rrows = audit_module._read_source(returns_path, 'returns', month)
        sale_export = []
        return_export = []
        summary = []
        sales_by_location = {name: [r for r in srows if audit_module._txt(r['location']) in conf['sales_locations']] for name,conf in audit_module.BRANDS.items()}
        returns_by_warehouse = {name: [r for r in rrows if audit_module._txt(r['warehouse']) in conf['return_locations']] for name,conf in audit_module.BRANDS.items()}
        for name,conf in audit_module.BRANDS.items():
            wb_sales,wb_returns = audit_module._read_template(root / conf['file'], conf)
            for kind, source, wb_rows, allowed in (
                ('sales', sales_by_location[name], wb_sales, conf['sales_locations']),
                ('returns', returns_by_warehouse[name], wb_returns, conf['return_locations']),
            ):
                extras = _find_additional(source, wb_rows, kind)
                total = Decimal('0')
                for item in extras:
                    sale = kind == 'sales'
                    loc = audit_module._txt(item['location' if sale else 'warehouse'])
                    money = audit_module._num(item['sales_amount' if sale else 'sale_amount'])
                    total += money
                    details = {
                        'Workbook business': name,
                        'Record type': 'Sales' if sale else 'Returns',
                        'Original worksheet': 'SalesRepItem' if sale else 'Ord.Rtn',
                        'Scope classification': 'Other location - review' if loc not in allowed else 'Mapped location - review',
                        'Location / warehouse': loc,
                        'Bill No': audit_module._txt(item['bill_no']),
                        'Return Code': audit_module._txt(item.get('return_code')),
                        'Product': audit_module._txt(item['product']),
                        'Quantity': str(audit_module._num(item['sold_qty' if sale else 'return_qty'])),
                        'Sale amount (LKR)': str(money),
                        'Transaction date': item['date' if sale else 'return_date'],
                        'Unit sales price': str(audit_module._num(item['sales_price'])) if sale else '',
                        'Cost amount': str(audit_module._num(item['cost_amount'])) if sale else '',
                        'Discount': str(audit_module._num(item['discount'])) if sale else '',
                        'Sales channel': audit_module._txt(item['sales_channel']) if sale else '',
                        'Paid by': audit_module._txt(item['paid_by']) if sale else '',
                    }
                    (sale_export if sale else return_export).append(details)
                expected = validated['businesses'][name][kind]['additional_in_template']
                if len(extras) != expected:
                    raise ValueError('Additional record count changed during report generation')
                summary.append({'Business': name, 'Record type': 'Sales' if kind=='sales' else 'Returns',
                    'Additional rows': len(extras), 'Additional amount (LKR)': str(total)})
        output = BytesIO()
        with ZipFile(output, 'w', ZIP_DEFLATED) as bundle:
            bundle.writestr('additional_sales.csv', _csv(sale_export, FIELDS))
            bundle.writestr('additional_returns.csv', _csv(return_export, FIELDS))
            bundle.writestr('additional_summary.csv', _csv(summary, ('Business','Record type','Additional rows','Additional amount (LKR)')))
            bundle.writestr('READ_ME_FIRST.txt',
                'CONFIDENTIAL - REVIEW ONLY\n'
                'These are workbook rows unmatched to this workbook business\'s mapped source transactions.\n'
                'They can belong to other business locations and must NOT be deleted automatically.\n'
                'They are not necessarily missing from the source exports. No workbook was modified.\n'
                'Keep this ZIP private; it contains invoice numbers and financial details.\n')
        if not validated['inputs_unchanged_sha256_verified']:
            raise RuntimeError('Original inputs did not pass SHA-256 verification')
        return output.getvalue()
