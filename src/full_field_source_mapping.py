"""Read-only source ingestion and mapping for synthetic sales export records.

No workbook saves. No business files are shipped. This is deliberately NOT
wired to the existing synthetic writer: it currently writes only core keys.
"""
from __future__ import annotations

from datetime import date, datetime
from decimal import Decimal, InvalidOperation
from pathlib import Path
import re

from openpyxl import load_workbook
from openpyxl.utils.cell import column_index_from_string
from src.business_schema import BUSINESS_SCHEMAS

SALES_FIELDS = (
    'bill_no', 'location', 'customer', 'address', 'created_by', 'city',
    'sales_channel', 'note', 'exchange_ref', 'contact_no1', 'contact_no2',
    'sales_person', 'invoice_type', 'reference', 'delivery_amount', 'product',
    'cost_price', 'sales_price', 'discount', 'exchange', 'sold_qty',
    'cost_amount', 'sales_amount', 'paid_by', 'date',
)
SALES_HEADERS = (
    'Bill No', 'Location', 'Customer', 'Address', 'Created By', 'City',
    'Sales Channel', 'Note', 'Exchange Ref', 'Contact No1', 'Contact No2',
    'Sales Person', 'Invoice Type', 'Reference', 'Delivery Amount', 'Product',
    'Cost Price', 'Sales Price', 'Discount', 'Exchange', 'Sold Qty',
    'Cost Amount', 'Sales Amount', 'Paid By', 'Date', 'Serials',
)
# Order derived from inspected SalesRepItem row-4 layouts, NOT guessed from row 3.
FIRST_COL = {'Be Dapper': 'Z', 'Be Desi': 'AB', 'Giggles': 'W',
             'Protecto': 'Y', 'Safety Sure': 'Y'}
MONEY_FIELDS = frozenset({
    'delivery_amount', 'cost_price', 'sales_price', 'discount', 'exchange',
    'cost_amount', 'sales_amount',
})
TEXT_FIELDS = frozenset(SALES_FIELDS) - MONEY_FIELDS - {'sold_qty', 'date'}
MANDATORY = frozenset({'bill_no', 'location', 'product', 'sold_qty', 'sales_amount', 'date'})
MONTH_RE = re.compile(r'^\d{4}-(0[1-9]|1[0-2])$')


def _month(reporting_month):
    if not isinstance(reporting_month, str) or MONTH_RE.fullmatch(reporting_month) is None:
        raise ValueError('reporting_month must be YYYY-MM')
    return reporting_month


def _date(value):
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    if isinstance(value, str):
        value = value.strip()
        # Synthetic export format intentionally ISO-only; reject ambiguous DD/MM/YYYY.
        try:
            return date.fromisoformat(value)
        except ValueError as exc:
            raise ValueError('Date must use ISO YYYY-MM-DD format') from exc
    raise ValueError('Invalid date value')


def _money(value, field):
    if value is None or value == '':
        return None
    if isinstance(value, bool):
        raise ValueError(f'Invalid monetary field {field}')
    try:
        result = Decimal(str(value))
    except (InvalidOperation, TypeError, ValueError) as exc:
        raise ValueError(f'Invalid monetary field {field}') from exc
    if not result.is_finite():
        raise ValueError(f'Non-finite monetary field {field}')
    return result


def _text(value, field):
    if value is None:
        return None
    result = str(value).strip()
    if result.startswith(('=', '+', '-', '@')):
        raise ValueError(f'Formula-like text blocked in {field}')
    return result or None


def normalize_sales_record(record, reporting_month):
    """Normalize a complete 25-field synthetic sale without modifying input."""
    _month(reporting_month)
    if not isinstance(record, dict):
        raise TypeError('Sales record must be a dictionary')
    missing = set(SALES_FIELDS) - record.keys()
    if missing:
        raise ValueError(f'Missing sales fields: {sorted(missing)}')
    result = {}
    for field in SALES_FIELDS:
        value = record[field]
        if field in TEXT_FIELDS:
            result[field] = _text(value, field)
        elif field in MONEY_FIELDS:
            result[field] = _money(value, field)
        elif field == 'sold_qty':
            qty = _money(value, field)
            if qty is None or qty <= 0 or qty != qty.to_integral_value():
                raise ValueError('Sold Qty must be a positive whole number')
            result[field] = int(qty)
        elif field == 'date':
            result[field] = _date(value)
    if any(result[f] is None for f in MANDATORY):
        raise ValueError('Missing mandatory sales value')
    if result['date'].strftime('%Y-%m') != reporting_month:
        raise ValueError('Sales date outside reporting month')
    return result


def sales_column_mapping(business):
    """All 25 source sales fields -> known worksheet column letters."""
    if business not in FIRST_COL or business not in BUSINESS_SCHEMAS:
        raise ValueError(f'Unknown business: {business!r}')
    start = column_index_from_string(FIRST_COL[business])
    # 25 sequential fields in row 4, including Note and Exchange Ref.
    from openpyxl.utils.cell import get_column_letter
    mapping = {field: get_column_letter(start + index)
               for index, field in enumerate(SALES_FIELDS)}
    configured = BUSINESS_SCHEMAS[business]['sales']['columns']
    for field, letter in configured.items():
        if mapping.get(field) != letter:
            raise ValueError(f'Conflict with verified business schema: {business}.{field}')
    return mapping


def sales_cell_values(record, business, reporting_month):
    """Prepare a dictionary of {Excel coordinate letter: validated value}.

    Does not touch a workbook; caller needs independent write safety and auditing.
    """
    values = normalize_sales_record(record, reporting_month)
    return {letter: values[field] for field, letter in sales_column_mapping(business).items()}


def read_synthetic_sales_export(path, reporting_month, *, sheet_name=None):
    """Read a synthetic .xlsx source export without saving it.

    Headers must be exactly the known 26 sales-export fields (including Serials).
    Returns all 25 mapped fields per row, excluding Serials (not yet mapped).
    Rejects extra columns, blank/orphan rows and ambiguous date strings.
    Do NOT pass real exports until ingestion rules for real date formats are approved.
    """
    _month(reporting_month)
    source = Path(path)
    if not source.is_file() or source.suffix.lower() != '.xlsx':
        raise ValueError('Source export must be an existing .xlsx file')
    wb = load_workbook(source, read_only=True, data_only=True)
    try:
        if sheet_name is None:
            if len(wb.worksheets) != 1:
                raise ValueError('Specify sheet_name for multi-sheet source')
            ws = wb.worksheets[0]
        else:
            if sheet_name not in wb:
                raise ValueError('Source worksheet not found')
            ws = wb[sheet_name]
        rows = ws.iter_rows(values_only=True)
        header = next(rows, None)
        if header is None or tuple(header) != SALES_HEADERS:
            raise ValueError('Source sales export headers do not exactly match expected layout')
        results = []
        seen_blank = False
        for row in rows:
            if all(value is None for value in row):
                seen_blank = True
                continue
            if seen_blank:
                raise ValueError('Nonblank record after blank row')
            if len(row) != len(SALES_HEADERS):
                raise ValueError('Invalid source row width')
            record = dict(zip(SALES_FIELDS, row[:len(SALES_FIELDS)]))
            results.append(normalize_sales_record(record, reporting_month))
        return results
    finally:
        wb.close()
