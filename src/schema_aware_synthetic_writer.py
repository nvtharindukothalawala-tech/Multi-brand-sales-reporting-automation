"""Safe, configuration-driven writer for SYNTHETIC test workbooks only.

This is deliberately independent of the existing production/test writer.
It never modifies its source workbook and never overwrites an output file.
"""
from __future__ import annotations

from collections import Counter
from copy import copy
from datetime import date, datetime
from decimal import Decimal, InvalidOperation
from pathlib import Path
import shutil
import tempfile

from openpyxl import load_workbook
from openpyxl.formula.translate import Translator
from openpyxl.utils.cell import column_index_from_string

from src.business_schema import BUSINESS_SCHEMAS

MARKER = 'SYNTHETIC_TEST_ONLY'

FIELDS = {
    'sales': ('bill_no', 'product', 'sold_qty', 'sales_amount', 'date'),
    'returns': ('return_code', 'bill_no', 'product', 'return_qty', 'sale_amount', 'return_date'),
}
HEADERS = {
    'sales': {'bill_no': 'Bill No', 'product': 'Product', 'sold_qty': 'Sold Qty',
              'sales_amount': 'Sales Amount', 'date': 'Date'},
    'returns': {'return_code': 'Return Code', 'bill_no': 'Bill No',
                'product': 'Product Description', 'return_qty': 'Return Qty',
                'sale_amount': 'Sale Amount', 'return_date': 'Return Date'},
}


def _date(value):
    if isinstance(value, datetime):
        return value.date().isoformat()
    if isinstance(value, date):
        return value.isoformat()
    if isinstance(value, str):
        try:
            return datetime.fromisoformat(value.strip().replace('Z', '+00:00')).date().isoformat()
        except ValueError as exc:
            raise ValueError(f'Invalid/ambiguous ISO date: {value!r}') from exc
    raise ValueError(f'Invalid date value: {type(value).__name__}')


def _decimal(value, label):
    if isinstance(value, bool):
        raise ValueError(f'Invalid {label}: boolean')
    try:
        result = Decimal(str(value))
    except (InvalidOperation, TypeError, ValueError) as exc:
        raise ValueError(f'Invalid {label}: {value!r}') from exc
    if not result.is_finite() or result < 0:
        raise ValueError(f'{label} must be a non-negative finite number')
    return result


def _identity(record, kind):
    if not isinstance(record, dict):
        raise ValueError('Each transaction must be a dictionary')
    fields = FIELDS[kind]
    if any(field not in record for field in fields):
        raise ValueError(f'Missing required {kind} fields: {sorted(set(fields) - record.keys())}')
    labels = ('bill_no', 'product') if kind == 'sales' else ('return_code', 'bill_no', 'product')
    for field in labels:
        if record[field] is None or not str(record[field]).strip():
            raise ValueError(f'Blank required field: {field}')
    qtyfield = 'sold_qty' if kind == 'sales' else 'return_qty'
    amtfield = 'sales_amount' if kind == 'sales' else 'sale_amount'
    datefield = 'date' if kind == 'sales' else 'return_date'
    qty = _decimal(record[qtyfield], qtyfield)
    if qty <= 0 or qty != qty.to_integral_value():
        raise ValueError(f'{qtyfield} must be a positive whole number')
    amount = _decimal(record[amtfield], amtfield)
    prefix = tuple(str(record[f]).strip() for f in labels)
    if any(value.startswith(('=', '+', '-', '@')) for value in prefix):
        raise ValueError('Formula-like text is blocked in synthetic transaction keys')
    return prefix + (qty, amount, _date(record[datefield]))


def _record_from_row(ws, row, columns, kind):
    return {field: ws[f'{columns[field]}{row}'].value for field in FIELDS[kind]}


def _validate_layout(ws, schema, kind):
    header_row = schema['header_row']
    first_row = schema['first_data_row']
    if not isinstance(header_row, int) or not isinstance(first_row, int) or first_row != header_row + 1:
        raise ValueError('Schema header/data row configuration is invalid')
    positions = []
    for field, expected in HEADERS[kind].items():
        letter = schema['columns'][field]
        positions.append(column_index_from_string(letter))
        actual = ws[f'{letter}{header_row}'].value
        if not isinstance(actual, str) or actual.strip().casefold() != expected.casefold():
            raise ValueError(f'Header mismatch: {kind}.{field} at {letter}{header_row}; found {actual!r}')
    if len(positions) != len(set(positions)):
        raise ValueError('Duplicate required field column mapping')
    if ws.tables:
        raise ValueError('Excel tables are not supported in this synthetic writer')
    if any(r.max_row >= first_row for r in ws.merged_cells.ranges):
        raise ValueError('Merged transaction-region cells are not supported')


def _existing_rows(ws, start_row):
    """Require contiguous, fully represented data; reject orphan/footer content."""
    last = start_row - 1
    seen_gap = False
    for row in range(start_row, ws.max_row + 1):
        occupied = any(cell.value is not None for cell in ws[row])
        if not occupied:
            seen_gap = True
        elif seen_gap:
            raise ValueError(f'Content after blank row in {ws.title}: row {row}')
        else:
            last = row
    return last


def _append(ws, rownum, record, kind, schema, prior_data_row):
    columns = schema['columns']
    # Preserve existing formulas and formatting on an existing transaction row.
    # Never copy the header row into the first transaction row.
    if prior_data_row is not None:
        for source_cell in ws[prior_data_row]:
            target = ws.cell(rownum, source_cell.column)
            if target.value is not None:
                raise ValueError(f'Unexpected occupied destination: {target.coordinate}')
            if source_cell.has_style:
                target._style = copy(source_cell._style)
            if source_cell.number_format:
                target.number_format = source_cell.number_format
            if source_cell.data_type == 'f' and source_cell.column not in {
                column_index_from_string(columns[f]) for f in FIELDS[kind]
            }:
                target.value = Translator(source_cell.value, origin=source_cell.coordinate).translate_formula(target.coordinate)
        if ws.row_dimensions[prior_data_row].height is not None:
            ws.row_dimensions[rownum].height = ws.row_dimensions[prior_data_row].height
    for field in FIELDS[kind]:
        value = record[field]
        if field in ('date', 'return_date'):
            value = date.fromisoformat(_date(value))
        elif field in ('sold_qty', 'return_qty'):
            value = int(_decimal(value, field))
        elif field in ('sales_amount', 'sale_amount'):
            value = float(_decimal(value, field))
        elif value is not None:
            value = str(value).strip()
            # Formula-like text has already been rejected by _identity.
        cell = ws[f'{columns[field]}{rownum}']
        if cell.value is not None:
            raise ValueError(f'Existing destination value: {cell.coordinate}')
        cell.value = value
        if field in ('date', 'return_date'):
            cell.number_format = 'yyyy-mm-dd'


def write_synthetic_business_workbook(source_path, output_path, business, kind, transactions):
    """Write only new occurrences to a separate, marked SYNTHETIC workbook.

    Returns {'existing': int, 'requested': int, 'added': int}.
    No output is created if validation fails; no existing output is overwritten.
    """
    if business not in BUSINESS_SCHEMAS:
        raise ValueError(f'Unknown business: {business!r}')
    if kind not in FIELDS:
        raise ValueError("kind must be 'sales' or 'returns'")
    source = Path(source_path)
    output = Path(output_path)
    if not source.is_file() or source.suffix.lower() != '.xlsx':
        raise ValueError('Source must be an existing .xlsx file')
    if source.resolve() == output.resolve():
        raise ValueError('Input and output paths must differ')
    if output.suffix.lower() != '.xlsx' or output.exists():
        raise ValueError('Output must be a new .xlsx path; overwriting is forbidden')
    if not isinstance(transactions, (list, tuple)):
        raise ValueError('transactions must be a list or tuple')
    # Validate ALL requested records before opening/saving the workbook.
    request_keys = [_identity(rec, kind) for rec in transactions]

    wb = load_workbook(source, read_only=False, data_only=False, keep_links=True)
    try:
        if 'SalesRepItem' not in wb or wb['SalesRepItem']['A1'].value != MARKER:
            raise PermissionError('BLOCKED: workbook is not marked SYNTHETIC_TEST_ONLY')
        schema = BUSINESS_SCHEMAS[business][kind]
        if schema['sheet'] not in wb:
            raise ValueError(f'Missing expected worksheet: {schema["sheet"]}')
        ws = wb[schema['sheet']]
        _validate_layout(ws, schema, kind)
        last = _existing_rows(ws, schema['first_data_row'])
        existing = Counter()
        for r in range(schema['first_data_row'], last + 1):
            existing[_identity(_record_from_row(ws, r, schema['columns'], kind), kind)] += 1
        target = Counter(request_keys)
        additional = target - existing
        needed = additional.copy()
        next_row = last + 1
        for rec, key in zip(transactions, request_keys):
            if needed[key] > 0:
                _append(ws, next_row, rec, kind, schema,
                        prior_data_row=next_row - 1 if next_row > schema['first_data_row'] else None)
                next_row += 1
                needed[key] -= 1
        added = next_row - (last + 1)
        # Verify in-memory transaction multiplicities before publishing.
        result = Counter(
            _identity(_record_from_row(ws, r, schema['columns'], kind), kind)
            for r in range(schema['first_data_row'], next_row)
        )
        if result != existing + additional:
            raise ValueError('Post-write reconciliation failed; output blocked')
        output.parent.mkdir(parents=True, exist_ok=True)
        tempname = None
        created_output = False
        try:
            with tempfile.NamedTemporaryFile(prefix='.synthetic_', suffix='.xlsx',
                                             dir=output.parent, delete=False) as f:
                tempname = Path(f.name)
            wb.save(tempname)
            # Exclusive creation: never overwrite output even if another process created it.
            with output.open('xb') as dest:
                created_output = True
                with tempname.open('rb') as src:
                    shutil.copyfileobj(src, dest)
        except BaseException:
            # Delete only a partial output created by this call, never a pre-existing file.
            if created_output:
                output.unlink(missing_ok=True)
            raise
        finally:
            if tempname is not None:
                tempname.unlink(missing_ok=True)
        return {'existing': sum(existing.values()), 'requested': len(transactions), 'added': added}
    finally:
        wb.close()
