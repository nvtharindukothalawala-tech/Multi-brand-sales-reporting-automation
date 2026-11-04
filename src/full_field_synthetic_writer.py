"""Full 25-field synthetic sales writer, additive to the existing project.

Never invoke this on a real company workbook or a copy of company data.
Workbook must be independently created and marked SYNTHETIC_TEST_ONLY.
No existing source is saved and no existing output is overwritten.
"""
from __future__ import annotations

from collections import Counter
from datetime import date
from decimal import Decimal
from pathlib import Path
from tempfile import TemporaryDirectory
import shutil

from openpyxl import load_workbook
from src.business_schema import BUSINESS_SCHEMAS
from src.full_field_source_mapping import (
    MONEY_FIELDS, SALES_FIELDS, normalize_sales_record, sales_column_mapping,
)
from src.schema_aware_synthetic_writer import (
    MARKER, _identity, _existing_rows, _record_from_row, _validate_layout,
    write_synthetic_business_workbook,
)

CORE_FIELDS = ('bill_no', 'product', 'sold_qty', 'sales_amount', 'date')
EXTRA_FIELDS = tuple(field for field in SALES_FIELDS if field not in CORE_FIELDS)


def _record_signature(record):
    """Stable full-record identity (text, dates, numbers), including nullable fields."""
    result = []
    for field in SALES_FIELDS:
        value = record[field]
        if isinstance(value, Decimal):
            result.append(value.normalize())
        elif isinstance(value, date):
            result.append(value.isoformat())
        else:
            result.append(value)
    return tuple(result)


def _stored_row(ws, row, mapping, month):
    record = {field: ws[f'{mapping[field]}{row}'].value for field in SALES_FIELDS}
    return normalize_sales_record(record, month)


def _safe_numeric(value, field):
    """Require monetary values representable as exact cents for roundtrip checks."""
    if value is None:
        return None
    if not isinstance(value, Decimal):
        raise ValueError(f'Unexpected {field} value')
    if value != value.quantize(Decimal('0.01')):
        raise ValueError(f'{field} exceeds cent precision; review before writing')
    if abs(value) > Decimal('999999999999.99'):
        raise ValueError(f'{field} exceeds safe synthetic workbook range')
    return float(value)


def _exclusive_copy(source, destination):
    """Destination must not exist; delete only partial file created by this call."""
    created = False
    try:
        with Path(destination).open('xb') as target:
            created = True
            with Path(source).open('rb') as src:
                shutil.copyfileobj(src, target)
    except BaseException:
        if created:
            Path(destination).unlink(missing_ok=True)
        raise


def write_full_field_synthetic_sales(source_path, output_path, business, transactions, month):
    """Write all 25 fields for synthetic sales, detecting full-field conflicts.

    Uses existing core synthetic writer to preserve formula/style/structure safeguards.
    Returns {'existing': n, 'requested': n, 'added': n, 'full_fields': 25}.
    Existing rows with the same core key but differing full data are refused,
    rather than silently treated as already present.
    """
    if business not in BUSINESS_SCHEMAS:
        raise ValueError('Unknown business')
    src, dest = Path(source_path), Path(output_path)
    if not src.is_file() or src.suffix.lower() != '.xlsx':
        raise ValueError('Existing .xlsx synthetic source required')
    if src.resolve() == dest.resolve() or dest.exists() or dest.suffix.lower() != '.xlsx':
        raise ValueError('Input/output overlap or output already exists')
    if not isinstance(transactions, (list, tuple)):
        raise ValueError('Transactions must be a list or tuple')

    normalized = [normalize_sales_record(row, month) for row in transactions]
    for rec in normalized:
        for field in MONEY_FIELDS:
            _safe_numeric(rec[field], field)
    mapping = sales_column_mapping(business)
    schema = BUSINESS_SCHEMAS[business]['sales']
    wb = load_workbook(src, read_only=False, data_only=False, keep_links=True)
    try:
        if 'SalesRepItem' not in wb or wb['SalesRepItem']['A1'].value != MARKER:
            raise PermissionError('BLOCKED: only independently generated synthetic workbooks')
        if schema['sheet'] not in wb:
            raise ValueError('Missing SalesRepItem worksheet')
        ws = wb[schema['sheet']]
        _validate_layout(ws, schema, 'sales')
        # Check all 25 headers, not only the five core fields.
        from src.full_field_source_mapping import SALES_HEADERS
        for field, label in zip(SALES_FIELDS, SALES_HEADERS):
            coord = f'{mapping[field]}{schema["header_row"]}'
            if str(ws[coord].value or '').strip().casefold() != label.casefold():
                raise ValueError(f'Full-field header mismatch: {coord} (expected {label})')
        last = _existing_rows(ws, schema['first_data_row'])
        original_rows = [_stored_row(ws, row, mapping, month)
                         for row in range(schema['first_data_row'], last + 1)]
        original_core = Counter(_identity(row, 'sales') for row in original_rows)
        requested_core = Counter(_identity(row, 'sales') for row in normalized)
        original_full = Counter(_record_signature(row) for row in original_rows)
        requested_full = Counter(_record_signature(row) for row in normalized)
        new_full = requested_full - original_full
        new_core = requested_core - original_core
        new_rows = []
        still_needed = new_full.copy()
        for row in normalized:
            signature = _record_signature(row)
            if still_needed[signature]:
                new_rows.append(row)
                still_needed[signature] -= 1
        # Reject any conflicting full row sharing an existing 5-field identity.
        if Counter(_identity(row, 'sales') for row in new_rows) != new_core:
            raise ValueError('Full-field mismatch for existing transaction key; output blocked')
        # Validate existing formula fields would not be overwritten by appended extras.
        if new_rows and last >= schema['first_data_row']:
            for field in EXTRA_FIELDS:
                if ws[f'{mapping[field]}{last}'].data_type == 'f':
                    raise ValueError(f'Formula occupies mapped field: {field}')
    finally:
        wb.close()

    # The original core writer appends first N records per key. Put required new
    # occurrences first, then existing matches, so appended rows are deterministic.
    # Supplying all requested records twice would alter core multiplicities;
    # instead keep only new rows, plus the original requested existing-occurrence
    # count per key. Core writer targets max(existing, requested), not sum.
    remainder = requested_core - Counter(_identity(r, 'sales') for r in new_rows)
    requested_ordered = list(new_rows)
    for row in normalized:
        key = _identity(row, 'sales')
        if remainder[key]:
            requested_ordered.append(row)
            remainder[key] -= 1

    with TemporaryDirectory(prefix='.full_field_synthetic_') as temporary:
        staged_core = Path(temporary) / 'core.xlsx'
        result = write_synthetic_business_workbook(
            src, staged_core, business, 'sales', requested_ordered
        )
        if result['added'] != len(new_rows):
            raise ValueError('Core writer added unexpected number of sales lines')
        staged_full = Path(temporary) / 'full.xlsx'
        staged_wb = load_workbook(staged_core, read_only=False, data_only=False, keep_links=True)
        try:
            ws = staged_wb[schema['sheet']]
            for offset, rec in enumerate(new_rows):
                rownum = last + 1 + offset
                for field in EXTRA_FIELDS:
                    coord = f'{mapping[field]}{rownum}'
                    cell = ws[coord]
                    if cell.value is not None:
                        raise ValueError(f'Existing content in mapped field {coord}')
                    value = rec[field]
                    if field in MONEY_FIELDS:
                        value = _safe_numeric(value, field)
                    cell.value = value
            # Verify all 25 fields in every saved row in memory.
            for offset, rec in enumerate(new_rows):
                stored = _stored_row(ws, last + 1 + offset, mapping, month)
                if _record_signature(stored) != _record_signature(rec):
                    raise ValueError('25-field in-memory reconciliation failure')
            staged_wb.save(staged_full)
        finally:
            staged_wb.close()
        # Verify on-disk saved values before publication.
        verify_wb = load_workbook(staged_full, read_only=True, data_only=False)
        try:
            verify_ws = verify_wb[schema['sheet']]
            for offset, rec in enumerate(new_rows):
                stored = _stored_row(verify_ws, last + 1 + offset, mapping, month)
                if _record_signature(stored) != _record_signature(rec):
                    raise ValueError('Saved 25-field reconciliation failure')
        finally:
            verify_wb.close()
        dest.parent.mkdir(parents=True, exist_ok=True)
        _exclusive_copy(staged_full, dest)
    return {'existing': len(original_rows), 'requested': len(normalized),
            'added': len(new_rows), 'full_fields': 25}
