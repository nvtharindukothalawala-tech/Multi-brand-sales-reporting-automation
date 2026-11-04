"""Synthetic-only Excel writer with transaction checks and optional audit JSON.

This module is for test workbooks marked SYNTHETIC_TEST_ONLY, not production reports.
"""

import os
import tempfile
from collections import Counter
from copy import copy
from decimal import Decimal, InvalidOperation
from pathlib import Path

import pandas as pd
from openpyxl import load_workbook
from openpyxl.formula.translate import Translator

from src.audit_logging import (
    calculate_file_hash,
    create_audit_record,
    save_audit_log,
)
from src.financial_validation import reconcile_totals, calculate_totals
from src.transaction_matching import make_transaction_key


TEST_MARKER = "SYNTHETIC_TEST_ONLY"

COLUMN_MAP = {
    "bill_no": "Z",
    "product": "AO",
    "cost_price": "AP",
    "sales_price": "AQ",
    "sold_qty": "AT",
    "cost_amount": "AU",
    "sales_amount": "AV",
    "date": "AZ",
}

# Zero-based indexes for tuples returned by openpyxl.iter_rows(values_only=True).
COLUMN_INDEXES = {
    "bill_no": 25,
    "product": 40,
    "cost_price": 41,
    "sales_price": 42,
    "sold_qty": 45,
    "cost_amount": 46,
    "sales_amount": 47,
    "date": 51,
}


def read_existing_records(ws):
    """Read the eight transaction fields, starting at worksheet row 5."""
    records = []
    for row in ws.iter_rows(min_row=5, values_only=True):
        if len(row) < 52 or row[25] is None:
            continue
        records.append({
            field: row[index] for field, index in COLUMN_INDEXES.items()
        })
    return records


def validate_workbook_structure(ws):
    """Reject unsupported structures in a controlled synthetic worksheet."""
    if ws.tables:
        raise ValueError("Structured Excel tables are not supported by this writer.")

    for merged_range in ws.merged_cells.ranges:
        if merged_range.max_row >= 5 and merged_range.min_col <= 52:
            raise ValueError(
                f"Merged cells detected in the transaction area: {merged_range}"
            )

    footer_labels = {"TOTAL", "GRAND TOTAL", "SUBTOTAL", "SUMMARY"}
    for row_number in range(5, ws.max_row + 1):
        bill_no = ws[f"Z{row_number}"].value
        if isinstance(bill_no, str) and bill_no.strip().upper() in footer_labels:
            raise ValueError(f"Footer or total row detected at row {row_number}.")

        has_content = any(
            ws.cell(row=row_number, column=column).value is not None
            for column in range(1, 53)
        )
        if has_content and (bill_no is None or not str(bill_no).strip()):
            raise ValueError(
                f"Unsupported non-transaction row detected at row {row_number}."
            )
    return True


def validate_incoming_transactions(records):
    """Validate the complete incoming batch before any worksheet modifications."""
    required_fields = set(COLUMN_MAP)
    numeric_fields = {
        "sold_qty", "cost_price", "sales_price", "cost_amount", "sales_amount"
    }
    if not records:
        raise ValueError("No incoming transactions provided.")

    for row_number, record in enumerate(records, start=1):
        if not isinstance(record, dict):
            raise ValueError(f"Transaction {row_number}: Record must be a dictionary.")

        missing_fields = required_fields - set(record)
        if missing_fields:
            raise ValueError(
                f"Transaction {row_number}: Missing fields: {sorted(missing_fields)}"
            )

        for field in ("bill_no", "product"):
            value = record[field]
            if value is None or not str(value).strip() or pd.isna(value):
                raise ValueError(f"Transaction {row_number}: Invalid {field}.")

        for field in numeric_fields:
            value = record[field]
            if value is None or isinstance(value, bool) or pd.isna(value):
                raise ValueError(f"Transaction {row_number}: Invalid {field}.")
            try:
                number = Decimal(str(value).strip())
            except (InvalidOperation, ValueError, TypeError):
                raise ValueError(f"Transaction {row_number}: Invalid {field}.") from None
            if not number.is_finite():
                raise ValueError(f"Transaction {row_number}: Non-finite {field}.")
            if field == "sold_qty" and number <= 0:
                raise ValueError(
                    f"Transaction {row_number}: Sold quantity must be positive."
                )

        transaction_date = record["date"]
        if transaction_date is None or pd.isna(transaction_date):
            raise ValueError(f"Transaction {row_number}: Missing transaction date.")
        try:
            parsed_date = pd.to_datetime(transaction_date, errors="raise")
            if pd.isna(parsed_date):
                raise ValueError("Invalid date")
        except (ValueError, TypeError, OverflowError):
            raise ValueError(f"Transaction {row_number}: Invalid transaction date.") from None
    return True


def copy_row_formatting_and_formulas(ws, source_row, target_row):
    """Copy styles A:AZ and translate formulas in calculated columns A:Y."""
    for column in range(1, 53):
        source_cell = ws.cell(row=source_row, column=column)
        target_cell = ws.cell(row=target_row, column=column)
        if source_cell.has_style:
            target_cell.font = copy(source_cell.font)
            target_cell.fill = copy(source_cell.fill)
            target_cell.border = copy(source_cell.border)
            target_cell.alignment = copy(source_cell.alignment)
            target_cell.protection = copy(source_cell.protection)
            target_cell.number_format = source_cell.number_format

        if column <= 25:
            formula = source_cell.value
            if isinstance(formula, str) and formula.startswith("="):
                target_cell.value = Translator(
                    formula, origin=source_cell.coordinate
                ).translate_formula(target_cell.coordinate)


def write_synthetic_copy(
    source_path,
    output_path,
    incoming_records,
    audit_path=None,
    run_id=None,
):
    """Create a verified synthetic workbook and optionally a JSON audit file.

    Output paths must be new. Each temporary file is created on the same
    filesystem as its destination for safe, non-overwriting hard links.
    A failed second publication triggers a best-effort Excel rollback.
    """
    source_path = Path(source_path).resolve()
    output_path = Path(output_path).resolve()

    if source_path == output_path:
        raise ValueError("Input and output must be different files.")
    if not source_path.is_file():
        raise FileNotFoundError(source_path)
    if source_path.suffix.lower() != ".xlsx":
        raise ValueError("Only .xlsx input files are supported.")
    if output_path.suffix.lower() != ".xlsx":
        raise ValueError("Output must be an .xlsx file.")
    if output_path.exists():
        raise FileExistsError("Output already exists. Refusing overwrite.")

    if (audit_path is None) != (run_id is None):
        raise ValueError("Provide both audit_path and run_id, or neither.")
    if audit_path is not None:
        audit_path = Path(audit_path).resolve()
        if audit_path in {source_path, output_path}:
            raise ValueError("Audit path must differ from source and Excel output.")
        if audit_path.suffix.lower() != ".json":
            raise ValueError("Audit output must be a .json file.")
        if audit_path.exists():
            raise FileExistsError("Audit output already exists. Refusing overwrite.")
        if not isinstance(run_id, str) or not run_id.strip():
            raise ValueError("A non-empty run ID is required.")

    wb = load_workbook(source_path)
    try:
        if "SalesRepItem" not in wb.sheetnames:
            raise ValueError("SalesRepItem sheet is missing.")
        ws = wb["SalesRepItem"]
        if ws["A1"].value != TEST_MARKER:
            raise ValueError("Not a synthetic test workbook. Writing is blocked.")

        validate_workbook_structure(ws)
        validate_incoming_transactions(incoming_records)
        existing_records = read_existing_records(ws)

        remaining_existing = Counter(
            make_transaction_key(record) for record in existing_records
        )
        new_records = []
        matched_count = 0
        for record in incoming_records:
            key = make_transaction_key(record)
            if remaining_existing[key] > 0:
                remaining_existing[key] -= 1
                matched_count += 1
            else:
                new_records.append(record)

        next_row = max(5, ws.max_row + 1)
        if new_records and not existing_records:
            raise ValueError("No existing transaction template row.")

        for record in new_records:
            copy_row_formatting_and_formulas(ws, next_row - 1, next_row)
            for field, column in COLUMN_MAP.items():
                ws[f"{column}{next_row}"] = record[field]
            next_row += 1

        reconcile_totals(
            existing_records, new_records, read_existing_records(ws)
        )

        output_path.parent.mkdir(parents=True, exist_ok=True)
        if audit_path is not None:
            audit_path.parent.mkdir(parents=True, exist_ok=True)

        with tempfile.TemporaryDirectory(
            prefix=".synthetic_run_", dir=output_path.parent
        ) as excel_temp_dir:
            temp_excel = Path(excel_temp_dir) / "report.xlsx"
            wb.save(temp_excel)

            verified_wb = load_workbook(
                temp_excel, read_only=True, data_only=False
            )
            try:
                verified_records = read_existing_records(
                    verified_wb["SalesRepItem"]
                )
            finally:
                verified_wb.close()
            reconcile_totals(existing_records, new_records, verified_records)

            if audit_path is None:
                # Hard link will fail without overwriting if destination exists.
                os.link(temp_excel, output_path)
            else:
                output_hash = calculate_file_hash(temp_excel)
                audit_record = create_audit_record(
                    run_id=run_id,
                    existing_count=len(existing_records),
                    inserted_count=len(new_records),
                    reconciled=True,
                    output_hash=output_hash,
                )
                # Use verified, saved workbook values as the audit source.
                audit_record["financial_totals"] = {
                    field: str(value)
                    for field, value in calculate_totals(verified_records).items()
                }
                # Keep the temporary audit on the audit destination's filesystem.
                with tempfile.TemporaryDirectory(
                    prefix=".synthetic_audit_", dir=audit_path.parent
                ) as audit_temp_dir:
                    temp_audit = Path(audit_temp_dir) / "audit.json"
                    save_audit_log(audit_record, temp_audit)

                    published_excel = False
                    try:
                        os.link(temp_excel, output_path)
                        published_excel = True
                        # Does not replace a pre-existing audit destination.
                        os.link(temp_audit, audit_path)
                    except Exception:
                        if published_excel:
                            output_path.unlink(missing_ok=True)
                        raise
    finally:
        wb.close()

    result = {
        "matched": matched_count,
        "inserted": len(new_records),
        "existing_before": len(existing_records),
        "existing_after": len(existing_records) + len(new_records),
        "reconciled": True,
    }
    if audit_path is not None:
        result["audit_path"] = str(audit_path)
        result["run_id"] = run_id
    return result
