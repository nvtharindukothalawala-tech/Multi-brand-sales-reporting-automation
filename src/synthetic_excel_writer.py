
# ==========================================
# SYNTHETIC EXCEL WRITING ENGINE
# LESSON 35 - COMPLETE VERSION
# ==========================================

import os
import tempfile

from collections import Counter
from copy import copy
from datetime import date, datetime
from decimal import Decimal, InvalidOperation
from pathlib import Path

import pandas as pd

from openpyxl import load_workbook
from openpyxl.formula.translate import Translator

from src.transaction_matching import make_transaction_key
from src.financial_validation import reconcile_totals


# ==========================================
# 1. SYNTHETIC WORKBOOK CONFIGURATION
# ==========================================

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


# ==========================================
# 2. READ EXISTING TRANSACTIONS
# ==========================================

def read_existing_records(ws):
    """
    Read transaction records from the
    synthetic Excel worksheet.
    """

    records = []

    for row in ws.iter_rows(
        min_row=5,
        values_only=True,
    ):

        if len(row) < 52 or row[25] is None:
            continue

        record = {
            field: row[index]
            for field, index in COLUMN_INDEXES.items()
        }

        records.append(record)

    return records


# ==========================================
# 3. VALIDATE WORKBOOK STRUCTURE
# ==========================================

def validate_workbook_structure(ws):
    """
    Reject unsupported Excel structures.

    This validation does not change
    the original worksheet.
    """

    # --------------------------------------
    # CHECK 1: EXCEL TABLES
    # --------------------------------------

    if len(ws.tables) > 0:
        raise ValueError(
            "Structured Excel tables are not "
            "supported by this writer."
        )

    # --------------------------------------
    # CHECK 2: MERGED CELLS
    # --------------------------------------

    for merged_range in ws.merged_cells.ranges:

        intersects_transaction_area = (
            merged_range.max_row >= 5
            and merged_range.min_col <= 52
        )

        if intersects_transaction_area:
            raise ValueError(
                "Merged cells detected in "
                "the transaction area: "
                f"{merged_range}"
            )

    # --------------------------------------
    # CHECK 3: FOOTER AND TOTAL ROWS
    # --------------------------------------

    footer_labels = {
        "TOTAL",
        "GRAND TOTAL",
        "SUBTOTAL",
        "SUMMARY",
    }

    for row_number in range(5, ws.max_row + 1):

        bill_no = ws[f"Z{row_number}"].value

        if isinstance(bill_no, str):

            if bill_no.strip().upper() in footer_labels:
                raise ValueError(
                    "Footer or total row detected "
                    f"at row {row_number}."
                )

        has_content = any(
            ws.cell(
                row=row_number,
                column=column,
            ).value is not None
            for column in range(1, 53)
        )

        if not has_content:
            continue

        if bill_no is None or str(bill_no).strip() == "":
            raise ValueError(
                "Unsupported non-transaction row "
                f"detected at row {row_number}."
            )

    return True


# ==========================================
# 4. VALIDATE INCOMING TRANSACTIONS
# ==========================================

def validate_incoming_transactions(records):
    """
    Validate every incoming transaction
    before preparing any Excel changes.
    """

    required_fields = {
        "bill_no",
        "product",
        "sold_qty",
        "cost_price",
        "sales_price",
        "cost_amount",
        "sales_amount",
        "date",
    }

    numeric_fields = {
        "sold_qty",
        "cost_price",
        "sales_price",
        "cost_amount",
        "sales_amount",
    }

    if not records:
        raise ValueError(
            "No incoming transactions provided."
        )

    for row_number, record in enumerate(
        records,
        start=1,
    ):

        # ----------------------------------
        # CHECK 1: RECORD TYPE
        # ----------------------------------

        if not isinstance(record, dict):
            raise ValueError(
                f"Transaction {row_number}: "
                "Record must be a dictionary."
            )

        # ----------------------------------
        # CHECK 2: REQUIRED FIELDS
        # ----------------------------------

        missing_fields = (
            required_fields - set(record.keys())
        )

        if missing_fields:
            raise ValueError(
                f"Transaction {row_number}: "
                f"Missing fields: {sorted(missing_fields)}"
            )

        # ----------------------------------
        # CHECK 3: TEXT VALUES
        # ----------------------------------

        for field in ["bill_no", "product"]:

            value = record[field]

            if (
                value is None
                or not str(value).strip()
                or pd.isna(value)
            ):
                raise ValueError(
                    f"Transaction {row_number}: "
                    f"Invalid {field}."
                )

        # ----------------------------------
        # CHECK 4: NUMERIC VALUES
        # ----------------------------------

        for field in numeric_fields:

            value = record[field]

            if (
                value is None
                or isinstance(value, bool)
                or pd.isna(value)
            ):
                raise ValueError(
                    f"Transaction {row_number}: "
                    f"Invalid {field}."
                )

            try:
                number = Decimal(
                    str(value).strip()
                )

            except (InvalidOperation, ValueError):
                raise ValueError(
                    f"Transaction {row_number}: "
                    f"Invalid {field}."
                ) from None

            if not number.is_finite():
                raise ValueError(
                    f"Transaction {row_number}: "
                    f"Non-finite {field}."
                )

            if field == "sold_qty" and number <= 0:
                raise ValueError(
                    f"Transaction {row_number}: "
                    "Sold quantity must be positive."
                )

        # ----------------------------------
        # CHECK 5: TRANSACTION DATE
        # ----------------------------------

        transaction_date = record["date"]

        if (
            transaction_date is None
            or pd.isna(transaction_date)
        ):
            raise ValueError(
                f"Transaction {row_number}: "
                "Missing transaction date."
            )

        try:
            parsed_date = pd.to_datetime(
                transaction_date,
                errors="raise",
            )

            if pd.isna(parsed_date):
                raise ValueError("Invalid date")

        except (
            ValueError,
            TypeError,
            OverflowError,
        ):
            raise ValueError(
                f"Transaction {row_number}: "
                "Invalid transaction date."
            ) from None

    return True


# ==========================================
# 5. COPY FORMULAS AND FORMATTING
# ==========================================

def copy_row_formatting_and_formulas(
    ws,
    source_row,
    target_row,
):
    """
    Copy styles into a new synthetic row.

    Translate formulas in columns A:Y.
    Do not copy fixed transaction values.
    """

    for column in range(1, 53):

        source_cell = ws.cell(
            row=source_row,
            column=column,
        )

        target_cell = ws.cell(
            row=target_row,
            column=column,
        )

        # ----------------------------------
        # COPY CELL FORMATTING
        # ----------------------------------

        if source_cell.has_style:

            target_cell.font = copy(
                source_cell.font
            )

            target_cell.fill = copy(
                source_cell.fill
            )

            target_cell.border = copy(
                source_cell.border
            )

            target_cell.alignment = copy(
                source_cell.alignment
            )

            target_cell.protection = copy(
                source_cell.protection
            )

            target_cell.number_format = (
                source_cell.number_format
            )

        # ----------------------------------
        # TRANSLATE CALCULATED FORMULAS
        # ----------------------------------

        if column <= 25:

            formula = source_cell.value

            if (
                isinstance(formula, str)
                and formula.startswith("=")
            ):

                translated_formula = Translator(
                    formula,
                    origin=source_cell.coordinate,
                ).translate_formula(
                    target_cell.coordinate
                )

                target_cell.value = translated_formula


# ==========================================
# 6. WRITE SYNTHETIC WORKBOOK COPY
# ==========================================

def write_synthetic_copy(
    source_path,
    output_path,
    incoming_records,
):
    """
    Create a verified synthetic Excel copy.

    Features:
    - Protect original workbook
    - Block unsupported structures
    - Validate incoming transactions
    - Prevent repeated transaction insertion
    - Copy supported formulas and formatting
    - Reconcile financial totals
    - Save and verify temporary output
    - Refuse overwriting an existing output
    """

    source_path = Path(source_path).resolve()
    output_path = Path(output_path).resolve()

    # --------------------------------------
    # STEP 1: FILE SAFETY CHECKS
    # --------------------------------------

    if source_path == output_path:
        raise ValueError(
            "Input and output must be different files."
        )

    if not source_path.is_file():
        raise FileNotFoundError(source_path)

    if source_path.suffix.lower() != ".xlsx":
        raise ValueError(
            "Only .xlsx input files are supported."
        )

    if output_path.suffix.lower() != ".xlsx":
        raise ValueError(
            "Output must be an .xlsx file."
        )

    if output_path.exists():
        raise FileExistsError(
            "Output already exists. Refusing overwrite."
        )

    # --------------------------------------
    # STEP 2: LOAD WORKBOOK
    # --------------------------------------

    wb = load_workbook(source_path)

    try:

        if "SalesRepItem" not in wb.sheetnames:
            raise ValueError(
                "SalesRepItem sheet is missing."
            )

        ws = wb["SalesRepItem"]

        # Verify this is a synthetic workbook.
        if ws["A1"].value != TEST_MARKER:
            raise ValueError(
                "Not a synthetic test workbook. "
                "Writing is blocked."
            )

        # ----------------------------------
        # STEP 3: VALIDATE INPUTS
        # ----------------------------------

        validate_workbook_structure(ws)

        validate_incoming_transactions(
            incoming_records
        )

        # ----------------------------------
        # STEP 4: READ EXISTING RECORDS
        # ----------------------------------

        existing_records = read_existing_records(ws)

        # ----------------------------------
        # STEP 5: DETECT DUPLICATES
        # ----------------------------------

        remaining_existing = Counter(
            make_transaction_key(record)
            for record in existing_records
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

        # ----------------------------------
        # STEP 6: PREPARE NEW ROWS
        # ----------------------------------

        next_row = max(
            5,
            ws.max_row + 1,
        )

        if new_records and not existing_records:
            raise ValueError(
                "No existing transaction template row."
            )

        for record in new_records:

            previous_row = next_row - 1

            copy_row_formatting_and_formulas(
                ws=ws,
                source_row=previous_row,
                target_row=next_row,
            )

            for field, column in COLUMN_MAP.items():

                ws[f"{column}{next_row}"] = (
                    record[field]
                )

            next_row += 1

        # ----------------------------------
        # STEP 7: RECONCILE IN MEMORY
        # ----------------------------------

        output_records = read_existing_records(ws)

        reconciliation = reconcile_totals(
            existing_records,
            new_records,
            output_records,
        )

        if not reconciliation["reconciled"]:
            raise ValueError(
                "Financial reconciliation failed."
            )

        # ----------------------------------
        # STEP 8: CREATE TEMPORARY WORKBOOK
        # ----------------------------------

        output_path.parent.mkdir(
            parents=True,
            exist_ok=True,
        )

        fd, temporary_name = tempfile.mkstemp(
            prefix=".synthetic_",
            suffix=".xlsx",
            dir=output_path.parent,
        )

        os.close(fd)

        temporary_path = Path(temporary_name)

        try:

            # Save workbook to temporary file.
            wb.save(temporary_path)

            # ----------------------------------
            # STEP 9: READ SAVED WORKBOOK AGAIN
            # ----------------------------------

            verified_workbook = load_workbook(
                temporary_path,
                read_only=True,
                data_only=False,
            )

            try:

                verified_records = read_existing_records(
                    verified_workbook["SalesRepItem"]
                )

            finally:
                verified_workbook.close()

            # ----------------------------------
            # STEP 10: RECONCILE SAVED OUTPUT
            # ----------------------------------

            reconcile_totals(
                existing_records,
                new_records,
                verified_records,
            )

            # ----------------------------------
            # STEP 11: PUBLISH VERIFIED OUTPUT
            # ----------------------------------

            # Does not overwrite an existing file.
            # Requires filesystem hard-link support.
            os.link(
                temporary_path,
                output_path,
            )

        finally:

            # Clean up temporary Excel file.
            temporary_path.unlink(
                missing_ok=True
            )

    finally:

        # Always close the original workbook.
        wb.close()

    # --------------------------------------
    # STEP 12: RETURN SUMMARY
    # --------------------------------------

    return {
        "matched": matched_count,
        "inserted": len(new_records),
        "existing_before": len(existing_records),
        "existing_after": (
            len(existing_records) + len(new_records)
        ),
        "reconciled": True,
    }
