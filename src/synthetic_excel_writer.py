
from collections import Counter
from copy import copy
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.formula.translate import Translator

from src.transaction_matching import make_transaction_key

from decimal import Decimal, InvalidOperation
from datetime import date, datetime

import pandas as pd

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
    """Read transaction records from a test sheet."""

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
# WORKBOOK STRUCTURE VALIDATION
# ==========================================

def validate_workbook_structure(ws):
    """
    Validate the controlled synthetic layout.

    Reject:
    1. Structured Excel tables
    2. Merged cells in the transaction area
    3. Footer, total, or unsupported non-data rows

    This function does not modify the worksheet.
    """

    # --------------------------------------
    # CHECK 1: STRUCTURED EXCEL TABLES
    # --------------------------------------

    if len(ws.tables) > 0:
        raise ValueError(
            "Structured Excel tables are not "
            "supported by this writer."
        )

    # --------------------------------------
    # CHECK 2: MERGED TRANSACTION CELLS
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
    # CHECK 3: FOOTERS AND UNSUPPORTED ROWS
    # --------------------------------------

    footer_labels = {
        "TOTAL",
        "GRAND TOTAL",
        "SUBTOTAL",
        "SUMMARY",
    }

    for row_number in range(5, ws.max_row + 1):

        bill_no = ws[f"Z{row_number}"].value

        # Check for known footer labels.
        if isinstance(bill_no, str):

            if bill_no.strip().upper() in footer_labels:
                raise ValueError(
                    "Footer or total row detected "
                    f"at row {row_number}."
                )

        # An entirely empty row is acceptable.
        has_content = any(
            ws.cell(
                row=row_number,
                column=column,
            ).value is not None
            for column in range(1, 53)
        )

        if not has_content:
            continue

        # In our synthetic layout, every populated
        # transaction-area row must have a bill number.
        if bill_no is None or str(bill_no).strip() == "":
            raise ValueError(
                "Unsupported non-transaction row "
                f"detected at row {row_number}."
            )

    return True


# ==========================================
# 3. COPY FORMULAS AND FORMATTING
# ==========================================

def copy_row_formatting_and_formulas(
    ws,
    source_row,
    target_row,
):
    """
    Copy styles from an existing synthetic row
    into a new row.

    Translate formulas in calculated columns A:Y.
    Do not copy other cell values.
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

        # Copy cell formatting.
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

        # A:Y are the calculated columns
        # in this controlled synthetic layout.
        #
        # Copy formulas only, not fixed values.
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

                target_cell.value = (
                    translated_formula
                )


# ==========================================
# 4. WRITE A NEW SYNTHETIC WORKBOOK COPY
# ==========================================

def write_synthetic_copy(
    source_path,
    output_path,
    incoming_records,
):
    """
    Create a separate synthetic Excel workbook.

    Preserve the source workbook.
    Prevent duplicate transaction occurrences.
    Copy formatting and supported formulas.
    """

    source_path = Path(source_path).resolve()
    output_path = Path(output_path).resolve()

    # --------------------------------------
    # SAFETY CHECKS
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
    # LOAD THE SYNTHETIC WORKBOOK
    # --------------------------------------

    wb = load_workbook(source_path)

    try:

        if "SalesRepItem" not in wb.sheetnames:
            raise ValueError(
                "SalesRepItem sheet is missing."
            )

        ws = wb["SalesRepItem"]

        # Check synthetic workbook marker first.
        if ws["A1"].value != TEST_MARKER:
            raise ValueError(
                "Not a synthetic test workbook. "
                "Writing is blocked."
            )

        # Validate the synthetic workbook structure.
        validate_workbook_structure(ws)

        # Validate all incoming transactions
        # before preparing any Excel changes.
        validate_incoming_transactions(incoming_records)

        # Read existing workbook transactions.
        existing_records = read_existing_records(ws)



        # ----------------------------------
        # CHECK FOR EXISTING TRANSACTIONS
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
        # ADD ONLY NEW TRANSACTION ROWS
        # ----------------------------------

        next_row = max(
            5,
            ws.max_row + 1,
        )

        # Existing synthetic layout must have
        # a transaction row to use as a template.
        if new_records and not existing_records:
            raise ValueError(
                "No existing transaction template row."
            )

        for record in new_records:

            previous_row = next_row - 1

            # Copy formatting and translate formulas.
            copy_row_formatting_and_formulas(
                ws=ws,
                source_row=previous_row,
                target_row=next_row,
            )

            # Write the new transaction values.
            for field, column in COLUMN_MAP.items():

                ws[f"{column}{next_row}"] = (
                    record[field]
                )

            next_row += 1

        # ----------------------------------
        # SAVE TO A DIFFERENT FILE
        # ----------------------------------

        output_path.parent.mkdir(
            parents=True,
            exist_ok=True,
        )

        wb.save(output_path)

    finally:
        wb.close()

    # --------------------------------------
    # RETURN THE EXECUTION SUMMARY
    # --------------------------------------

    return {
        "matched": matched_count,
        "inserted": len(new_records),
        "existing_before": len(existing_records),
        "existing_after": (
            len(existing_records) + len(new_records)
        ),
    }


# ==========================================
# VALIDATE INCOMING TRANSACTIONS
# ==========================================

def validate_incoming_transactions(records):
    """
    Validate all incoming synthetic transactions.

    Reject invalid data before Excel writing.
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

    if not records:
        raise ValueError(
            "No incoming transactions provided."
        )

    numeric_fields = {
        "sold_qty",
        "cost_price",
        "sales_price",
        "cost_amount",
        "sales_amount",
    }

    for row_number, record in enumerate(
        records, start=1
    ):

        if not isinstance(record, dict):
            raise ValueError(
                f"Transaction {row_number}: "
                "Record must be a dictionary."
            )

        # ----------------------------------
        # CHECK 1: REQUIRED FIELDS
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
        # CHECK 2: BILL NUMBER AND PRODUCT
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
        # CHECK 3: NUMERIC VALUES
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
                number = Decimal(str(value).strip())

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

            # Sales quantity must be positive.
            if field == "sold_qty" and number <= 0:
                raise ValueError(
                    f"Transaction {row_number}: "
                    "Sold quantity must be positive."
                )

        # ----------------------------------
        # CHECK 4: TRANSACTION DATE
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

        except (ValueError, TypeError, OverflowError):
            raise ValueError(
                f"Transaction {row_number}: "
                "Invalid transaction date."
            ) from None

    return True
