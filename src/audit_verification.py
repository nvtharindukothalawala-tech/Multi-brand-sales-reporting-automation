# ==========================================
# AUDIT INTEGRITY VERIFICATION ENGINE
# LESSON 54 - COMPLETE VERSION
# ==========================================

import json
from decimal import Decimal, InvalidOperation
from pathlib import Path
from zipfile import BadZipFile

from openpyxl import load_workbook

from src.audit_logging import calculate_file_hash


SYNTHETIC_MARKER = "SYNTHETIC_TEST_ONLY"
SALES_SHEET = "SalesRepItem"

# Column numbers are 1-based (Z=26, AO=41, AZ=52).
REQUIRED_COLUMNS = {
    "bill_no": 26,
    "product": 41,
    "cost_price": 42,
    "sales_price": 43,
    "sold_qty": 46,
    "cost_amount": 47,
    "sales_amount": 48,
    "date": 52,
}

NUMERIC_FIELDS = {
    "cost_price",
    "sales_price",
    "sold_qty",
    "cost_amount",
    "sales_amount",
}


def verify_audit_integrity(excel_path, audit_path):
    """Check audit metadata, transaction counts, and the report checksum."""
    excel_path = Path(excel_path)
    audit_path = Path(audit_path)

    if not excel_path.is_file():
        raise FileNotFoundError(f"Excel report not found: {excel_path}")
    if not audit_path.is_file():
        raise FileNotFoundError(f"Audit file not found: {audit_path}")

    try:
        with audit_path.open("r", encoding="utf-8") as file:
            audit_record = json.load(file)
    except json.JSONDecodeError:
        raise ValueError("Audit file contains invalid JSON.") from None

    if not isinstance(audit_record, dict):
        raise ValueError("Audit JSON must contain an object.")

    # Unversioned records are legacy and may not contain financial totals.
    # Version 2 is strict: all three totals must be present and valid.
    schema_version = audit_record.get("audit_schema_version", 1)
    if type(schema_version) is not int or schema_version not in (1, 2):
        raise ValueError("Unsupported audit schema version.")
    if schema_version == 2 and "financial_totals" not in audit_record:
        raise ValueError("Version 2 audit is missing financial totals.")

    if audit_record.get("report_type") != SYNTHETIC_MARKER:
        raise ValueError("Audit report type is invalid.")
    if audit_record.get("reconciled") is not True:
        raise ValueError("Audit reconciliation was not successful.")
    if audit_record.get("status") != "VERIFIED":
        raise ValueError("Audit status is not VERIFIED.")

    for field in ("existing_count", "inserted_count", "final_count"):
        value = audit_record.get(field)
        if type(value) is not int or value < 0:
            raise ValueError(f"Invalid audit transaction count: {field}")

    existing_count = audit_record["existing_count"]
    inserted_count = audit_record["inserted_count"]
    final_count = audit_record["final_count"]

    expected_final_count = existing_count + inserted_count
    if final_count != expected_final_count:
        raise ValueError(
            "Audit transaction count mismatch: "
            f"expected {expected_final_count}, found {final_count}."
        )

    recorded_hash = audit_record.get("output_sha256")
    if (
        not isinstance(recorded_hash, str)
        or len(recorded_hash) != 64
        or any(character not in "0123456789abcdef" for character in recorded_hash.lower())
    ):
        raise ValueError("Audit contains an invalid SHA-256 checksum.")

    actual_hash = calculate_file_hash(excel_path)
    if actual_hash.lower() != recorded_hash.lower():
        raise ValueError(
            "Audit integrity verification failed: "
            "Excel file checksum does not match."
        )

    return {
        "verified": True,
        "run_id": audit_record.get("run_id"),
        "status": "VERIFIED",
        "report_type": SYNTHETIC_MARKER,
        "output_sha256": actual_hash,
        "existing_count": existing_count,
        "inserted_count": inserted_count,
        "final_count": final_count,
        "audit_schema_version": schema_version,
    }


def _validate_transaction_row(row, row_number):
    """Validate required fields in one synthetic transaction row."""
    for field, column_number in REQUIRED_COLUMNS.items():
        value = row[column_number - 1]
        if value is None or (isinstance(value, str) and not value.strip()):
            raise ValueError(
                f"Invalid transaction at row {row_number}: missing {field}."
            )

        if field in NUMERIC_FIELDS:
            if isinstance(value, bool):
                raise ValueError(f"Invalid {field} at row {row_number}.")
            try:
                number = Decimal(str(value))
                if not number.is_finite():
                    raise ValueError("Non-finite number")
            except (InvalidOperation, ValueError, TypeError):
                raise ValueError(f"Invalid {field} at row {row_number}.") from None

            if field == "sold_qty" and number <= 0:
                raise ValueError(f"Invalid sold quantity at row {row_number}.")


def verify_workbook_audit(excel_path, audit_path):
    """Verify audit integrity, Excel structure and required transaction values."""
    excel_path = Path(excel_path)
    result = verify_audit_integrity(excel_path, audit_path)

    if excel_path.suffix.lower() != ".xlsx":
        raise ValueError("Expected an .xlsx workbook.")

    try:
        workbook = load_workbook(
            excel_path,
            read_only=True,
            data_only=False,
        )
    except (BadZipFile, OSError, ValueError, KeyError) as error:
        raise ValueError("Invalid or unreadable Excel workbook.") from error

    try:
        if SALES_SHEET not in workbook.sheetnames:
            raise ValueError("Required SalesRepItem worksheet is missing.")

        worksheet = workbook[SALES_SHEET]
        if worksheet["A1"].value != SYNTHETIC_MARKER:
            raise ValueError("Synthetic workbook marker is missing or incorrect.")

        actual_transaction_count = 0
        actual_financial_totals = {
            "sold_qty": Decimal("0"),
            "cost_amount": Decimal("0"),
            "sales_amount": Decimal("0"),
        }
        for row_number, row in enumerate(
            worksheet.iter_rows(
                min_row=5,
                min_col=1,
                max_col=52,
                values_only=True,
            ),
            start=5,
        ):
            bill_no = row[REQUIRED_COLUMNS["bill_no"] - 1]
            if bill_no is None or (isinstance(bill_no, str) and not bill_no.strip()):
                # A blank, formatted row is permitted. Transaction values
                # without a Bill Number are not: otherwise they could be
                # silently excluded from both row counts and financial sums.
                transaction_columns = (
                    column for field, column in REQUIRED_COLUMNS.items()
                    if field != "bill_no"
                )
                if any(
                    row[column - 1] is not None
                    and not (
                        isinstance(row[column - 1], str)
                        and not row[column - 1].strip()
                    )
                    for column in transaction_columns
                ):
                    raise ValueError(
                        f"Orphan transaction data at row {row_number}: "
                        "Bill Number is missing."
                    )
                continue

            _validate_transaction_row(row, row_number)
            for field in actual_financial_totals:
                cell_value = row[REQUIRED_COLUMNS[field] - 1]
                actual_financial_totals[field] += Decimal(str(cell_value))
            actual_transaction_count += 1

        expected_count = result["final_count"]
        if actual_transaction_count != expected_count:
            raise ValueError(
                "Workbook transaction count mismatch: "
                f"Audit expected {expected_count}, "
                f"Excel contains {actual_transaction_count}."
            )

        # Old audit records have no totals; new ones are checked strictly.
        with Path(audit_path).open("r", encoding="utf-8") as audit_file:
            stored_audit = json.load(audit_file)
        stored_totals = stored_audit.get("financial_totals")
        if result["audit_schema_version"] == 2 and stored_totals is None:
            raise ValueError("Version 2 audit is missing financial totals.")
        if "financial_totals" in stored_audit and stored_totals is None:
            raise ValueError("Audit financial totals must not be null.")
        if stored_totals is not None:
            if not isinstance(stored_totals, dict) or set(stored_totals) != set(actual_financial_totals):
                raise ValueError("Invalid audit financial totals structure.")
            for field, actual in actual_financial_totals.items():
                recorded = stored_totals[field]
                if not isinstance(recorded, str):
                    raise ValueError(f"Invalid recorded financial total: {field}.")
                try:
                    expected = Decimal(recorded)
                except InvalidOperation:
                    raise ValueError(f"Invalid recorded financial total: {field}.") from None
                if not expected.is_finite() or expected != actual:
                    raise ValueError(
                        f"Workbook financial total mismatch for {field}: "
                        f"audit={recorded}, workbook={actual}."
                    )
    finally:
        workbook.close()

    result["workbook_verified"] = True
    result["required_sheet"] = SALES_SHEET
    result["synthetic_marker_verified"] = True
    result["actual_transaction_count"] = actual_transaction_count
    result["transaction_count_verified"] = True
    result["transaction_fields_verified"] = True
    result["financial_totals_checked"] = stored_totals is not None
    result["financial_totals_verified"] = stored_totals is not None
    result["legacy_audit"] = result["audit_schema_version"] == 1
    return result
