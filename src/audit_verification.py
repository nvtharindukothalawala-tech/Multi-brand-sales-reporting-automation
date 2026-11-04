
# ==========================================
# AUDIT INTEGRITY VERIFICATION ENGINE
# LESSON 52 - COMPLETE VERSION
# ==========================================

import json

from pathlib import Path
from zipfile import BadZipFile

from openpyxl import load_workbook

from src.audit_logging import calculate_file_hash


# ==========================================
# 1. VERIFY AUDIT INTEGRITY
# ==========================================

def verify_audit_integrity(
    excel_path,
    audit_path,
):
    """
    Verify a synthetic report against its
    JSON audit record.

    Checks:
    1. Required files exist
    2. Audit JSON is valid
    3. Report type is synthetic
    4. Financial reconciliation succeeded
    5. Audit status is VERIFIED
    6. Transaction counts are valid
    7. Transaction counts are consistent
    8. Recorded checksum is valid
    9. Actual checksum matches the audit
    """

    excel_path = Path(excel_path)
    audit_path = Path(audit_path)

    # --------------------------------------
    # CHECK 1: REQUIRED FILES EXIST
    # --------------------------------------

    if not excel_path.is_file():
        raise FileNotFoundError(
            f"Excel report not found: {excel_path}"
        )

    if not audit_path.is_file():
        raise FileNotFoundError(
            f"Audit file not found: {audit_path}"
        )

    # --------------------------------------
    # CHECK 2: READ AUDIT JSON
    # --------------------------------------

    try:
        with open(
            audit_path,
            "r",
            encoding="utf-8",
        ) as file:
            audit_record = json.load(file)

    except json.JSONDecodeError:
        raise ValueError(
            "Audit file contains invalid JSON."
        ) from None

    if not isinstance(audit_record, dict):
        raise ValueError(
            "Audit JSON must contain an object."
        )

    # --------------------------------------
    # CHECK 3: SYNTHETIC REPORT TYPE
    # --------------------------------------

    if audit_record.get("report_type") != (
        "SYNTHETIC_TEST_ONLY"
    ):
        raise ValueError(
            "Audit report type is invalid."
        )

    # --------------------------------------
    # CHECK 4: RECONCILIATION STATUS
    # --------------------------------------

    if audit_record.get("reconciled") is not True:
        raise ValueError(
            "Audit reconciliation was not successful."
        )

    # --------------------------------------
    # CHECK 5: AUDIT STATUS
    # --------------------------------------

    if audit_record.get("status") != "VERIFIED":
        raise ValueError(
            "Audit status is not VERIFIED."
        )

    # --------------------------------------
    # CHECK 6: VALIDATE TRANSACTION COUNTS
    # --------------------------------------

    count_fields = (
        "existing_count",
        "inserted_count",
        "final_count",
    )

    for field in count_fields:

        value = audit_record.get(field)

        # Reject missing, negative, decimal,
        # string, and boolean count values.
        if (
            type(value) is not int
            or value < 0
        ):
            raise ValueError(
                f"Invalid audit transaction count: "
                f"{field}"
            )

    # --------------------------------------
    # CHECK 7: COUNT CONSISTENCY
    # --------------------------------------

    existing_count = audit_record["existing_count"]
    inserted_count = audit_record["inserted_count"]
    final_count = audit_record["final_count"]

    expected_final_count = (
        existing_count + inserted_count
    )

    if final_count != expected_final_count:
        raise ValueError(
            "Audit transaction count mismatch: "
            f"expected {expected_final_count}, "
            f"found {final_count}."
        )

    # --------------------------------------
    # CHECK 8: RECORDED CHECKSUM
    # --------------------------------------

    recorded_hash = audit_record.get(
        "output_sha256"
    )

    if (
        not isinstance(recorded_hash, str)
        or len(recorded_hash) != 64
        or any(
            character not in "0123456789abcdef"
            for character in recorded_hash.lower()
        )
    ):
        raise ValueError(
            "Audit contains an invalid SHA-256 checksum."
        )

    # --------------------------------------
    # CHECK 9: CALCULATE ACTUAL CHECKSUM
    # --------------------------------------

    actual_hash = calculate_file_hash(
        excel_path
    )

    if actual_hash.lower() != recorded_hash.lower():
        raise ValueError(
            "Audit integrity verification failed: "
            "Excel file checksum does not match."
        )

    # --------------------------------------
    # VERIFICATION SUCCESSFUL
    # --------------------------------------

    return {
        "verified": True,
        "run_id": audit_record.get("run_id"),
        "status": "VERIFIED",
        "report_type": "SYNTHETIC_TEST_ONLY",
        "output_sha256": actual_hash,
        "existing_count": existing_count,
        "inserted_count": inserted_count,
        "final_count": final_count,
    }


# ==========================================
# 2. VERIFY SYNTHETIC EXCEL WORKBOOK
# ==========================================

def verify_workbook_audit(
    excel_path,
    audit_path,
):
    """
    Verify the integrity and structure
    of a synthetic Excel workbook.

    Checks:
    1. Audit JSON and SHA-256 checksum
    2. XLSX file extension
    3. Excel workbook is readable
    4. SalesRepItem worksheet exists
    5. Synthetic marker exists in A1
    6. Count actual Excel transactions
    7. Compare actual count with audit
    """

    excel_path = Path(excel_path)

    # --------------------------------------
    # CHECK 1: AUDIT INTEGRITY
    # --------------------------------------

    result = verify_audit_integrity(
        excel_path,
        audit_path,
    )

    # --------------------------------------
    # CHECK 2: XLSX EXTENSION
    # --------------------------------------

    if excel_path.suffix.lower() != ".xlsx":
        raise ValueError(
            "Expected an .xlsx workbook."
        )

    # --------------------------------------
    # CHECK 3: OPEN THE WORKBOOK
    # --------------------------------------

    try:
        workbook = load_workbook(
            excel_path,
            read_only=True,
            data_only=False,
        )

    except (
        BadZipFile,
        OSError,
        ValueError,
        KeyError,
    ) as error:
        raise ValueError(
            "Invalid or unreadable Excel workbook."
        ) from error

    try:

        # ----------------------------------
        # CHECK 4: REQUIRED WORKSHEET
        # ----------------------------------

        if "SalesRepItem" not in workbook.sheetnames:
            raise ValueError(
                "Required SalesRepItem "
                "worksheet is missing."
            )

        worksheet = workbook["SalesRepItem"]

        # ----------------------------------
        # CHECK 5: SYNTHETIC MARKER
        # ----------------------------------

        if worksheet["A1"].value != (
            "SYNTHETIC_TEST_ONLY"
        ):
            raise ValueError(
                "Synthetic workbook marker "
                "is missing or incorrect."
            )

        # ----------------------------------
        # CHECK 6: COUNT ACTUAL TRANSACTIONS
        # ----------------------------------

        # Transaction records begin at row 5.
        # Bill numbers are stored in column Z.
        #
        # Column Z = Excel column number 26.

        actual_transaction_count = 0

        for row in worksheet.iter_rows(
            min_row=5,
            min_col=26,
            max_col=26,
            values_only=True,
        ):

            bill_no = row[0]

            # Only count rows with
            # a non-empty Bill Number.
            if (
                bill_no is not None
                and str(bill_no).strip() != ""
            ):
                actual_transaction_count += 1

        # ----------------------------------
        # CHECK 7: COMPARE WITH AUDIT COUNT
        # ----------------------------------

        expected_transaction_count = (
            result["final_count"]
        )

        if (
            actual_transaction_count
            != expected_transaction_count
        ):
            raise ValueError(
                "Workbook transaction count mismatch: "
                f"Audit expected "
                f"{expected_transaction_count}, "
                f"Excel contains "
                f"{actual_transaction_count}."
            )

    finally:

        # Always close the workbook.
        workbook.close()

    # --------------------------------------
    # VERIFICATION SUCCESSFUL
    # --------------------------------------

    result["workbook_verified"] = True
    result["required_sheet"] = "SalesRepItem"
    result["synthetic_marker_verified"] = True

    # Additional verification metadata
    result["actual_transaction_count"] = (
        actual_transaction_count
    )

    result["transaction_count_verified"] = True

    return result
