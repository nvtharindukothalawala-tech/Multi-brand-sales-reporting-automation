
# ==========================================
# AUDIT INTEGRITY VERIFICATION ENGINE
# LESSON 45
# ==========================================

import json

from pathlib import Path

from src.audit_logging import calculate_file_hash

from zipfile import BadZipFile

from openpyxl import load_workbook

# ==========================================
# 1. VERIFY AUDIT INTEGRITY
# ==========================================

def verify_audit_integrity(
    excel_path,
    audit_path,
):
    """
    Verify that a synthetic report matches
    its JSON audit record.

    Checks:
    1. Required files exist
    2. Audit JSON is valid
    3. Report is marked synthetic
    4. Reconciliation succeeded
    5. Audit status is VERIFIED
    6. Recorded checksum matches the file
    """

    excel_path = Path(excel_path)
    audit_path = Path(audit_path)

    # --------------------------------------
    # CHECK 1: REQUIRED FILES
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
    # CHECK 6: TRANSACTION COUNT VALIDATION
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
    # CHECK 6: RECORDED CHECKSUM
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
    # CHECK 7: CALCULATE ACTUAL CHECKSUM
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
        "existing_count": audit_record.get(
            "existing_count"
        ),
        "inserted_count": audit_record.get(
            "inserted_count"
        ),
        "final_count": audit_record.get(
            "final_count"
        ),
    }
# ==========================================
# LESSON 49
# SYNTHETIC WORKBOOK AUDIT VERIFICATION
# ==========================================

def verify_workbook_audit(
    excel_path,
    audit_path,
):
    """
    Verify the integrity and structure
    of a synthetic Excel workbook.

    Checks:
    1. Audit checksum and metadata
    2. XLSX file extension
    3. Valid Excel workbook
    4. Required SalesRepItem worksheet
    5. Synthetic marker in cell A1
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

    finally:

        # Always close the workbook.
        workbook.close()

    # --------------------------------------
    # VERIFICATION SUCCESSFUL
    # --------------------------------------

    result["workbook_verified"] = True
    result["required_sheet"] = "SalesRepItem"
    result["synthetic_marker_verified"] = True

    return result
