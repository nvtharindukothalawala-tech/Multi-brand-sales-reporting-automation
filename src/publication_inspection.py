"""Read-only inspection of a synthetic Excel report / audit publication pair.

This checks the result after publication, including after a possible process
crash. It never deletes, overwrites, or attempts automatic recovery.
"""

from pathlib import Path

from src.audit_verification import verify_workbook_audit


class IncompletePublicationError(ValueError):
    """Only one of the two expected publication files is present."""


def inspect_publication(excel_path, audit_path):
    """Return a verified status for an expected report and audit pair.

    NOT_PUBLISHED: neither file exists.
    VERIFIED: both files exist and the complete workbook audit passes.
    Incomplete or invalid pairs raise an exception; no files are changed.
    """
    excel_path = Path(excel_path)
    audit_path = Path(audit_path)

    if excel_path.resolve() == audit_path.resolve():
        raise ValueError("Excel and audit paths must be different.")
    if excel_path.suffix.lower() != ".xlsx":
        raise ValueError("Excel path must have an .xlsx extension.")
    if audit_path.suffix.lower() != ".json":
        raise ValueError("Audit path must have a .json extension.")

    excel_exists = excel_path.exists()
    audit_exists = audit_path.exists()

    if not excel_exists and not audit_exists:
        return {"publication_status": "NOT_PUBLISHED", "verified": False}
    if excel_exists != audit_exists:
        missing = "audit JSON" if excel_exists else "Excel workbook"
        raise IncompletePublicationError(
            f"Incomplete publication: missing {missing}. "
            "Review the files manually; no cleanup was performed."
        )
    if not excel_path.is_file() or not audit_path.is_file():
        raise ValueError("Publication paths must point to regular files.")

    result = verify_workbook_audit(excel_path, audit_path)
    return {"publication_status": "VERIFIED", **result}
