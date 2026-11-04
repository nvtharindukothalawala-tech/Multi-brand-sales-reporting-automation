"""Read-only recovery assessment for SYNTHETIC_TEST_ONLY report publications.

The assessment NEVER modifies source, Excel report, or audit JSON files.
It does not claim that a two-file publication is crash-atomic, and it
never automatically approves publication or deletion.
"""

from pathlib import Path

from src.publication_inspection import (
    IncompletePublicationError,
    inspect_publication,
)


def assess_publication_recovery(excel_path, audit_path):
    """Classify a synthetic publication pair and recommend a safe next action.

    Statuses:
      NOT_PUBLISHED: neither output exists; no recovery is necessary.
      VERIFIED: both files exist and all applicable checks have passed.
      INCOMPLETE: exactly one file exists; preserve it for manual review.
      REVIEW_REQUIRED: both files exist but checks fail (or paths invalid).

    This function only reads metadata and files; it never changes them.
    A VERIFIED legacy audit may lack financial-total verification: inspect
    ``financial_totals_verified`` before relying on financial correctness.
    """
    excel_path = Path(excel_path)
    audit_path = Path(audit_path)

    try:
        result = inspect_publication(excel_path, audit_path)
    except IncompletePublicationError as error:
        return {
            "recovery_status": "INCOMPLETE",
            "verified": False,
            "manual_review_required": True,
            "safe_to_use": False,
            "reason": str(error),
            "recommended_action": (
                "Preserve existing files. Check for interruption, then "
                "review the missing counterpart; do not overwrite or delete automatically."
            ),
        }
    except (ValueError, FileNotFoundError, OSError) as error:
        return {
            "recovery_status": "REVIEW_REQUIRED",
            "verified": False,
            "manual_review_required": True,
            "safe_to_use": False,
            "reason": str(error),
            "recommended_action": (
                "Preserve both paths and investigate the failed verification. "
                "Do not automatically regenerate, overwrite, or delete files."
            ),
        }

    if result["publication_status"] == "NOT_PUBLISHED":
        return {
            "recovery_status": "NOT_PUBLISHED",
            "verified": False,
            "manual_review_required": False,
            "safe_to_use": False,
            "reason": "Neither publication file exists.",
            "recommended_action": (
                "No published output is present. A new synthetic-only run "
                "may be planned through the normal approval process."
            ),
        }

    financial_verified = result.get("financial_totals_verified") is True
    # A legacy pair can pass integrity checks but lack financial audit totals.
    return {
        "recovery_status": "VERIFIED",
        "verified": True,
        "manual_review_required": not financial_verified,
        "safe_to_use": financial_verified,
        "financial_totals_verified": financial_verified,
        "legacy_audit": result.get("legacy_audit", False),
        "reason": (
            "Workbook, checksum, counts, and financial totals verified."
            if financial_verified
            else "Basic verification passed, but financial totals were not verified."
        ),
        "recommended_action": (
            "Retain the verified pair; no recovery action is needed."
            if financial_verified
            else "Review the legacy audit before relying on financial totals."
        ),
    }
