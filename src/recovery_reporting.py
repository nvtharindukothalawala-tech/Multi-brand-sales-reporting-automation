"""Create a standalone JSON report from a read-only publication recovery assessment.

This writes ONLY a new, explicitly chosen JSON report path. It never modifies
an Excel workbook or its audit JSON. Existing reports are never overwritten.
"""

import json
import os
import tempfile
from datetime import datetime, timezone
from pathlib import Path

from src.publication_recovery import assess_publication_recovery


REPORT_SCHEMA_VERSION = 1


def create_recovery_report(excel_path, audit_path, report_path, run_id=None):
    """Assess publication safety and exclusively publish a human-reviewable JSON report.

    Returns the report dictionary. Neither input file is modified. The report
    destination's parent directory must already exist. Publishing uses a
    private temporary file and an exclusive hard link, avoiding replacement of
    any existing report. A crash may leave a temporary file for manual cleanup.
    """
    excel_path = Path(excel_path)
    audit_path = Path(audit_path)
    report_path = Path(report_path)

    if excel_path.resolve() == audit_path.resolve():
        raise ValueError("Excel and audit paths must be different.")
    if report_path.suffix.lower() != ".json":
        raise ValueError("Recovery report must be a .json file.")
    if excel_path.suffix.lower() != ".xlsx" or audit_path.suffix.lower() != ".json":
        raise ValueError("Expected .xlsx Excel and .json audit paths.")
    if report_path.resolve() in (excel_path.resolve(), audit_path.resolve()):
        raise ValueError("Recovery report must not overwrite an input file.")
    if not report_path.parent.is_dir():
        raise FileNotFoundError("Recovery report parent directory does not exist.")
    if report_path.exists() or report_path.is_symlink():
        raise FileExistsError(f"Recovery report already exists: {report_path}")
    if run_id is not None and (not isinstance(run_id, str) or not run_id.strip()):
        raise ValueError("run_id must be a nonempty string if provided.")

    assessment = assess_publication_recovery(excel_path, audit_path)
    report = {
        "report_type": "SYNTHETIC_PUBLICATION_RECOVERY_ASSESSMENT",
        "report_schema_version": REPORT_SCHEMA_VERSION,
        "created_at_utc": datetime.now(timezone.utc).isoformat(),
        "run_id": run_id,
        "excel_path": str(excel_path),
        "audit_path": str(audit_path),
        "recovery_status": assessment["recovery_status"],
        "verified": assessment["verified"],
        "safe_to_use": assessment["safe_to_use"],
        "manual_review_required": assessment["manual_review_required"],
        "financial_totals_verified": assessment.get("financial_totals_verified", False),
        "legacy_audit": assessment.get("legacy_audit", False),
        "reason": assessment["reason"],
        "recommended_action": assessment["recommended_action"],
    }

    temporary_path = None
    try:
        # The temporary file and report are on the same filesystem.
        descriptor, temporary_name = tempfile.mkstemp(
            prefix=".recovery_report_", suffix=".tmp", dir=report_path.parent
        )
        temporary_path = Path(temporary_name)
        with os.fdopen(descriptor, "w", encoding="utf-8", newline="\n") as handle:
            json.dump(report, handle, indent=2, ensure_ascii=False, allow_nan=False)
            handle.write("\n")
            handle.flush()
            os.fsync(handle.fileno())
        # Exclusive create: fails rather than replace an existing destination.
        os.link(temporary_path, report_path)
    finally:
        if temporary_path is not None:
            temporary_path.unlink(missing_ok=True)

    return report
