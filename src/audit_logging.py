
# ==========================================
# FINANCIAL AUDIT LOGGING
# LESSON 40
# ==========================================

import hashlib
import json

from datetime import datetime, timezone
from pathlib import Path


# ==========================================
# 1. CALCULATE FILE CHECKSUM
# ==========================================

def calculate_file_hash(file_path):
    """
    Calculate the SHA-256 checksum of a file.

    A checksum helps identify whether
    file contents have changed.
    """

    file_path = Path(file_path)

    sha256 = hashlib.sha256()

    with open(file_path, "rb") as file:

        for chunk in iter(
            lambda: file.read(1024 * 1024),
            b"",
        ):
            sha256.update(chunk)

    return sha256.hexdigest()


# ==========================================
# 2. CREATE AN AUDIT RECORD
# ==========================================

def create_audit_record(
    run_id,
    existing_count,
    inserted_count,
    reconciled,
    output_hash=None,
):
    """
    Create an audit record for a
    synthetic reporting execution.

    The record contains only metadata,
    not transaction-level financial data.
    """

    if not isinstance(run_id, str) or not run_id.strip():
        raise ValueError(
            "A non-empty run ID is required."
        )

    if (
        not isinstance(existing_count, int)
        or isinstance(existing_count, bool)
        or existing_count < 0
    ):
        raise ValueError(
            "Existing count must be a "
            "non-negative integer."
        )

    if (
        not isinstance(inserted_count, int)
        or isinstance(inserted_count, bool)
        or inserted_count < 0
    ):
        raise ValueError(
            "Inserted count must be a "
            "non-negative integer."
        )

    if not isinstance(reconciled, bool):
        raise ValueError(
            "Reconciled must be True or False."
        )

    if output_hash is not None:
        if (
            not isinstance(output_hash, str)
            or len(output_hash) != 64
            or any(
                character not in "0123456789abcdefABCDEF"
                for character in output_hash
            )
        ):
            raise ValueError(
                "Output hash must be a valid "
                "SHA-256 hexadecimal string."
            )

    timestamp = datetime.now(
        timezone.utc
    ).isoformat()

    return {
        "run_id": run_id,
        "timestamp": timestamp,
        "report_type": "SYNTHETIC_TEST_ONLY",
        "existing_count": existing_count,
        "inserted_count": inserted_count,
        "final_count": (
            existing_count + inserted_count
        ),
        "reconciled": reconciled,
        "status": (
            "VERIFIED" if reconciled else "FAILED"
        ),
        "output_sha256": output_hash,
    }


# ==========================================
# 3. SAVE JSON AUDIT LOG
# ==========================================

def save_audit_log(record, output_path):
    """
    Save an audit record as JSON.

    Refuse to overwrite an existing
    audit file.
    """

    if not isinstance(record, dict):
        raise ValueError(
            "Audit record must be a dictionary."
        )

    output_path = Path(output_path)

    output_path.parent.mkdir(
        parents=True,
        exist_ok=True,
    )

    # Exclusive creation prevents overwriting
    # an existing audit file.
    with open(
        output_path,
        "x",
        encoding="utf-8",
    ) as file:

        json.dump(
            record,
            file,
            indent=4,
            ensure_ascii=False,
        )

        file.write("\n")

    return output_path
