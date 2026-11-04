"""CLI for read-only synthetic publication recovery reports.

Run from project root:
    python -m src.recovery_cli --excel report.xlsx --audit audit.json --output recovery_report.json

Exit codes: 0 = verified and safe; 2 = review/not published; 1 = CLI/report error.
"""

import argparse
import json
import sys

from src.recovery_reporting import create_recovery_report


def main(argv=None):
    parser = argparse.ArgumentParser(
        description="Generate a JSON recovery assessment for a synthetic Excel/audit pair."
    )
    parser.add_argument("--excel", required=True, help="Path to synthetic .xlsx output")
    parser.add_argument("--audit", required=True, help="Path to matching audit .json")
    parser.add_argument("--output", required=True, help="New recovery report .json path")
    parser.add_argument("--run-id", default=None, help="Optional nonempty run identifier")
    args = parser.parse_args(argv)

    try:
        report = create_recovery_report(
            excel_path=args.excel,
            audit_path=args.audit,
            report_path=args.output,
            run_id=args.run_id,
        )
    except (ValueError, OSError) as error:
        print(f"Recovery report not created: {error}", file=sys.stderr)
        return 1

    print(json.dumps({
        "recovery_status": report["recovery_status"],
        "safe_to_use": report["safe_to_use"],
        "manual_review_required": report["manual_review_required"],
        "recommended_action": report["recommended_action"],
        "report_path": args.output,
    }, indent=2))
    return 0 if report["safe_to_use"] is True else 2


if __name__ == "__main__":
    sys.exit(main())
