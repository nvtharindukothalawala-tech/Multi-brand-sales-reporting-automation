"""CLI for read-only synthetic publication recovery reports.

Usage:
    python -m src.recovery_cli --excel report.xlsx --audit audit.json --output recovery_report.json
    python -m src.recovery_cli --excel report.xlsx --audit audit.json --output recovery_report.json --format text

Exit codes: 0 = verified and safe; 2 = review/not published; 1 = report error.
Both display formats write the same JSON report; neither modifies the inputs.
"""

import argparse
import json
import sys

from src.recovery_reporting import create_recovery_report


def _display_result(report, output, output_format):
    """Display a compact structured result, or a human-readable summary."""
    summary = {
        "recovery_status": report["recovery_status"],
        "safe_to_use": report["safe_to_use"],
        "manual_review_required": report["manual_review_required"],
        "recommended_action": report["recommended_action"],
        "report_path": output,
    }
    if output_format == "json":
        print(json.dumps(summary, indent=2, ensure_ascii=False))
    else:
        print(f"Recovery status : {summary['recovery_status']}")
        print(f"Safe to use     : {'YES' if summary['safe_to_use'] else 'NO'}")
        print(f"Manual review   : {'YES' if summary['manual_review_required'] else 'NO'}")
        print(f"Recommended step: {summary['recommended_action']}")
        print(f"Saved report    : {summary['report_path']}")


def main(argv=None):
    parser = argparse.ArgumentParser(
        description="Generate a JSON recovery assessment for a synthetic Excel/audit pair."
    )
    parser.add_argument("--excel", required=True, help="Path to synthetic .xlsx output")
    parser.add_argument("--audit", required=True, help="Path to matching audit .json")
    parser.add_argument("--output", required=True, help="New recovery report .json path")
    parser.add_argument("--run-id", default=None, help="Optional nonempty run identifier")
    parser.add_argument(
        "--format", choices=("json", "text"), default="json",
        help="Console display format (default: json). The saved report is always JSON.",
    )
    parser.add_argument(
        "--error-format", choices=("text", "json"), default="text",
        help="Error display format on stderr (default: text).",
    )
    args = parser.parse_args(argv)

    try:
        report = create_recovery_report(
            excel_path=args.excel,
            audit_path=args.audit,
            report_path=args.output,
            run_id=args.run_id,
        )
    except (ValueError, OSError) as error:
        message = f"Recovery report not created: {error}"
        if args.error_format == "json":
            print(json.dumps({"error": message, "report_created": False}), file=sys.stderr)
        else:
            print(message, file=sys.stderr)
        return 1

    _display_result(report, args.output, args.format)
    return 0 if report["safe_to_use"] is True else 2


if __name__ == "__main__":
    sys.exit(main())
