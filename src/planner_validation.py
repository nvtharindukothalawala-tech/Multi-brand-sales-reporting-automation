
import pandas as pd


def validate_reporting_month(dates, expected_month):
    """
    Check that every transaction belongs
    to the expected reporting month.
    """

    if not pd.Series([expected_month]).str.fullmatch(
        r"\d{4}-(0[1-9]|1[0-2])"
    ).all():
        raise ValueError("Invalid reporting month format")

    if len(dates) == 0:
        raise ValueError("No transactions provided")

    parsed = pd.to_datetime(
        pd.Series(dates),
        errors="coerce",
        format="mixed"
    )

    if parsed.isna().any():
        raise ValueError("Invalid transaction dates found")

    actual_months = parsed.dt.strftime("%Y-%m")

    if not actual_months.eq(expected_month).all():
        raise ValueError(
            "Transactions outside the reporting month"
        )

    return True


def validate_scope_approval(statuses):
    """
    Check whether all reporting-scope
    rules are approved.
    """

    if len(statuses) == 0:
        return False

    normalized = [
        str(status).strip().upper()
        for status in statuses
    ]

    return all(
        status == "APPROVED"
        for status in normalized
    )


def plan_update_decision(dates, expected_month, statuses):
    """
    Return a safe dry-run decision.
    This function never authorizes Excel writes.
    """

    try:
        validate_reporting_month(dates, expected_month)
    except ValueError:
        return "BLOCKED_INVALID_REPORTING_MONTH"

    if not validate_scope_approval(statuses):
        return "BLOCKED_PENDING_APPROVAL"

    return "REVIEW_REQUIRED"
