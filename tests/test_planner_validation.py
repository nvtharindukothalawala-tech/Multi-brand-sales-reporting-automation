
import pytest

from src.planner_validation import (
    validate_reporting_month,
    validate_scope_approval,
    plan_update_decision,
)


def test_correct_reporting_month():

    dates = [
        "2026-10-01",
        "2026-10-15",
        "2026-10-31",
    ]

    assert validate_reporting_month(
        dates, "2026-10"
    ) is True


def test_wrong_reporting_month():

    dates = [
        "2026-10-01",
        "2026-11-01",
    ]

    with pytest.raises(ValueError):
        validate_reporting_month(
            dates, "2026-10"
        )


def test_invalid_date():

    with pytest.raises(ValueError):
        validate_reporting_month(
            ["invalid-date"], "2026-10"
        )


def test_pending_business_approval():

    statuses = [
        "APPROVED",
        "REVIEW",
    ]

    assert validate_scope_approval(
        statuses
    ) is False


def test_all_approvals():

    statuses = [
        "APPROVED",
        "APPROVED",
    ]

    assert validate_scope_approval(
        statuses
    ) is True


def test_missing_approval_rules():

    assert validate_scope_approval([]) is False


def test_invalid_month_blocks_update():

    result = plan_update_decision(
        ["2026-11-01"],
        "2026-10",
        ["APPROVED"]
    )

    assert result == "BLOCKED_INVALID_REPORTING_MONTH"


def test_approved_scope_still_requires_review():

    result = plan_update_decision(
        ["2026-10-01"],
        "2026-10",
        ["APPROVED"]
    )

    assert result == "REVIEW_REQUIRED"
