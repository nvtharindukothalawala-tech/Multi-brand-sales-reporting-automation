
import pytest

from src.financial_validation import (
    calculate_totals,
    reconcile_totals,
)


def sale(bill, quantity, amount):
    return {
        "bill_no": bill,
        "sold_qty": quantity,
        "cost_amount": amount * 0.6,
        "sales_amount": amount,
    }


def test_financial_totals():

    records = [
        sale("TEST001", 1, 1500),
        sale("TEST002", 2, 3000),
    ]

    totals = calculate_totals(records)

    assert totals["sold_qty"] == 3
    assert totals["sales_amount"] == 4500
    assert totals["cost_amount"] == 2700


def test_successful_reconciliation():

    existing = [sale("TEST001", 1, 1500)]
    inserted = [sale("TEST002", 1, 1500)]

    output = existing + inserted

    result = reconcile_totals(
        existing,
        inserted,
        output,
    )

    assert result["reconciled"] is True
    assert result["output_totals"]["sales_amount"] == 3000
    assert result["differences"]["sales_amount"] == 0


def test_reconciliation_detects_wrong_amount():

    existing = [sale("TEST001", 1, 1500)]
    inserted = [sale("TEST002", 1, 1500)]

    output = [
        sale("TEST001", 1, 1500),
        sale("TEST002", 1, 1400),
    ]

    with pytest.raises(ValueError):
        reconcile_totals(
            existing,
            inserted,
            output,
        )
