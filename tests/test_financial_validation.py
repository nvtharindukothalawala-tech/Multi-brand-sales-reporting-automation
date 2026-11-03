
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

# ==========================================
# LESSON 36 - TRANSACTION ACCURACY TESTS
# ==========================================

from decimal import Decimal


def detailed_sale(bill_no, amount, date_value="2026-10-10"):
    return {
        "bill_no": bill_no,
        "product": "TEST_SHIRT",
        "cost_price": 1000,
        "sales_price": amount,
        "sold_qty": 1,
        "cost_amount": 1000,
        "sales_amount": amount,
        "date": date_value,
    }


# ==========================================
# TEST 4: DETECT COMPENSATING AMOUNT ERRORS
# ==========================================

def test_reconciliation_detects_offsetting_errors():

    existing = [
        detailed_sale("TEST001", 1500),
        detailed_sale("TEST002", 2500),
    ]

    # Total is still 4000, but amounts are wrong.
    output = [
        detailed_sale("TEST001", 1400),
        detailed_sale("TEST002", 2600),
    ]

    with pytest.raises(ValueError):
        reconcile_totals(
            existing,
            [],
            output,
        )


# ==========================================
# TEST 5: DETECT CHANGED TRANSACTION DATE
# ==========================================

def test_reconciliation_detects_changed_date():

    existing = [
        detailed_sale(
            "TEST001",
            1500,
            "2026-10-10",
        ),
    ]

    # Same financial amounts, different date.
    output = [
        detailed_sale(
            "TEST001",
            1500,
            "2026-10-11",
        ),
    ]

    with pytest.raises(ValueError):
        reconcile_totals(
            existing,
            [],
            output,
        )


# ==========================================
# TEST 6: EXACT DECIMAL PRECISION
# ==========================================

def test_financial_decimal_precision():

    records = [
        {
            "sold_qty": 1,
            "cost_amount": Decimal("0.05"),
            "sales_amount": Decimal("0.10"),
        },
        {
            "sold_qty": 1,
            "cost_amount": Decimal("0.15"),
            "sales_amount": Decimal("0.20"),
        },
    ]

    totals = calculate_totals(records)

    assert totals["sold_qty"] == Decimal("2")

    assert totals["cost_amount"] == Decimal("0.20")

    assert totals["sales_amount"] == Decimal("0.30")
