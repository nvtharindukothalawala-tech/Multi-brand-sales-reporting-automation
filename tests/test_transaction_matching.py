
from src.transaction_matching import (
    make_transaction_key,
    compare_transactions,
)


def sample_sale(
    bill="INV001",
    product="Shirt",
    amount=2500,
):
    return {
        "bill_no": bill,
        "product": product,
        "sold_qty": 1,
        "cost_price": 1500,
        "sales_price": 2500,
        "cost_amount": 1500,
        "sales_amount": amount,
        "date": "2026-10-01",
    }


def test_identical_transactions_match():

    incoming = [sample_sale()]
    existing = [sample_sale()]

    result = compare_transactions(
        incoming, existing
    )

    assert result["matched"] == 1
    assert result["new"] == 0
    assert result["existing_only"] == 0


def test_new_transaction_detected():

    incoming = [sample_sale(bill="INV002")]
    existing = [sample_sale(bill="INV001")]

    result = compare_transactions(
        incoming, existing
    )

    assert result["matched"] == 0
    assert result["new"] == 1
    assert result["existing_only"] == 1


def test_repeated_identical_lines():

    incoming = [
        sample_sale(),
        sample_sale(),
    ]

    existing = [sample_sale()]

    result = compare_transactions(
        incoming, existing
    )

    assert result["matched"] == 1
    assert result["new"] == 1
    assert result["existing_only"] == 0


def test_text_normalization():

    first = sample_sale(
        bill=" INV001 ",
        product="SHIRT"
    )

    second = sample_sale(
        bill="inv001",
        product="shirt"
    )

    assert make_transaction_key(
        first
    ) == make_transaction_key(second)


def test_different_sales_amount():

    incoming = [
        sample_sale(amount=2600)
    ]

    existing = [
        sample_sale(amount=2500)
    ]

    result = compare_transactions(
        incoming, existing
    )

    assert result["matched"] == 0
    assert result["new"] == 1
