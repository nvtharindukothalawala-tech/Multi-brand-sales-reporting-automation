
from decimal import Decimal, InvalidOperation


def calculate_totals(records):
    """
    Calculate transaction quantity and
    financial totals using Decimal.
    """

    totals = {
        "sold_qty": Decimal("0"),
        "cost_amount": Decimal("0"),
        "sales_amount": Decimal("0"),
    }

    for row_number, record in enumerate(records, start=1):
        for field in totals:

            try:
                amount = Decimal(str(record[field]))

                if not amount.is_finite():
                    raise ValueError("Non-finite number")

            except (
                KeyError,
                InvalidOperation,
                ValueError,
            ):
                raise ValueError(
                    f"Invalid {field} in record {row_number}"
                ) from None

            totals[field] += amount

    return totals


def reconcile_totals(
    existing_records,
    inserted_records,
    output_records,
):
    """
    Check output = existing + inserted.

    Raise an error if totals do not match.
    """

    existing = calculate_totals(existing_records)
    inserted = calculate_totals(inserted_records)
    actual = calculate_totals(output_records)

    differences = {}

    for field in existing:

        expected = existing[field] + inserted[field]
        differences[field] = actual[field] - expected

    if any(value != 0 for value in differences.values()):
        raise ValueError(
            f"Financial reconciliation failed: {differences}"
        )

    expected_rows = (
        len(existing_records) + len(inserted_records)
    )

    if len(output_records) != expected_rows:
        raise ValueError(
            "Transaction row count reconciliation failed."
        )

    return {
        "reconciled": True,
        "existing_totals": existing,
        "inserted_totals": inserted,
        "output_totals": actual,
        "differences": differences,
    }
