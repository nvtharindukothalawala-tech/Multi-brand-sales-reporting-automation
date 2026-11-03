
# ==========================================
# FINANCIAL VALIDATION ENGINE
# LESSON 37
# ==========================================

from collections import Counter
from datetime import date, datetime
from decimal import Decimal, InvalidOperation

from src.transaction_matching import make_transaction_key


# ==========================================
# 1. CONFIGURATION
# ==========================================

FINANCIAL_FIELDS = (
    "sold_qty",
    "cost_amount",
    "sales_amount",
)

NUMERIC_FIELDS = {
    "sold_qty",
    "cost_price",
    "sales_price",
    "cost_amount",
    "sales_amount",
}

TRANSACTION_FIELDS = {
    "bill_no",
    "product",
    "cost_price",
    "sales_price",
    "sold_qty",
    "cost_amount",
    "sales_amount",
    "date",
}


# ==========================================
# 2. CALCULATE FINANCIAL TOTALS
# ==========================================

def calculate_totals(records):
    """
    Calculate totals using Decimal.

    Reject missing, invalid, or non-finite
    financial values.
    """

    totals = {
        field: Decimal("0")
        for field in FINANCIAL_FIELDS
    }

    for row_number, record in enumerate(
        records,
        start=1,
    ):

        for field in FINANCIAL_FIELDS:

            try:
                amount = Decimal(
                    str(record[field])
                )

                if not amount.is_finite():
                    raise ValueError(
                        "Non-finite number"
                    )

            except (
                KeyError,
                InvalidOperation,
                ValueError,
                TypeError,
            ):
                raise ValueError(
                    f"Invalid {field} "
                    f"in record {row_number}"
                ) from None

            totals[field] += amount

    return totals


# ==========================================
# 3. NORMALIZE PARTIAL TEST RECORDS
# ==========================================

def normalize_value(field, value):
    """
    Normalize values for record comparison.

    Supports partial synthetic test records
    and simple Excel date representations.
    """

    if value is None:
        return None

    if field in NUMERIC_FIELDS:

        try:
            number = Decimal(str(value))

            if not number.is_finite():
                raise ValueError(
                    f"Invalid numeric value: {field}"
                )

            return number.normalize()

        except (
            InvalidOperation,
            ValueError,
            TypeError,
        ):
            raise ValueError(
                f"Invalid numeric value: {field}"
            ) from None

    if field == "date":

        if isinstance(value, datetime):
            return value.isoformat()

        if isinstance(value, date):
            return value.isoformat()

        text = str(value).strip()

        try:
            return datetime.fromisoformat(
                text
            ).isoformat()

        except ValueError:
            return text

    if isinstance(value, str):
        return " ".join(
            value.strip().upper().split()
        )

    return str(value)


# ==========================================
# 4. CREATE TRANSACTION COMPARISON KEY
# ==========================================

def create_record_key(record):
    """
    Create a key for transaction-level checks.

    Use the existing transaction matching engine
    for full transaction records.

    Use a normalized key for partial test records.
    """

    if not isinstance(record, dict):
        raise ValueError(
            "Transaction must be a dictionary."
        )

    # Reuse the existing transaction matching
    # logic for complete sales records.
    if TRANSACTION_FIELDS.issubset(record.keys()):

        return (
            "full",
            make_transaction_key(record),
        )

    # Partial records are used by the earlier
    # financial unit tests.
    return (
        "partial",
        tuple(
            (
                field,
                normalize_value(
                    field,
                    record[field],
                ),
            )
            for field in sorted(record.keys())
        ),
    )


# ==========================================
# 5. VALIDATE INDIVIDUAL TRANSACTIONS
# ==========================================

def verify_transaction_records(
    existing_records,
    inserted_records,
    output_records,
):
    """
    Verify that the output contains exactly
    the expected transaction occurrences.

    Detect:
    - Changed amounts
    - Changed dates
    - Missing transactions
    - Extra transactions
    - Incorrect duplicate occurrences
    """

    expected_records = (
        list(existing_records)
        + list(inserted_records)
    )

    expected_counter = Counter(
        create_record_key(record)
        for record in expected_records
    )

    actual_counter = Counter(
        create_record_key(record)
        for record in output_records
    )

    if expected_counter != actual_counter:

        missing = (
            expected_counter - actual_counter
        )

        unexpected = (
            actual_counter - expected_counter
        )

        raise ValueError(
            "Transaction-level reconciliation "
            "failed. "
            f"Missing occurrences: {sum(missing.values())}. "
            f"Unexpected occurrences: "
            f"{sum(unexpected.values())}."
        )

    return True


# ==========================================
# 6. COMPLETE FINANCIAL RECONCILIATION
# ==========================================

def reconcile_totals(
    existing_records,
    inserted_records,
    output_records,
):
    """
    Verify:
    1. Transaction row counts
    2. Financial quantities and amounts
    3. Individual transaction occurrences

    Raise ValueError when any check fails.
    """

    # --------------------------------------
    # CHECK 1: CALCULATE TOTALS
    # --------------------------------------

    existing = calculate_totals(
        existing_records
    )

    inserted = calculate_totals(
        inserted_records
    )

    actual = calculate_totals(
        output_records
    )

    # --------------------------------------
    # CHECK 2: COMPARE FINANCIAL TOTALS
    # --------------------------------------

    differences = {}

    for field in FINANCIAL_FIELDS:

        expected = (
            existing[field] + inserted[field]
        )

        differences[field] = (
            actual[field] - expected
        )

    if any(
        value != 0
        for value in differences.values()
    ):
        raise ValueError(
            "Financial reconciliation failed: "
            f"{differences}"
        )

    # --------------------------------------
    # CHECK 3: COMPARE ROW COUNTS
    # --------------------------------------

    expected_rows = (
        len(existing_records)
        + len(inserted_records)
    )

    if len(output_records) != expected_rows:
        raise ValueError(
            "Transaction row count "
            "reconciliation failed."
        )

    # --------------------------------------
    # CHECK 4: VERIFY EVERY TRANSACTION
    # --------------------------------------

    verify_transaction_records(
        existing_records,
        inserted_records,
        output_records,
    )

    # --------------------------------------
    # RECONCILIATION SUCCESSFUL
    # --------------------------------------

    return {
        "reconciled": True,
        "existing_totals": existing,
        "inserted_totals": inserted,
        "output_totals": actual,
        "differences": differences,
    }
