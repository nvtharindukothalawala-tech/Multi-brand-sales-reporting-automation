
from collections import Counter
from decimal import Decimal, InvalidOperation
from datetime import datetime, date

import pandas as pd


def clean_text(value):
    """Standardize text values for matching."""

    if pd.isna(value):
        return ""

    return str(value).strip().casefold()


def clean_number(value):
    """Standardize numbers to two decimals."""

    if pd.isna(value):
        return None

    try:
        return Decimal(str(value)).quantize(
            Decimal("0.01")
        )
    except (InvalidOperation, ValueError):
        return None


def clean_date(value):
    """Convert supported date values to YYYY-MM-DD."""

    if value is None or pd.isna(value):
        return None

    if isinstance(value, (datetime, date)):
        return value.strftime("%Y-%m-%d")

    parsed = pd.to_datetime(
        str(value).strip(),
        errors="coerce"
    )

    if pd.isna(parsed):
        return None

    return parsed.strftime("%Y-%m-%d")


def make_transaction_key(record):
    """Build a candidate transaction-line matching key."""

    return (
        clean_text(record["bill_no"]),
        clean_text(record["product"]),
        clean_number(record["sold_qty"]),
        clean_number(record["cost_price"]),
        clean_number(record["sales_price"]),
        clean_number(record["cost_amount"]),
        clean_number(record["sales_amount"]),
        clean_date(record["date"]),
    )


def compare_transactions(incoming_records, existing_records):
    """Compare transaction occurrences without losing repeated lines."""

    incoming_counter = Counter(
        make_transaction_key(record)
        for record in incoming_records
    )

    existing_counter = Counter(
        make_transaction_key(record)
        for record in existing_records
    )

    matching = incoming_counter & existing_counter
    incoming_only = incoming_counter - existing_counter
    existing_only = existing_counter - incoming_counter

    return {
        "matched": sum(matching.values()),
        "new": sum(incoming_only.values()),
        "existing_only": sum(existing_only.values()),
    }
