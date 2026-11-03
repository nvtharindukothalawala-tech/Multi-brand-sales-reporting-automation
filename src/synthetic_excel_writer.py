
from collections import Counter
from pathlib import Path

from openpyxl import load_workbook

from src.transaction_matching import make_transaction_key


# Only synthetic workbooks with this marker are allowed.
TEST_MARKER = "SYNTHETIC_TEST_ONLY"

COLUMN_MAP = {
    "bill_no": "Z",
    "product": "AO",
    "cost_price": "AP",
    "sales_price": "AQ",
    "sold_qty": "AT",
    "cost_amount": "AU",
    "sales_amount": "AV",
    "date": "AZ",
}

COLUMN_INDEXES = {
    "bill_no": 25,
    "product": 40,
    "cost_price": 41,
    "sales_price": 42,
    "sold_qty": 45,
    "cost_amount": 46,
    "sales_amount": 47,
    "date": 51,
}


def read_existing_records(ws):
    """Read the synthetic workbook transactions."""

    records = []

    for row in ws.iter_rows(
        min_row=5,
        values_only=True,
    ):
        if len(row) < 52 or row[25] is None:
            continue

        record = {
            field: row[index]
            for field, index in COLUMN_INDEXES.items()
        }

        records.append(record)

    return records


def write_synthetic_copy(
    source_path,
    output_path,
    incoming_records,
):
    """
    Create a new workbook containing missing
    synthetic transactions.

    Never overwrite the input workbook.
    """

    source_path = Path(source_path).resolve()
    output_path = Path(output_path).resolve()

    if source_path == output_path:
        raise ValueError(
            "Input and output must be different files."
        )

    if not source_path.is_file():
        raise FileNotFoundError(source_path)

    if source_path.suffix.lower() != ".xlsx":
        raise ValueError("Only .xlsx files are supported.")

    if output_path.exists():
        raise FileExistsError(
            "Output already exists. Refusing overwrite."
        )

    wb = load_workbook(source_path)

    try:
        if "SalesRepItem" not in wb.sheetnames:
            raise ValueError(
                "SalesRepItem sheet is missing."
            )

        ws = wb["SalesRepItem"]

        if ws["A1"].value != TEST_MARKER:
            raise ValueError(
                "Not a synthetic test workbook. "
                "Writing is blocked."
            )

        existing_records = read_existing_records(ws)

        # Count matching transaction occurrences.
        remaining_existing = Counter(
            make_transaction_key(record)
            for record in existing_records
        )

        new_records = []
        matched_count = 0

        for record in incoming_records:
            key = make_transaction_key(record)

            if remaining_existing[key] > 0:
                remaining_existing[key] -= 1
                matched_count += 1
            else:
                new_records.append(record)

        # Add only missing transaction occurrences.
        next_row = max(5, ws.max_row + 1)

        for record in new_records:
            for field, column in COLUMN_MAP.items():
                ws[f"{column}{next_row}"] = record[field]

            next_row += 1

        output_path.parent.mkdir(
            parents=True,
            exist_ok=True,
        )

        wb.save(output_path)

    finally:
        wb.close()

    return {
        "matched": matched_count,
        "inserted": len(new_records),
        "existing_before": len(existing_records),
        "existing_after": (
            len(existing_records) + len(new_records)
        ),
    }
