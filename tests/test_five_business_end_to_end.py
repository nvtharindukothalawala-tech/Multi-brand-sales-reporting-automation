"""End-to-end integration for synthetic FIVE-business sales/returns workbooks.

No real business data, no private paths, no writes outside pytest tmp_path.
The existing writer and configuration are imported but never modified.
"""

from collections import Counter
from datetime import date, datetime
from decimal import Decimal
from hashlib import sha256

import pytest
from openpyxl import Workbook, load_workbook

from src.business_schema import BUSINESS_SCHEMAS
from src.schema_aware_synthetic_writer import write_synthetic_business_workbook

EXPECTED_BUSINESSES = {"Be Dapper", "Be Desi", "Giggles", "Protecto", "Safety Sure"}
HEADER_LABELS = {
    "sales": {
        "bill_no": "Bill No", "product": "Product", "sold_qty": "Sold Qty",
        "sales_amount": "Sales Amount", "date": "Date",
    },
    "returns": {
        "return_code": "Return Code", "bill_no": "Bill No",
        "product": "Product Description", "return_qty": "Return Qty",
        "sale_amount": "Sale Amount", "return_date": "Return Date",
    },
}


def create_fictional_template(path, business, marked=True):
    """Create a minimal workbook from scratch; never copy a company template."""
    wb = Workbook()
    sales = wb.active
    sales.title = "SalesRepItem"
    returns = wb.create_sheet("Ord.Rtn")
    if marked:
        sales["A1"] = "SYNTHETIC_TEST_ONLY"
    for kind, ws in (("sales", sales), ("returns", returns)):
        config = BUSINESS_SCHEMAS[business][kind]
        for field, label in HEADER_LABELS[kind].items():
            ws[f"{config['columns'][field]}{config['header_row']}"] = label
    wb.save(path)
    wb.close()


def sales_for(business):
    slug = business.replace(" ", "").upper()
    return [
        {"bill_no": f"SYN-{slug}-01", "product": "Fictional Blue Shirt", "sold_qty": 2,
         "sales_amount": "1500.25", "date": "2026-10-08"},
        {"bill_no": f"SYN-{slug}-02", "product": "Fictional Red Shirt", "sold_qty": 3,
         "sales_amount": "2400.50", "date": "2026-10-10"},
        {"bill_no": f"SYN-{slug}-01", "product": "Fictional Blue Shirt", "sold_qty": 2,
         "sales_amount": "1500.25", "date": "2026-10-08"},
    ]


def returns_for(business):
    slug = business.replace(" ", "").upper()
    return [
        {"return_code": f"RTN-{slug}-01", "bill_no": f"SYN-{slug}-01",
         "product": "Fictional Blue Shirt", "return_qty": 1,
         "sale_amount": "750.25", "return_date": "2026-10-15"},
        {"return_code": f"RTN-{slug}-02", "bill_no": f"SYN-{slug}-02",
         "product": "Fictional Red Shirt", "return_qty": 1,
         "sale_amount": "800.50", "return_date": "2026-10-16"},
    ]


def read_rows(book, business, kind):
    config = BUSINESS_SCHEMAS[business][kind]
    ws = book[config["sheet"]]
    fields = tuple(HEADER_LABELS[kind])
    rows = []
    for row_num in range(config["first_data_row"], ws.max_row + 1):
        values = {f: ws[f"{config['columns'][f]}{row_num}"].value for f in fields}
        if any(value is not None for value in values.values()):
            rows.append(values)
    return rows


def normalize_record(record, kind):
    text_fields = ("bill_no", "product") if kind == "sales" else ("return_code", "bill_no", "product")
    qty_field = "sold_qty" if kind == "sales" else "return_qty"
    money_field = "sales_amount" if kind == "sales" else "sale_amount"
    date_field = "date" if kind == "sales" else "return_date"
    date_value = record[date_field]
    if isinstance(date_value, (date, datetime)):
        date_value = date_value.isoformat()[:10]
    return (
        *(str(record[f]).strip() for f in text_fields),
        int(record[qty_field]),
        Decimal(str(record[money_field])),
        date_value,
    )


def assert_financial_reconciliation(rows, expected, kind):
    assert Counter(normalize_record(row, kind) for row in rows) == Counter(
        normalize_record(row, kind) for row in expected
    )
    qty = "sold_qty" if kind == "sales" else "return_qty"
    amount = "sales_amount" if kind == "sales" else "sale_amount"
    assert sum(int(row[qty]) for row in rows) == sum(int(row[qty]) for row in expected)
    assert sum(Decimal(str(row[amount])) for row in rows) == sum(
        Decimal(str(row[amount])) for row in expected
    )


@pytest.mark.parametrize("business", sorted(EXPECTED_BUSINESSES))
def test_five_business_sales_returns_full_chain(tmp_path, business):
    """Create -> append sales -> append returns -> read back -> rerun both idempotently."""
    assert set(BUSINESS_SCHEMAS) == EXPECTED_BUSINESSES
    source = tmp_path / "fictional_input.xlsx"
    sales_output = tmp_path / "sales_output.xlsx"
    final_output = tmp_path / "combined_output.xlsx"
    rerun_sales = tmp_path / "rerun_sales.xlsx"
    rerun_returns = tmp_path / "rerun_returns.xlsx"
    create_fictional_template(source, business)
    original_digest = sha256(source.read_bytes()).hexdigest()

    sales = sales_for(business)
    returns = returns_for(business)
    sales_result = write_synthetic_business_workbook(source, sales_output, business, "sales", sales)
    assert sales_result == {"existing": 0, "requested": 3, "added": 3}
    return_result = write_synthetic_business_workbook(sales_output, final_output, business, "returns", returns)
    assert return_result == {"existing": 0, "requested": 2, "added": 2}
    assert sha256(source.read_bytes()).hexdigest() == original_digest

    book = load_workbook(final_output, data_only=False)
    try:
        assert book["SalesRepItem"]["A1"].value == "SYNTHETIC_TEST_ONLY"
        assert_financial_reconciliation(read_rows(book, business, "sales"), sales, "sales")
        assert_financial_reconciliation(read_rows(book, business, "returns"), returns, "returns")
    finally:
        book.close()

    # Rerunning with identical transactions must not insert any duplicate occurrence.
    again_sales = write_synthetic_business_workbook(final_output, rerun_sales, business, "sales", sales)
    assert again_sales == {"existing": 3, "requested": 3, "added": 0}
    again_returns = write_synthetic_business_workbook(rerun_sales, rerun_returns, business, "returns", returns)
    assert again_returns == {"existing": 2, "requested": 2, "added": 0}
    book = load_workbook(rerun_returns, data_only=False)
    try:
        assert_financial_reconciliation(read_rows(book, business, "sales"), sales, "sales")
        assert_financial_reconciliation(read_rows(book, business, "returns"), returns, "returns")
    finally:
        book.close()


@pytest.mark.parametrize("business", sorted(EXPECTED_BUSINESSES))
def test_unmarked_workbook_is_never_published(tmp_path, business):
    source = tmp_path / "unmarked.xlsx"
    output = tmp_path / "blocked.xlsx"
    create_fictional_template(source, business, marked=False)
    old_hash = sha256(source.read_bytes()).hexdigest()
    with pytest.raises(PermissionError, match="SYNTHETIC_TEST_ONLY"):
        write_synthetic_business_workbook(source, output, business, "sales", sales_for(business))
    assert not output.exists()
    assert sha256(source.read_bytes()).hexdigest() == old_hash


def test_no_cross_business_rows_in_complete_batch(tmp_path):
    """Simulate a monthly batch across all five distinct output workbooks."""
    assert set(BUSINESS_SCHEMAS) == EXPECTED_BUSINESSES
    summaries = {}
    all_bills = set()
    for i, business in enumerate(sorted(EXPECTED_BUSINESSES)):
        brand_dir = tmp_path / f"brand_{i}"
        brand_dir.mkdir()
        source = brand_dir / "fictional_source.xlsx"
        sales_output = brand_dir / "with_sales.xlsx"
        completed = brand_dir / "with_sales_and_returns.xlsx"
        create_fictional_template(source, business)
        sales = sales_for(business)
        returns = returns_for(business)
        write_synthetic_business_workbook(source, sales_output, business, "sales", sales)
        write_synthetic_business_workbook(sales_output, completed, business, "returns", returns)
        workbook = load_workbook(completed)
        try:
            sales_rows = read_rows(workbook, business, "sales")
            return_rows = read_rows(workbook, business, "returns")
            assert_financial_reconciliation(sales_rows, sales, "sales")
            assert_financial_reconciliation(return_rows, returns, "returns")
            bills = {record["bill_no"] for record in sales_rows}
            assert not (bills & all_bills), f"Cross-business bill collision: {business}"
            all_bills.update(bills)
            sales_total = sum(Decimal(str(r["sales_amount"])) for r in sales_rows)
            returns_total = sum(Decimal(str(r["sale_amount"])) for r in return_rows)
            summaries[business] = (len(sales_rows), len(return_rows), sales_total - returns_total)
        finally:
            workbook.close()
    assert len(summaries) == 5
    assert sum(v[0] for v in summaries.values()) == 15
    assert sum(v[1] for v in summaries.values()) == 10
    assert sum(v[2] for v in summaries.values()) == Decimal("19251.25")


def test_output_collision_keeps_prior_file(tmp_path):
    business = "Giggles"
    source = tmp_path / "source.xlsx"
    output = tmp_path / "output.xlsx"
    create_fictional_template(source, business)
    output.write_bytes(b"existing output must not change")
    with pytest.raises(ValueError, match="overwriting"):
        write_synthetic_business_workbook(source, output, business, "sales", sales_for(business))
    assert output.read_bytes() == b"existing output must not change"
