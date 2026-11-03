
from pathlib import Path

from openpyxl import Workbook, load_workbook

from src.transaction_matching import compare_transactions


# ----------------------------------------
# 1. CREATE SYNTHETIC TRANSACTION
# ----------------------------------------

def sample_sale(bill_no, product="Shirt"):
    return {
        "bill_no": bill_no,
        "product": product,
        "sold_qty": 1,
        "cost_price": 1500,
        "sales_price": 2500,
        "cost_amount": 1500,
        "sales_amount": 2500,
        "date": "2026-10-01"
    }


# ----------------------------------------
# 2. CREATE TEMPORARY EXCEL WORKBOOK
# ----------------------------------------

def create_test_workbook(file_path, records):

    wb = Workbook()
    ws = wb.active
    ws.title = "SalesRepItem"

    # Match the structure of the Be Dapper
    # sales transaction area.
    columns = {
        "bill_no": "Z",
        "product": "AO",
        "cost_price": "AP",
        "sales_price": "AQ",
        "sold_qty": "AT",
        "cost_amount": "AU",
        "sales_amount": "AV",
        "date": "AZ"
    }

    for row_number, record in enumerate(
        records, start=5
    ):
        for field, column in columns.items():
            ws[f"{column}{row_number}"] = record[field]

    wb.save(file_path)
    wb.close()


# ----------------------------------------
# 3. READ EXCEL TRANSACTIONS
# ----------------------------------------

def read_test_workbook(file_path):

    wb = load_workbook(
        file_path,
        read_only=True,
        data_only=False
    )

    records = []

    try:
        ws = wb["SalesRepItem"]

        for row in ws.iter_rows(
            min_row=5,
            values_only=True
        ):
            if len(row) < 52 or row[25] is None:
                continue

            records.append({
                "bill_no": row[25],
                "product": row[40],
                "cost_price": row[41],
                "sales_price": row[42],
                "sold_qty": row[45],
                "cost_amount": row[46],
                "sales_amount": row[47],
                "date": row[51]
            })

    finally:
        wb.close()

    return records


# ----------------------------------------
# 4. TEST: ALL RECORDS MATCH
# ----------------------------------------

def test_excel_all_transactions_match(tmp_path):

    excel_file = tmp_path / "test_sales.xlsx"

    records = [
        sample_sale("INV001"),
        sample_sale("INV002")
    ]

    create_test_workbook(excel_file, records)

    existing = read_test_workbook(excel_file)

    result = compare_transactions(records, existing)

    assert result["matched"] == 2
    assert result["new"] == 0
    assert result["existing_only"] == 0


# ----------------------------------------
# 5. TEST: DETECT NEW TRANSACTION
# ----------------------------------------

def test_excel_detects_new_sale(tmp_path):

    excel_file = tmp_path / "test_sales.xlsx"

    existing_records = [
        sample_sale("INV001")
    ]

    incoming_records = [
        sample_sale("INV001"),
        sample_sale("INV002")
    ]

    create_test_workbook(
        excel_file,
        existing_records
    )

    existing = read_test_workbook(excel_file)

    result = compare_transactions(
        incoming_records,
        existing
    )

    assert result["matched"] == 1
    assert result["new"] == 1
    assert result["existing_only"] == 0


# ----------------------------------------
# 6. TEST: KEEP REPEATED PRODUCT LINES
# ----------------------------------------

def test_excel_repeated_lines(tmp_path):

    excel_file = tmp_path / "test_sales.xlsx"

    existing_records = [
        sample_sale("INV001")
    ]

    incoming_records = [
        sample_sale("INV001"),
        sample_sale("INV001")
    ]

    create_test_workbook(
        excel_file,
        existing_records
    )

    existing = read_test_workbook(excel_file)

    result = compare_transactions(
        incoming_records,
        existing
    )

    assert result["matched"] == 1
    assert result["new"] == 1
    assert result["existing_only"] == 0
