
"""
Five-business Excel schema configuration.

Metadata only.
Does not open, modify, or save workbooks.
"""

BUSINESS_SCHEMAS = {
    "Be Dapper": {
        "sales": {
            "sheet": "SalesRepItem",
            "header_row": 4,
            "first_data_row": 5,
            "columns": {
                "bill_no": "Z",
                "location": "AA",
                "product": "AO",
                "sold_qty": "AT",
                "sales_amount": "AV",
                "date": "AX",
                "note": "AG",
                "exchange_ref": "AH",
            },
        },
        "returns": {
            "sheet": "Ord.Rtn",
            "header_row": 2,
            "first_data_row": 3,
            "columns": {
                "return_code": "G",
                "bill_no": "H",
                "product": "L",
                "return_qty": "P",
                "sale_amount": "U",
                "return_date": "V",
            },
        },
    },
    "Be Desi": {
        "sales": {
            "sheet": "SalesRepItem",
            "header_row": 4,
            "first_data_row": 5,
            "columns": {
                "bill_no": "AB",
                "location": "AC",
                "product": "AQ",
                "sold_qty": "AV",
                "sales_amount": "AX",
                "date": "AZ",
                "note": "AI",
                "exchange_ref": "AJ",
            },
        },
        "returns": {
            "sheet": "Ord.Rtn",
            "header_row": 2,
            "first_data_row": 3,
            "columns": {
                "return_code": "F",
                "bill_no": "G",
                "product": "K",
                "return_qty": "O",
                "sale_amount": "T",
                "return_date": "U",
            },
        },
    },
    "Giggles": {
        "sales": {
            "sheet": "SalesRepItem",
            "header_row": 4,
            "first_data_row": 5,
            "columns": {
                "bill_no": "W",
                "location": "X",
                "product": "AL",
                "sold_qty": "AQ",
                "sales_amount": "AS",
                "date": "AU",
                "note": "AD",
                "exchange_ref": "AE",
            },
        },
        "returns": {
            "sheet": "Ord.Rtn",
            "header_row": 2,
            "first_data_row": 3,
            "columns": {
                "return_code": "F",
                "bill_no": "G",
                "product": "K",
                "return_qty": "O",
                "sale_amount": "T",
                "return_date": "U",
            },
        },
    },
    "Protecto": {
        "sales": {
            "sheet": "SalesRepItem",
            "header_row": 4,
            "first_data_row": 5,
            "columns": {
                "bill_no": "Y",
                "location": "Z",
                "product": "AN",
                "sold_qty": "AS",
                "sales_amount": "AU",
                "date": "AW",
                "note": "AF",
                "exchange_ref": "AG",
            },
        },
        "returns": {
            "sheet": "Ord.Rtn",
            "header_row": 2,
            "first_data_row": 3,
            "columns": {
                "return_code": "F",
                "bill_no": "G",
                "product": "K",
                "return_qty": "O",
                "sale_amount": "T",
                "return_date": "U",
            },
        },
    },
    "Safety Sure": {
        "sales": {
            "sheet": "SalesRepItem",
            "header_row": 4,
            "first_data_row": 5,
            "columns": {
                "bill_no": "Y",
                "location": "Z",
                "product": "AN",
                "sold_qty": "AS",
                "sales_amount": "AU",
                "date": "AW",
                "note": "AF",
                "exchange_ref": "AG",
            },
        },
        "returns": {
            "sheet": "Ord.Rtn",
            "header_row": 1,
            "first_data_row": 2,
            "columns": {
                "return_code": "F",
                "bill_no": "G",
                "product": "K",
                "return_qty": "O",
                "sale_amount": "T",
                "return_date": "U",
            },
        },
    },
}
