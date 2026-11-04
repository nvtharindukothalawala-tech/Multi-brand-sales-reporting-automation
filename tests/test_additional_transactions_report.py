from io import StringIO
from src.additional_transactions_report import _find_additional, _safe, _csv


def sale(bill, product='P1'):
    return {'bill_no':bill,'product':product,'sold_qty':1,'sales_amount':100}


def test_additional_rows_are_multiset_differences():
    src=[sale('A')]
    existing=[sale('A'),sale('A'),sale('B')]
    assert [r['bill_no'] for r in _find_additional(src,existing,'sales')] == ['A','B']


def test_no_extra_rows_when_same_transactions():
    assert _find_additional([sale('A'),sale('B')],[sale('B'),sale('A')],'sales') == []


def test_csv_formula_values_escaped():
    assert _safe('=CMD()') == "'=CMD()"
    assert "'=CMD()" in _csv([{'Item':'=CMD()'}],('Item',)).decode('utf-8-sig')


def test_invalid_kind_fails():
    import pytest
    with pytest.raises(ValueError):
        _find_additional([],[],'bad')
