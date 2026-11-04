"""Fictional synthetic fixtures for the real-data READ-ONLY reconciliation engine."""
from __future__ import annotations

from datetime import datetime
from pathlib import Path
from copy import deepcopy
import pytest
from openpyxl import Workbook
from openpyxl.utils.cell import column_index_from_string

from src.real_business_dry_run import (
    BRANDS, SALES_FIELDS, SALES_HEADERS, RETURN_HEADERS,
    audit, _export_report, _hash, _compare,
)


def make_fixture(tmp_path):
    sales_path=tmp_path/'sales.xlsx'; returns_path=tmp_path/'returns.xlsx'
    templates=tmp_path/'templates'; templates.mkdir()
    sw=Workbook(); sh=sw.active; sh.append(SALES_HEADERS)
    rw=Workbook(); rh=rw.active; rh.append(RETURN_HEADERS)
    for brand,conf in BRANDS.items():
        bill='TEST-'+brand.replace(' ','').upper()
        sale={f:None for f in SALES_FIELDS}
        sale.update(bill_no=bill, location=conf['sales_locations'][0], customer='TEST CUSTOMER',
            product='FICTIONAL PRODUCT', sold_qty=1, sales_amount=100, cost_amount=50,
            cost_price=50, sales_price=100, discount=0, exchange=0, delivery_amount=0,
            paid_by='TEST CREDIT',date='2026-10-12 11:30:00')
        sh.append([sale.get(f) for f in SALES_FIELDS]+[None])
        ret={'return_code':'R-'+bill,'bill_no':bill,'warehouse':conf['return_locations'][0],
             'product':'FICTIONAL PRODUCT','return_qty':1,'sale_amount':10,'return_date':'2026-10-12'}
        return_row=[None]*len(RETURN_HEADERS)
        for name,v in {'Return Code':ret['return_code'],'Bill No':ret['bill_no'],'Warehouse':ret['warehouse'],
            'Product Description':ret['product'],'Return Qty':ret['return_qty'],
            'Sale Amount':ret['sale_amount'],'Return Date':ret['return_date']}.items():
            return_row[RETURN_HEADERS.index(name)]=v
        rh.append(return_row)
        w=Workbook(); ss=w.active;ss.title='SalesRepItem';rr=w.create_sheet('Ord.Rtn')
        start=conf['sales_col']
        for i,h in enumerate(SALES_HEADERS[:25]):ss.cell(4,start+i,h)
        for i,f in enumerate(SALES_FIELDS):
            if f in ('paid_by','date'):continue
            ss.cell(5,start+i,sale[f])
        # Real template contains two numeric calculated cells between headers
        # and the actual Paid By / Date data fields.
        ss.cell(5,start+23,0)
        ss.cell(5,start+24,100)
        ss.cell(5,start+25,'TEST CREDIT')
        ss.cell(5,start+26,'2026-10-12 11:30:00')
        hr=conf['return_header_row']
        header_map={'return_code':'Return Code','bill_no':'Bill No','warehouse':'Warehouse',
              'product':'Product Description','return_qty':'Return Qty','sale_amount':'Sale Amount','return_date':'Return Date'}
        for f,letter in conf['returns_col'].items():
            idx=column_index_from_string(letter)
            rr.cell(hr,idx,header_map[f]);rr.cell(hr+1,idx,ret[f])
        w.save(templates/conf['file']);w.close()
    sw.save(sales_path);sw.close();rw.save(returns_path);rw.close()
    return sales_path,returns_path,templates


def test_read_only_five_business_full_match(tmp_path):
    sales,returns,templates=make_fixture(tmp_path)
    paths=[sales,returns,*templates.glob('*.xlsx')]
    hashes={p:_hash(p) for p in paths}
    result=audit(sales,returns,templates,'2026-10')
    assert result['status']=='NO_INSERTS_NEEDED_REVIEW_EXTRAS'
    assert result['source_sales']['records']==5
    assert result['source_returns']['records']==5
    assert result['financial_partition_check'] is True
    assert result['inputs_unchanged_sha256_verified'] is True
    assert all(d[k]['matched_identity']==1 and d[k]['full_field_different_matched_rows']==0
        for d in result['businesses'].values() for k in ('sales','returns'))
    assert hashes=={p:_hash(p) for p in paths}


def test_report_exports_only_aggregate_data_and_blocks_overwrite(tmp_path):
    sales,returns,templates=make_fixture(tmp_path)
    result=audit(sales,returns,templates,'2026-10')
    output=tmp_path/'private_report'
    _export_report(result,output)
    files=sorted(p.name for p in output.iterdir())
    assert files==['READ_ME_FIRST.txt','business_summary.csv','field_differences.csv','real_data_audit.json']
    for file in output.iterdir():
        assert 'TEST CUSTOMER' not in file.read_text(encoding='utf-8-sig')
    with pytest.raises(FileExistsError): _export_report(result,output)


def test_unexpected_month_fails_closed(tmp_path):
    sales,returns,templates=make_fixture(tmp_path)
    with pytest.raises(ValueError,match='outside'):
        audit(sales,returns,templates,'2026-09')


def test_missing_source_sale_reports_missing_not_appended(tmp_path):
    sales,returns,templates=make_fixture(tmp_path)
    wb=Workbook();ws=wb.active;ws.title='SalesRepItem';
    # Template file still required to match known headers: modify a fictional
    # core transaction to make the source transaction absent.
    from openpyxl import load_workbook
    selected=templates/BRANDS['Be Dapper']['file']
    w=load_workbook(selected);ss=w['SalesRepItem']
    ss.cell(5,BRANDS['Be Dapper']['sales_col']).value='DIFFERENT-BILL'
    w.save(selected);w.close()
    r=audit(sales,returns,templates,'2026-10')
    assert r['businesses']['Be Dapper']['sales']['missing_from_template']==1
    assert r['status']=='REVIEW_REQUIRED'


def test_detect_changed_nonkey_field(tmp_path):
    sales,returns,templates=make_fixture(tmp_path)
    from openpyxl import load_workbook
    selected=templates/BRANDS['Giggles']['file']; w=load_workbook(selected)
    ss=w['SalesRepItem'];ss.cell(5,BRANDS['Giggles']['sales_col']+2,'CHANGED CUSTOMER')
    w.save(selected);w.close()
    r=audit(sales,returns,templates,'2026-10')
    assert r['businesses']['Giggles']['sales']['matched_identity']==1
    assert r['businesses']['Giggles']['sales']['field_difference_counts']['customer']==1
    assert r['status']=='REVIEW_REQUIRED'


def test_extra_other_location_is_not_deleted(tmp_path):
    sales,returns,templates=make_fixture(tmp_path)
    from openpyxl import load_workbook
    selected=templates/BRANDS['Be Desi']['file'];w=load_workbook(selected);ss=w['SalesRepItem']
    first=BRANDS['Be Desi']['sales_col']
    for i in range(25): ss.cell(6,first+i,ss.cell(5,first+i).value)
    ss.cell(6,first,'EXTRA-BILL')
    ss.cell(6,first+1,'Raja & Rani Shop - Raja & Rani Shop')
    ss.cell(6,first+25,'TEST CREDIT');ss.cell(6,first+26,'2026-10-12')
    w.save(selected);w.close()
    r=audit(sales,returns,templates,'2026-10')
    v=r['businesses']['Be Desi']['sales']
    assert v['additional_in_template']==1
    assert v['additional_location_analysis']['additional_other_locations']==1
    assert r['status']=='NO_INSERTS_NEEDED_REVIEW_EXTRAS'


def test_invalid_header_is_rejected(tmp_path):
    sales,returns,templates=make_fixture(tmp_path)
    from openpyxl import load_workbook
    selected=templates/BRANDS['Protecto']['file'];w=load_workbook(selected);w['SalesRepItem'].cell(4,BRANDS['Protecto']['sales_col'],'WRONG HEADER');w.save(selected);w.close()
    with pytest.raises(ValueError,match='header mismatch'):
        audit(sales,returns,templates,'2026-10')
