"""Confidential, READ-ONLY five-brand workbook reconciliation.

Only openpyxl.load_workbook(read_only=True) is used for input files.
No original workbook is ever saved or copied into an output workbook.
Output contains aggregate metrics, not personal customer data.
"""
from __future__ import annotations

import argparse
import csv
import hashlib
import json
from collections import Counter, defaultdict
from datetime import date, datetime
from decimal import Decimal, InvalidOperation
from pathlib import Path
from tempfile import TemporaryDirectory
from zipfile import BadZipFile

from openpyxl import load_workbook
from openpyxl.utils.datetime import from_excel

SALES_HEADERS = ('Bill No','Location','Customer','Address','Created By','City','Sales Channel','Note','Exchange Ref','Contact No1','Contact No2','Sales Person','Invoice Type','Reference','Delivery Amount','Product','Cost Price','Sales Price','Discount','Exchange','Sold Qty','Cost Amount','Sales Amount','Paid By','Date','Serials')
RETURN_HEADERS = ('#','Outlet','Return Code','Bill No','Warehouse','Customer','Sales Person','Product Description','Cost Price','Sale Price','Discount','Return Qty','Remarks','Payment','Delivery Amount','Cost Amount','Sale Amount','Return Date','Created On','Created By')
SALES_FIELDS = ('bill_no','location','customer','address','created_by','city','sales_channel','note','exchange_ref','contact_no1','contact_no2','sales_person','invoice_type','reference','delivery_amount','product','cost_price','sales_price','discount','exchange','sold_qty','cost_amount','sales_amount','paid_by','date')
RETURN_FIELDS = ('return_code','bill_no','warehouse','product','return_qty','sale_amount','return_date')
SALES_MONEY = frozenset(('delivery_amount','cost_price','sales_price','discount','exchange','cost_amount','sales_amount'))
RETURN_MONEY = frozenset(('sale_amount',))
BRANDS = {
 'Be Dapper': {'file':'BeDapper_ChannelWise_Report.xlsx','sales_col':26,'returns_col':{'return_code':'G','bill_no':'H','warehouse':'I','product':'L','return_qty':'P','sale_amount':'U','return_date':'V'},'return_header_row':2,'sales_locations':('Be Dapper Shop - Be Dapper Shop','Be Dapper Wellawatte Shop - Be Dapper Wellawatte Shop'),'return_locations':('Warehouse - Be Dapper Shop',)},
 'Be Desi': {'file':'BeDesi_ChannelWise_Report.xlsx','sales_col':28,'returns_col':{'return_code':'F','bill_no':'G','warehouse':'H','product':'K','return_qty':'O','sale_amount':'T','return_date':'U'},'return_header_row':2,'sales_locations':('Be Desi Shop - Be Desi Shop',),'return_locations':('Warehouse - Be Desi Shop',)},
 'Giggles': {'file':'Giggles_ChannelWise_Report.xlsx','sales_col':23,'returns_col':{'return_code':'F','bill_no':'G','warehouse':'H','product':'K','return_qty':'O','sale_amount':'T','return_date':'U'},'return_header_row':2,'sales_locations':('Giggles Shop - Giggles Shop',),'return_locations':('Warehouse - Giggles Shop','Stock - WH - Giggles Dehiwala')},
 'Protecto': {'file':'Protecto_ChannelWise_Report.xlsx','sales_col':25,'returns_col':{'return_code':'F','bill_no':'G','warehouse':'H','product':'K','return_qty':'O','sale_amount':'T','return_date':'U'},'return_header_row':2,'sales_locations':('Protecto Shop - Protecto Shop',),'return_locations':('Warehouse - Protecto Shop','Stock - WH - Protecto Dehiwala')},
 'Safety Sure': {'file':'SafetySure_ChannelWise_Report.xlsx','sales_col':25,'returns_col':{'return_code':'F','bill_no':'G','warehouse':'H','product':'K','return_qty':'O','sale_amount':'T','return_date':'U'},'return_header_row':1,'sales_locations':('Safety Sure Shop - Safety Sure Shop',),'return_locations':('Warehouse - Safety Sure Shop',)},
}


def _hash(path):
 h=hashlib.sha256()
 with Path(path).open('rb') as stream:
  for buf in iter(lambda: stream.read(1024*1024),b''): h.update(buf)
 return h.hexdigest()


def _txt(value):
 if value is None: return ''
 # Normalizes Excel integer identifiers exported as floats.
 if isinstance(value,float) and value.is_integer(): return str(int(value))
 return str(value).strip()


def _num(value):
 if value is None or str(value).strip()=='': return Decimal('0')
 if isinstance(value,bool): raise ValueError('Boolean numeric field')
 try: x=Decimal(str(value))
 except (InvalidOperation,ValueError,TypeError): raise ValueError('Invalid monetary or quantity value') from None
 if not x.is_finite(): raise ValueError('Non-finite monetary or quantity value')
 return x.normalize()


def _day(value):
 if value is None or value=='': return ''
 if isinstance(value,datetime): return value.date().isoformat()
 if isinstance(value,date): return value.isoformat()
 if isinstance(value,(int,float)):
  try: return from_excel(value).date().isoformat()
  except (ValueError,OverflowError,TypeError,AttributeError): return 'INVALID_NUMERIC_DATE'
 s=str(value).strip()
 if len(s)>=10:
  try: return date.fromisoformat(s[:10]).isoformat()
  except ValueError: pass
 return 'INVALID_DATE'


def _norm(field,value):
 if field in SALES_MONEY or field in RETURN_MONEY or field in ('sold_qty','return_qty'):
  return _num(value)
 if field in ('date','return_date'): return _day(value)
 return _txt(value)


def _sale_identity(row):
 return tuple(_norm(f,row[f]) for f in ('bill_no','product','sold_qty','sales_amount'))


def _return_identity(row):
 return tuple(_norm(f,row[f]) for f in ('return_code','bill_no','product','return_qty','sale_amount'))


def _sale_signature(row):
 return tuple(_norm(f,row[f]) for f in SALES_FIELDS)


def _return_signature(row):
 return tuple(_norm(f,row[f]) for f in RETURN_FIELDS)


def _read_source(path,kind,month):
 expected=SALES_HEADERS if kind=='sales' else RETURN_HEADERS
 result=[]
 wb=load_workbook(path,read_only=True,data_only=True)
 try:
  if len(wb.worksheets)!=1: raise ValueError(f'{kind}: expected one source worksheet')
  ws=wb.active; it=ws.iter_rows(values_only=True)
  first=tuple(next(it,()))
  if first!=expected: raise ValueError(f'{kind}: unexpected headers or column order')
  for row_index,row in enumerate(it,2):
   if all(v is None for v in row): continue
   if len(row)!=len(expected): raise ValueError(f'{kind}: source row {row_index} has wrong width')
   raw=dict(zip(expected,row))
   if kind=='sales': record={f:raw[h] for f,h in zip(SALES_FIELDS,SALES_HEADERS[:25])}; date_field='date'
   else:
    record={'return_code':raw['Return Code'],'bill_no':raw['Bill No'],'warehouse':raw['Warehouse'],
       'product':raw['Product Description'],'return_qty':raw['Return Qty'],'sale_amount':raw['Sale Amount'],'return_date':raw['Return Date']}
    date_field='return_date'
   d=_day(record[date_field])
   if not d.startswith(month+'-'): raise ValueError(f'{kind}: row {row_index} outside {month} or invalid date')
   result.append(record)
 finally: wb.close()
 return result


def _read_template(path,conf):
 sales=[]; returns=[]
 wb=load_workbook(path,read_only=True,data_only=False)
 try:
  if 'SalesRepItem' not in wb or 'Ord.Rtn' not in wb: raise ValueError('Required sheet missing')
  ws=wb['SalesRepItem']; start=conf['sales_col']
  actual=tuple(ws.cell(4,start+j).value for j in range(25))
  if actual!=SALES_HEADERS[:25]: raise ValueError('SalesRepItem row-4 header mismatch')
  for row in ws.iter_rows(min_row=5,values_only=True):
   vals=row[start-1:start+24]
   if not vals or all(v is None for v in vals): continue
   d=dict(zip(SALES_FIELDS,vals))
   # Observed on ALL five real templates: the row-4 Paid By / Date
   # labels are two columns LEFT of their actual row-5+ values.
   # The intervening two cells contain numeric calculated values.
   # This read-only auditor follows the VERIFIED DATA positions, not headers.
   d['paid_by'] = row[start+24] if len(row)>start+24 else None
   d['date'] = row[start+25] if len(row)>start+25 else None
   if not (_txt(d['bill_no']) and _txt(d['product'])):
    # Sidebar/summary cells are not an identifiable sale transaction.
    continue
   sales.append(d)
  ws=wb['Ord.Rtn']; mapping=conf['returns_col']; header_row=conf['return_header_row']
  expected={'return_code':'Return Code','bill_no':'Bill No','warehouse':'Warehouse','product':'Product Description','return_qty':'Return Qty','sale_amount':'Sale Amount','return_date':'Return Date'}
  from openpyxl.utils.cell import column_index_from_string
  indexes={f:column_index_from_string(c)-1 for f,c in mapping.items()}
  for f,idx in indexes.items():
   if ws.cell(header_row,idx+1).value != expected[f]: raise ValueError('Ord.Rtn header mismatch: '+f)
  for row in ws.iter_rows(min_row=header_row+1,values_only=True):
   d={f:(row[idx] if idx<len(row) else None) for f,idx in indexes.items()}
   if not any(v is not None for v in d.values()): continue
   if not (_txt(d['return_code']) and _txt(d['bill_no']) and _txt(d['product'])): continue
   returns.append(d)
 finally: wb.close()
 return sales,returns


def _compare(source,template,kind):
 """Multiset matches including repeated identical lines; no PII in results."""
 identity=_sale_identity if kind=='sales' else _return_identity
 signature=_sale_signature if kind=='sales' else _return_signature
 fields=SALES_FIELDS if kind=='sales' else RETURN_FIELDS
 sgroup=defaultdict(list); tgroup=defaultdict(list)
 for r in source: sgroup[identity(r)].append(r)
 for r in template: tgroup[identity(r)].append(r)
 matched=0; missing=0; additional=0; field_drift=Counter(); changed_rows=0
 # First pair identical full records, then pair remaining rows and report drift.
 for k in sgroup.keys() | tgroup.keys():
  src=sgroup[k]; tmp=tgroup[k]; matches=min(len(src),len(tmp))
  matched+=matches; missing+=max(len(src)-len(tmp),0); additional+=max(len(tmp)-len(src),0)
  sc=Counter(signature(r) for r in src)
  tc=Counter(signature(r) for r in tmp)
  exact=sc & tc
  for sig,n in exact.items():
   sc[sig]-=n; tc[sig]-=n
  residual_s=[sig for sig,n in sc.items() for _ in range(n)]
  residual_t=[sig for sig,n in tc.items() for _ in range(n)]
  for a,b in zip(residual_s,residual_t):
   changed_rows+=1
   for f,av,bv in zip(fields,a,b):
    if av!=bv: field_drift[f]+=1
 return {'source_records':len(source),'template_records':len(template),'matched_identity':matched,
    'missing_from_template':missing,'additional_in_template':additional,
    'full_field_different_matched_rows':changed_rows,'field_difference_counts':dict(sorted(field_drift.items()))}



def _additional_worksheet_lines(source,template,kind,expected_locations):
    """Identify extra existing template rows as aggregate location counts only."""
    identity=_sale_identity if kind=='sales' else _return_identity
    counter=Counter(identity(r) for r in source)
    unexpected=Counter(); in_scope=0; outside_scope=0
    location_field='location' if kind=='sales' else 'warehouse'
    for record in template:
        key=identity(record)
        if counter[key]:
            counter[key]-=1
            continue
        location=_txt(record[location_field])
        unexpected[location]+=1
        if location in expected_locations: in_scope+=1
        else: outside_scope+=1
    return {'additional_same_business_location':in_scope,
            'additional_other_locations':outside_scope,
            'additional_location_counts':dict(sorted(unexpected.items()))}


def _totals(records,amount,quantity):
 return {'records':len(records), 'quantity':str(sum((_num(r[quantity]) for r in records),Decimal(0))),
   'amount':str(sum((_num(r[amount]) for r in records),Decimal(0)).quantize(Decimal('0.01')))}


def audit(sales_path,returns_path,templates_dir,month='2026-10'):
 if len(month)!=7 or month[4]!='-': raise ValueError('Month must use YYYY-MM')
 try: date.fromisoformat(month+'-01')
 except ValueError: raise ValueError('Invalid month') from None
 inputs=[Path(sales_path).resolve(),Path(returns_path).resolve()]
 templates={b:(Path(templates_dir)/c['file']).resolve() for b,c in BRANDS.items()}
 inputs+=list(templates.values())
 if len(set(inputs))!=len(inputs) or any(not p.is_file() for p in inputs): raise ValueError('Seven distinct existing input files required')
 before={p:_hash(p) for p in inputs}
 sales=_read_source(inputs[0],'sales',month); returns=_read_source(inputs[1],'returns',month)
 sales_by={b:[] for b in BRANDS}; return_by={b:[] for b in BRANDS}; other_sales=[];other_returns=[]
 sales_map={loc:b for b,c in BRANDS.items() for loc in c['sales_locations']}
 return_map={loc:b for b,c in BRANDS.items() for loc in c['return_locations']}
 for record in sales:
  b=sales_map.get(_txt(record['location']))
  (sales_by[b] if b else other_sales).append(record)
 for record in returns:
  b=return_map.get(_txt(record['warehouse']))
  (return_by[b] if b else other_returns).append(record)
 result={'status':'REVIEW_REQUIRED','mode':'REAL_INPUT_READ_ONLY_NO_WRITES','month':month,
   'validation_scope':'25 sales fields and 7 returns fields; matches by tested transaction keys',
   'important_real_template_layout_note':'In SalesRepItem row 4, Paid By and Date header labels are two columns left of actual transaction values in all five current real templates; this audit reads verified data positions without modifying headers.',
   'source_sales':_totals(sales,'sales_amount','sold_qty'),
   'source_returns':_totals(returns,'sale_amount','return_qty'),
   'unmapped_sales':_totals(other_sales,'sales_amount','sold_qty'),
   'unmapped_returns':_totals(other_returns,'sale_amount','return_qty'),
   'businesses':{},'input_sha256':{p.name:before[p] for p in inputs},
   'warnings':[]}
 for b,conf in BRANDS.items():
  current_sales,current_returns=_read_template(templates[b],conf)
  sx=_compare(sales_by[b],current_sales,'sales'); rx=_compare(return_by[b],current_returns,'returns')
  ss=_totals(sales_by[b],'sales_amount','sold_qty'); rs=_totals(return_by[b],'sale_amount','return_qty')
  sx['additional_location_analysis']=_additional_worksheet_lines(sales_by[b],current_sales,'sales',conf['sales_locations'])
  rx['additional_location_analysis']=_additional_worksheet_lines(return_by[b],current_returns,'returns',conf['return_locations'])
  result['businesses'][b]={'sales':sx,'returns':rx,'source_sales_totals':ss,'source_returns_totals':rs,
    'source_net_amount':str(Decimal(ss['amount'])-Decimal(rs['amount']))}
  if any(x['missing_from_template'] for x in (sx,rx)):
   result['warnings'].append(f'{b}: source lines missing from existing workbook (review only; no writes)')
  if any(x['full_field_different_matched_rows'] for x in (sx,rx)):
   result['warnings'].append(f'{b}: matched keys have field differences (manual review required)')
  if any(x['additional_in_template'] for x in (sx,rx)):
   result['warnings'].append(f'{b}: workbook contains additional lines; do not delete automatically')
 after={p:_hash(p) for p in inputs}
 if before!=after: raise RuntimeError('One or more original inputs changed during read-only audit; discard results')
 result['inputs_unchanged_sha256_verified']=True
 result['mapped_sales_records']=sum(len(a) for a in sales_by.values())
 result['mapped_return_records']=sum(len(a) for a in return_by.values())
 result['mapped_sales_amount']=str(sum((Decimal(_totals(a,'sales_amount','sold_qty')['amount']) for a in sales_by.values()),Decimal(0)))
 result['mapped_returns_amount']=str(sum((Decimal(_totals(a,'sale_amount','return_qty')['amount']) for a in return_by.values()),Decimal(0)))
 result['financial_partition_check']=(Decimal(result['mapped_sales_amount'])+Decimal(result['unmapped_sales']['amount'])==Decimal(result['source_sales']['amount'])
  and Decimal(result['mapped_returns_amount'])+Decimal(result['unmapped_returns']['amount'])==Decimal(result['source_returns']['amount']))
 if not result['financial_partition_check']: raise ValueError('Financial partition does not reconcile')
 no_missing=all(not d[k]['missing_from_template'] for d in result['businesses'].values() for k in ('sales','returns'))
 no_drift=all(not d[k]['full_field_different_matched_rows'] for d in result['businesses'].values() for k in ('sales','returns'))
 result['status']='NO_INSERTS_NEEDED_REVIEW_EXTRAS' if no_missing and no_drift else 'REVIEW_REQUIRED'
 result['production_write_approved']=False
 return result


def _export_report(result,directory):
 out=Path(directory).resolve()
 if out.exists(): raise FileExistsError('Output directory exists: refuse to overwrite')
 # Fail-closed: reports go in a new directory, with no private record-level exports.
 out.mkdir(parents=True,exist_ok=False)
 try:
  (out/'real_data_audit.json').write_text(json.dumps(result,indent=2,ensure_ascii=False)+'\n',encoding='utf-8')
  with (out/'business_summary.csv').open('w',encoding='utf-8-sig',newline='') as f:
   w=csv.writer(f);w.writerow(['Business','Sales source','Sales matched','Sales missing','Sales additional','Sales full-field differences','Return source','Return matched','Return missing','Return additional','Return field differences','Sales amount','Return amount','Net amount'])
   for b,d in result['businesses'].items():
    s=d['sales'];r=d['returns'];w.writerow([b,s['source_records'],s['matched_identity'],s['missing_from_template'],s['additional_in_template'],s['full_field_different_matched_rows'],r['source_records'],r['matched_identity'],r['missing_from_template'],r['additional_in_template'],r['full_field_different_matched_rows'],d['source_sales_totals']['amount'],d['source_returns_totals']['amount'],d['source_net_amount']])
  with (out/'field_differences.csv').open('w',encoding='utf-8-sig',newline='') as f:
   w=csv.writer(f);w.writerow(['Business','Type','Field','Different matched records'])
   for b,d in result['businesses'].items():
    for kind in ('sales','returns'):
     for field,count in d[kind]['field_difference_counts'].items(): w.writerow([b,kind,field,count])
  (out/'READ_ME_FIRST.txt').write_text('CONFIDENTIAL - PRIVATE AUDIT\nRead-only; original input SHA-256 hashes verified unchanged.\nNo original files were written.\nNo production workbook update is authorized or performed.\nDo not publish source hashes or audit summaries to a public repository without company permission.\nAdditional template records must not be deleted automatically.\n',encoding='utf-8')
 except BaseException:
  for p in out.iterdir():
   if p.is_file(): p.unlink()
  out.rmdir()
  raise
 return out


def main(argv=None):
 parser=argparse.ArgumentParser(description='READ-ONLY real business exports / five templates reconciliation')
 parser.add_argument('--sales',required=True);parser.add_argument('--returns',required=True)
 parser.add_argument('--templates-dir',required=True);parser.add_argument('--output',required=True)
 parser.add_argument('--month',default='2026-10')
 args=parser.parse_args(argv)
 inputs=[Path(args.sales).resolve(),Path(args.returns).resolve()]
 templates=[(Path(args.templates_dir)/conf['file']).resolve() for conf in BRANDS.values()]
 target=Path(args.output).resolve()
 if target in inputs+templates or any(target.is_relative_to(p) for p in inputs+templates): raise ValueError('Output must not overlap input file')
 result=audit(args.sales,args.returns,args.templates_dir,args.month)
 dest=_export_report(result,args.output)
 print('REAL BUSINESS READ-ONLY AUDIT:',result['status'])
 for b,d in result['businesses'].items():
  s=d['sales'];r=d['returns']
  print(f"{b}: sales {s['matched_identity']}/{s['source_records']} (missing {s['missing_from_template']}, extra {s['additional_in_template']}, drift {s['full_field_different_matched_rows']}); returns {r['matched_identity']}/{r['source_records']} (missing {r['missing_from_template']}, extra {r['additional_in_template']}, drift {r['full_field_different_matched_rows']})")
 print('Unmapped sales:',result['unmapped_sales']['records'],'| returns:',result['unmapped_returns']['records'])
 print('Original inputs SHA-256 unchanged: YES')
 print('PRIVATE audit output:',dest)
 return result

if __name__=='__main__': main()
