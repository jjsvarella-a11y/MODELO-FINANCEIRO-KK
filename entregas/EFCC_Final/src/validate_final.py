from reference_io import load_reference
"""Independent recalculation checks against executed Python reference and invariants."""
from pathlib import Path
from zipfile import ZipFile
from lxml import etree as E
import json,sys,posixpath,math
ROOT=Path(__file__).resolve().parents[1];WORK=Path('/workspace/kk01_efcc');T=lambda n:'{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'+n
R='{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id'
def read(z):
 w=E.fromstring(z.read('xl/workbook.xml'));rr={n.get('Id'):n.get('Target') for n in E.fromstring(z.read('xl/_rels/workbook.xml.rels'))};strings=[''.join(n.itertext()) for n in E.fromstring(z.read('xl/sharedStrings.xml'))] if 'xl/sharedStrings.xml' in z.namelist() else [];out={};errors=[];maxlen=0;fcount=0
 for s in w.find(T('sheets')):
  sn=s.get('name');p=rr[s.get(R)];p=p.lstrip('/') if p.startswith('/') else posixpath.normpath('xl/'+p);r=E.fromstring(z.read(p));cells={}
  for c in r.iter(T('c')):
   v=c.findtext(T('v'));typ=c.get('t');f=c.find(T('f'))
   if f is not None:fcount+=1;maxlen=max(maxlen,len(f.text or ''))
   if typ=='e':errors.append([sn,c.get('r'),v])
   if typ=='s':v=strings[int(v)] if v is not None else None
   elif typ=='inlineStr':v=''.join(c.find(T('is')).itertext())
   elif typ not in ['str','e'] and v is not None:
    try:v=float(v)
    except ValueError:errors.append([sn,c.get('r'),'invalid numeric '+v])
   cells[c.get('r')]=v
  out[sn]=cells
 return out,errors,{'formula_count':fcount,'max_formula_characters':maxlen,'sheets':len(out)}
def letter(i):
 out=''
 while i>=0:out=chr(i%26+65)+out;i=i//26-1
 return out
def validate(source):
 ref=load_reference(ROOT);diff=[];checks=[];n=0;maxdelta=0
 with ZipFile(source) as z:cells,errors,stats=read(z)
 for label,rr in ref['scenarios'].items():
  sn='V06 Motor' if label=='Base' else 'V06 Referencia' if label=='Reference Case' else f'V06 S{list(ref["scenarios"]).index(label):02}' if list(ref['scenarios']).index(label)<25 else f'V06 M{list(ref["scenarios"]).index(label)-24:02}' if list(ref['scenarios']).index(label)<55 else f'V06 O{list(ref["scenarios"]).index(label)-54:02}'
  for m,d in enumerate(rr,5):
   for k,v in d.items():
    if not isinstance(v,(int,float)):continue
    addr=letter(ref['main_columns'][k])+str(m);got=cells[sn].get(addr);n+=1
    if not isinstance(got,(int,float)):diff.append([label,addr,k,v,got]);continue
    delta=abs(got-v);maxdelta=max(maxdelta,delta)
    if delta>max(.00002,abs(v)*1e-9):diff.append([label,addr,k,v,got])
  # Independent accounting identities, restrictions and reserve movements.
  for m,d in enumerate(rr,1):
   def fail(name,val):
    if abs(val)>.00002:checks.append([label,m,name,val])
   fail('cash_identity',d['cash_check']);fail('debt_identity',d['debt_check']);fail('ffe_fund_identity',d['ffe_fund_check']);fail('inventory_94',d['inventory_check'])
   if d['dividend']>.00002:
    if d['debt']>.005 or d['funding_gap_cumulative']>.005 or d['cash']<d['reserve']-.005 or d['dividend']>d['trailing_12m']/ref['EFCC']['effective']['coverage']+.005:checks.append([label,m,'dividend_constraint',d['dividend']])
   if d['debt']<-.005 or d['cumulative_draw']>30000000+.005:checks.append([label,m,'debt_limit',d['debt']])
 # Individual schedules and consolidated cash reconcile at all months.
 for m in range(5,185):
  v=cells['V06 Conciliacao Vendas'].get('D'+str(m))
  if not isinstance(v,(int,float)) or abs(v)>.00002:checks.append(['commercial',m-4,'monthly_reconciliation',v])
 for r in range(5,99):
  v=cells['V06 Disponibilidade'].get('W'+str(r))
  if not isinstance(v,(int,float)) or abs(v)>.00002:checks.append(['commercial',r,'unit_schedule',v])
 result={'status':'PASS' if not (errors or diff or checks) else 'FAIL','platform':'LibreOffice independent recalculation after invalidating all formula caches','excel_native':'NOT TESTED','scenario_count':len(ref['scenarios']),'numeric_monthly_comparisons':n,'maximum_absolute_numeric_difference':maxdelta,'errors':errors,'differences':diff,'invariant_failures':checks,**stats}
 (ROOT/'Validacao_Financeira.json').write_text(json.dumps(result,indent=2,ensure_ascii=False));print({k:len(v) if isinstance(v,list) else v for k,v in result.items()});assert not (errors or diff or checks),(errors[:4],diff[:4],checks[:4]);return result
if __name__=='__main__':validate(Path(sys.argv[1]))
