"""Recalculate report under PT, EN and an actual interest-rate mutation."""
from pathlib import Path
from zipfile import ZipFile,ZIP_DEFLATED
from lxml import etree as E
import sys,json,concurrent.futures,openpyxl,math,os
from recalculate import recalculate
from extend import paths,DASHES,T
from build_report import SECTIONS
ROOT=Path(__file__).resolve().parents[1];WORK=Path(os.environ.get('KK01_REPORT_WORKDIR','/workspace/kk01_dash10'));manifest=json.loads((ROOT/'Fontes_DASH10.json').read_text())
def mutate(value,address,name):
 source=WORK/'Lab_DASH10.xlsx';dest=WORK/(name+'.xlsx')
 with ZipFile(source) as z,ZipFile(dest,'w',ZIP_DEFLATED) as zz:
  _,lookup=paths(z)
  for p in z.infolist():
   data=z.read(p.filename)
   if p.filename==lookup['V06 Controlos']:
    r=E.fromstring(data);c=r.find('.//'+T('c')+'[@r="'+address+'"]');assert c is not None
    for child in list(c):c.remove(child)
    if isinstance(value,str):c.set('t','inlineStr');E.SubElement(E.SubElement(c,T('is')),T('t')).text=value
    else:c.attrib.pop('t',None);E.SubElement(c,T('v')).text=str(value)
    data=E.tostring(r,encoding='UTF-8',xml_declaration=True)
   zz.writestr(p,data)
 return recalculate(dest,WORK/('recalc_'+name))
def check(path,language,base=None):
 w=openpyxl.load_workbook(path,read_only=True,data_only=True);wf=openpyxl.load_workbook(path,read_only=True,data_only=False);s=w['DASH 10 V06'];f=wf['DASH 10 V06'];fail=[];count=0;cache={}
 def cell(sn,addr):
  if sn not in cache:cache[sn]={c.coordinate:c.value for row in w[sn] for c in row if hasattr(c,'coordinate')}
  return cache[sn].get(addr)
 import re
 for binding in manifest['financial_bindings']:
  formula=binding['formula'];direct=re.fullmatch(r"'([^']+)'!\$?([A-Z]+)\$?(\d+)",formula);sum_=re.fullmatch(r"SUM\('([^']+)'!([A-Z]+)(\d+):([A-Z]+)(\d+)\)",formula)
  if direct:expected=cell(direct[1],direct[2]+direct[3])
  elif sum_:expected=sum(cell(sum_[1],sum_[2]+str(i)) or 0 for i in range(int(sum_[3]),int(sum_[5])+1))
  else:raise AssertionError(formula)
  actual=s[binding['cell']].value;count+=1
  if not isinstance(actual,(int,float)) or not isinstance(expected,(int,float)) or abs(actual-expected)>max(.01,abs(expected)*1e-9):fail.append([binding['cell'],actual,expected])
  count+=1
  if f[binding['cell']].data_type!='f':fail.append(['manual financial cell',binding['cell']])
 for i,(pt,en) in enumerate(SECTIONS,1):
  expected=f'{i:02} — '+(pt if language=='PT' else en);actual=s.cell((i-1)*50+8,1).value;count+=1
  if expected!=actual:fail.append(['section',i,actual,expected])
 for sn in DASHES:
  count+=1
  expected='10 — Relatório de Viabilidade' if language=='PT' else '10 — Investment Report'
  if w[sn]['B43'].value!=expected:fail.append(['navigation',sn,w[sn]['B43'].value])
 errors=[]
 for row in s:
  for c in row:
   if c.data_type=='e':errors.append([c.coordinate,c.value])
 if errors:fail.append(['report errors',errors])
 if base:
  for a,b in zip(w['V06 Motor'].iter_rows(min_row=5,max_row=184,values_only=True),base):
   for x,y in zip(a,b):
    count+=1
    if isinstance(x,(int,float)) and isinstance(y,(int,float)):
     if abs(x-y)>max(.01,abs(y)*1e-9):fail.append(['language financial change',x,y])
    elif x!=y:fail.append(['language financial type change',x,y])
 out={'file':str(path),'language':language,'comparisons':count,'passed':not fail,'failures':fail[:20],'report_npv':s['A26'].value};w.close();wf.close();return out
if __name__=='__main__':
 pt=WORK/'recalc_PT/out/Lab_DASH10.xlsx';results=[check(pt,'PT')]
 w=openpyxl.load_workbook(pt,read_only=True,data_only=True);baseline=list(w['V06 Motor'].iter_rows(min_row=5,max_row=184,values_only=True));w.close()
 with concurrent.futures.ThreadPoolExecutor(max_workers=2) as pool:
  en,rate=[x.result() for x in [pool.submit(mutate,'EN','B5','EN'),pool.submit(mutate,.08,'B41','Juro8')]]
 results.extend([check(en,'EN',baseline),check(rate,'PT')])
 # Rate mutation must affect both source and report; compare explicit investor NPV cells.
 src=openpyxl.load_workbook(pt,read_only=True,data_only=True);changed=openpyxl.load_workbook(rate,read_only=True,data_only=True);before=src['DASH 10 V06']['A26'].value;after=changed['DASH 10 V06']['A26'].value
 results.append({'test':'Interest change propagates to report investor NPV','passed':isinstance(before,(int,float)) and isinstance(after,(int,float)) and abs(after-before)>.01,'before':before,'after':after});src.close();changed.close()
 (ROOT/'Testes_DASH10.json').write_text(json.dumps(results,ensure_ascii=False,indent=2));print(json.dumps(results,ensure_ascii=False,indent=2));raise SystemExit(not all(r['passed'] for r in results))
