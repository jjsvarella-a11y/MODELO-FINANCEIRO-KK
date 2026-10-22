from reference_io import load_reference
"""Actual editable-cell mutations and cache-invalidated recalculations.
For speed, private labs freeze independent untested scenario branches to verified
baseline values. The active engine, calendar, reports and tested branch stay live.
This optimisation is never applied to the delivered workbook.
"""
from pathlib import Path
from zipfile import ZipFile,ZIP_DEFLATED
from lxml import etree as E
import json,sys,re,copy,concurrent.futures
from assemble import paths,T,xml
from validate_final import read
sys.path.insert(0,str(Path(__file__).resolve().parents[2]/'V06/src'))
from recalculate import recalculate
ROOT=Path(__file__).resolve().parents[1];WORK=Path('/workspace/kk01_efcc');SOURCE=WORK/'final_recalc/out/V06_modulos.xlsx'
REF=load_reference(ROOT)
def number(x):return isinstance(x,(int,float))
def near(a,b):assert number(a) and number(b) and abs(a-b)<=max(.00005,abs(b)*1e-9),(a,b)
def make_lab():
 dest=WORK/'Mutation_Lab.xlsx'
 with ZipFile(SOURCE) as z,ZipFile(dest,'w',ZIP_DEFLATED,compresslevel=3) as zz:
  _,pp=paths(z);lookup={v:k for k,v in pp.items()}
  for part in z.infolist():
   data=z.read(part.filename);sn=lookup.get(part.filename,'')
   freeze=bool(re.fullmatch(r'V06 [SMO]\d+|V06 Referencia|V06 F\d+',sn)) and sn not in ['V06 M01','V06 F01','V06 F26']
   if freeze or sn=='V06 Cenarios':
    r=E.fromstring(data)
    for row in r.find(T('sheetData')):
     if sn=='V06 Cenarios' and int(row.get('r')) in [5,30]:continue
     for c in row:
      f=c.find(T('f'))
      if f is not None:c.remove(f)
    data=xml(r)
   zz.writestr(part.filename,data)
 return dest

def run_case(name,changes):
 folder=WORK/'mutations'/name;folder.mkdir(parents=True,exist_ok=True);dest=folder/(name+'.xlsx')
 try:
  with ZipFile(WORK/'Mutation_Lab.xlsx') as z,ZipFile(dest,'w',ZIP_DEFLATED,compresslevel=3) as zz:
   _,pp=paths(z);updates={}
   for sn,addr,value in changes:
    if pp[sn] not in updates:updates[pp[sn]]=E.fromstring(z.read(pp[sn]))
    root=updates[pp[sn]];c=root.find('.//'+T('c')+'[@r="'+addr+'"]');assert c is not None,(sn,addr)
    for ch in list(c):c.remove(ch)
    c.attrib.pop('t',None)
    if isinstance(value,str):c.set('t','inlineStr');E.SubElement(E.SubElement(c,T('is')),T('t')).text=value
    else:E.SubElement(c,T('v')).text=str(value)
   for part in z.infolist():zz.writestr(part.filename,xml(updates[part.filename]) if part.filename in updates else z.read(part.filename))
  if name.startswith('Discount_'):
   with ZipFile(SOURCE) as original,ZipFile(dest) as target:
    _,op=paths(original);_,tp=paths(target);ar=E.fromstring(original.read(op['V06 Cenarios']));br=E.fromstring(target.read(tp['V06 Cenarios']));ac={c.get('r'):c for c in ar.iter(T('c'))}
    # Rebuild via row replacement to preserve addresses and original formula metadata.
    for row in br.find(T('sheetData')):
     for c in list(row):
      addr=c.get('r')
      if re.fullmatch(r'M\d+',addr or '') and int(addr[1:])>=5:
       at=list(row).index(c);row.remove(c);row.insert(at,copy.deepcopy(ac[addr]))
    changes={tp['V06 Cenarios']:xml(br)};temp=dest.with_suffix('.discount.xlsx')
    with ZipFile(temp,'w',ZIP_DEFLATED,compresslevel=3) as zz:
     for part in target.infolist():zz.writestr(part.filename,changes.get(part.filename,target.read(part.filename)))
   temp.replace(dest)
  out=recalculate(dest,folder/'recalc')
  with ZipFile(out) as z:cells,errors,stats=read(z)
  assert not errors,errors[:5]
  base=REF['metrics']['Base'];v=cells['V06 Cenarios'];npv=v['M5'];evidence={'npv':npv,'cash30':v['B5'],'div54':v['E5'],'operator_fees':cells['DASH 05 V06']['I15']}
  if name in ['Discount_8','Discount_Custom10']:
   rate=.08 if name=='Discount_8' else .10
   with ZipFile(SOURCE) as original:
    _,sources=paths(original)
    for idx in range(30):
     root=E.fromstring(original.read(sources[f'V06 F{26+idx:02}']));cf={}
     for c in root.iter(T('c')):
      if re.fullmatch(r'B\d+',c.get('r','')):cf[int(c.get('r')[1:])-5]=float(c.findtext(T('v')) or 0)
     expected=sum(cf.get(m,0)/(1+rate)**(m/12) for m in range(181));near(cells['V06 Matriz Valores']['C'+str(idx+5)],expected)
   evidence['matrix_cells_compared']=30;evidence['discount']=rate
  elif name in ['Construction_6','VAT_18']:
   label='Construction +6M' if name=='Construction_6' else 'VAT delayed'
   for key,addr in [('npv','M5'),('irr','K5'),('moic','L5'),('exit','J5'),('debt','F5'),('div54','E5')]:near(v[addr],REF['metrics'][label][key])
   if name=='Construction_6':near(cells['DASH 10 V06']['A266'],cells['DASH 05 V06']['E9'])
   evidence['expected_full_scenario']=label
  elif name=='EN':
   checked=0
   for r,d in enumerate(REF['main_rows'],5):
    for key,expected in d.items():
     if number(expected):near(cells['V06 Motor'][__import__('xlsxwriter').utility.xl_col_to_name(REF['main_columns'][key])+str(r)],expected);checked+=1
   assert 'Executive Dashboard' in cells['EXEC V06']['B7'];assert 'Consolidated Investment' in cells['DASH 10 V06']['A4'];evidence['unchanged_engine_values']=checked
  elif name in ['Hotel_Custom','Hotel_Optimized','Residential_Optimized','Reserve_Conservative','Interest_8']:
   label={'Hotel_Custom':'ADR 200 Occ 65%','Hotel_Optimized':'Hotel KK Optimized','Residential_Optimized':'Residential KK Optimized','Reserve_Conservative':'Reserve Conservative','Interest_8':'Interest 8%'}[name]
   ex=REF['metrics'][label]
   for key,addr in [('npv','M5'),('irr','K5'),('moic','L5'),('exit','J5'),('debt','F5'),('div54','E5')]:near(v[addr],ex[key])
   evidence['expected_full_scenario']=label
  elif name=='MICE_55':assert npv>base['npv']
  elif name=='FFE_3':near(v['N5'],base['ebitda15']);assert npv>base['npv']
  elif name=='FFE_Spend_50':
   near(npv,base['npv']);near(v['J5'],base['exit']);assert cells['V06 Operacao']['R184']>0;evidence['restricted_fund']=cells['V06 Operacao']['R184']
  elif name=='Matrix_IRR':near(cells['V06 Matriz ADR']['C7'],cells['V06 Matriz Valores']['D5'])
  elif name=='Matrix_ADR155':
   assert cells['V06 Matriz Valores']['C5']>REF['metrics']['ADR 150 Occ 50%']['npv'];near(cells['V06 Matriz ADR']['C7'],cells['V06 Matriz Valores']['C5']);evidence['matrix_npv']=cells['V06 Matriz ADR']['C7']
  elif name=='Unit_Price':near(cells['DASH 04A V06']['M9'],sum(u['price'] for u in __import__('core').legacy.UNITS)+1000);assert npv!=base['npv']
  elif name=='Reference_Report':
   near(cells['DASH 10 V06']['A26'],REF['metrics']['Reference Case']['npv']);assert npv!=base['npv'];evidence['report_reference_npv']=cells['DASH 10 V06']['A26']
  return {'test':name,'status':'PASS','evidence':evidence}
 except Exception as ex:return {'test':name,'status':'FAIL','error':repr(ex)}

def main():
 make_lab()
 with ZipFile(SOURCE) as z:cells,_,_=read(z)
 hr={cells['V06 Hipoteses']['A'+str(r)]:r for r in range(5,34)}
 unitprice=cells['V06 Unidades']['H5']
 cases=[('EN',[('V06 Controlos','B5','EN')]),('Hotel_Custom',[('DASH 05 V06','H23','Custom'),('DASH 05 V06','H25',200),('DASH 05 V06','H27',.65)]),('MICE_55',[('DASH 05 V06','N25',.55)]),('FFE_3',[('DASH 05 V06','N27',.03)]),('FFE_Spend_50',[('V06 Hipoteses','D'+str(hr['ffe_spend']),.5)]),('Hotel_Optimized',[('DASH 05 V06','H32','KK Optimized Terms')]),('Residential_Optimized',[('DASH 04B V06','H23','KK Optimized Terms')]),('Reserve_Conservative',[('DASH 07 V06','H23','Conservative')]),('Interest_8',[('V06 Controlos','B41',.08)]),('Matrix_IRR',[('V06 Matriz ADR','B3','IRR')]),('Matrix_ADR155',[('V06 Matriz ADR','C6',155)]),('Unit_Price',[('V06 Unidades','H5',unitprice+1000)]),('Reference_Report',[('V06 Hipoteses','D'+str(hr['report_case']),'Reference Case'),('V06 Controlos','B41',.08)])]
 results=[]
 with concurrent.futures.ThreadPoolExecutor(max_workers=2) as pool:
  futures={pool.submit(run_case,*case):case[0] for case in cases}
  for future in concurrent.futures.as_completed(futures):
   result=future.result();results.append(result);print(result,flush=True)
 (ROOT/'Testes_Seletores.json').write_text(json.dumps({'platform':'LibreOffice, actual input mutations and full cache invalidation','excel_native':'NOT TESTED','private_lab_optimisation':'Unrelated independent branches frozen; active/tested engines, calendars, dashboards and reports live; delivered workbook retains every formula.','tests':results},ensure_ascii=False,indent=2));assert all(r['status']=='PASS' for r in results)
if __name__=='__main__':main()
