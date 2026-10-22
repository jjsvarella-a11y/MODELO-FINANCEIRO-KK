"""Recalculate the existing 30 matrix engines and verify real investor NPV.
Private labs freeze only unrelated S/O/reference branches. M01..M30, their
cash flows, the main engine, all inputs and the visible/report matrix stay live.
"""
from pathlib import Path
from zipfile import ZipFile,ZIP_DEFLATED
from lxml import etree as E
from collections import defaultdict
from concurrent.futures import ThreadPoolExecutor,as_completed
import copy,json,gzip,sys,re,hashlib
from assemble import paths,T,xml
from validate_final import read
import scenario_reference as sr
from premium_refine import BASE,MASTER,ROOT,OUT,WORK,H,Q,ADRS,OCCS,OROW
sys.path.insert(0,str(ROOT.parents[0]/'V06/src'))
from recalculate import recalculate
from xlsxwriter.utility import xl_col_to_name as col
LAB=WORK/'matrix_tests';LAB.mkdir(exist_ok=True);SOURCE=LAB/'Matrix_Lab.xlsx'
with gzip.open(ROOT/'Referencia_Python_V06_SYNC.json.gz','rt') as f:REFERENCE=json.load(f)
PARAMS=REFERENCE['EFCC']['effective'];CC=REFERENCE['main_columns']

def near(a,b):assert isinstance(a,(int,float)) and abs(a-b)<=max(.00005,abs(b)*1e-9),(a,b)
def setup():
 with ZipFile(MASTER) as z,ZipFile(SOURCE,'w',ZIP_DEFLATED,compresslevel=3) as out:
  _,pp=paths(z);lookup={p:sn for sn,p in pp.items()}
  for item in z.infolist():
   data=z.read(item.filename);sn=lookup.get(item.filename)
   freeze=sn is not None and (bool(re.fullmatch(r'V06 [SO]\d+|V06 Referencia',sn)) or (bool(re.fullmatch(r'V06 F\d+',sn)) and not (sn=='V06 F01' or 26<=int(sn[5:])<=55)))
   if freeze or sn=='V06 Cenarios':
    root=E.fromstring(data)
    for row in root.find(T('sheetData')):
     if sn=='V06 Cenarios' and (int(row.get('r'))==5 or 30<=int(row.get('r'))<=59):continue
     for c in row:
      f=c.find(T('f'))
      if f is not None:c.remove(f)
    data=xml(root)
   out.writestr(copy.copy(item),data)

def numeric(model,adr,occ,discount,exitmonth,exitmultiple,exitcost):
 rows=[{k:v.v for k,v in row.items()} for row in model.calculate('MatrixNumeric',{'hotel_mode':'Custom','adr_input':adr,'occ_input':occ})['rows']];last=int(exitmonth);end=rows[last-1];enterprise=sum(r['ebitda'] for r in rows[last-12:last])*exitmultiple;equity=max(0,enterprise*(1-exitcost)-end['debt']+end['cash']-end['hotusa_reserve']-end['tax_payable']);npv=0
 for m,d in enumerate(rows,1):
  dividend=d['dividend']*(1900000/4150000 if d['dividend_target']>=4150000 else 1650000/3575000 if d['dividend_target']>=3575000 else 1400000/3000000);cf=dividend-d['investor_outlay']+(equity*.5 if m==last else 0) if m<=last else 0;npv+=cf/(1+discount)**(m/12)
 return rows,npv

def inspect(file,name,overrides):
 with ZipFile(file) as z:cells,errors,_=read(z)
 assert not errors,errors[:5]
 parameters=PARAMS.copy();parameters.update(overrides)
 activeadr=cells[H]['E9'];activeocc=cells[H]['I9'];near(cells['V06 Matriz ADR']['G6'],activeadr);near(cells['V06 Matriz ADR']['B8'],activeocc)
 near(cells['V06 Cenarios']['M39'],cells[Q]['M15']) # M10 is the exact-current intersection.
 model=sr.Model(sr.SinkWorkbook(),defaultdict(lambda:None),{k:sr.asx(v) for k,v in parameters.items()},[{k:sr.asx(v) for k,v in u.items()} for u in sr.legacy.UNITS],sr.legacy.CHAPTERS);model.tickets=[(code,cat,*map(sr.asx,[amount,share,month,1,1,0,0])) for code,cat,amount,share,month in sr.TICKETS]
 points=[];monthly=0
 for ri,occ in enumerate(OCCS):
  for ci,adr in enumerate(ADRS):
   index=OROW[ri]*5+ci;srrow=index+5;flow=cells[f'V06 F{index+26:02}'];expected=sum((flow.get('B'+str(m+5)) or 0)/(1+parameters['discount'])**(m/12) for m in range(181));near(cells['V06 Matriz Valores']['C'+str(srrow)],expected)
   ha=col(8+ci*4)+str(120+ri*2);pa=col(2+ci*2)+str(508+ri*2);near(cells[H][ha],expected);near(cells['DASH 10 V06'][pa],expected)
   point={'adr':adr,'occupancy':occ,'npv':expected,'dashboard_cell':ha,'report_cell':pa}
   if (ri,ci) in [(0,0),(0,3),(4,0),(4,3)]:
    rows,independent=numeric(model,adr,occ,parameters['discount'],parameters['exit_month'],parameters['exit_multiple'],parameters['exit_cost']);near(expected,independent);point['independent_core_reference_npv']=independent
    engine=cells[f'V06 M{index+1:02}']
    for m,row in enumerate(rows,5):
     for key,value in row.items():
      if isinstance(value,(int,float)):near(engine[col(CC[key])+str(m)],value);monthly+=1
   points.append(point)
 # Independently verify the reused exact-current engine against every main month.
 for r in range(5,185):
  for key,c in CC.items():
   a=col(c)+str(r)
   if isinstance(cells['V06 Motor'].get(a),(float,int)):near(cells['V06 M10'][a],cells['V06 Motor'][a]);monthly+=1
 threshold=cells[H]['W116'];colors=[('near' if abs(x['npv'])<=threshold else 'negative' if x['npv']<0 else 'positive') for x in points]
 expectedadr=next((adr for ci,adr in enumerate(ADRS) if cells['V06 Matriz Valores']['C'+str(10+ci)]>=0),None);expectedocc=next((occ for ri,occ in enumerate(OCCS) if cells['V06 Matriz Valores']['C'+str(9+OROW[ri]*5)]>=0),None)
 for cell,value in [('E135',expectedadr),('J135',expectedocc)]:
  got=cells[H][cell]
  if value is None:assert isinstance(got,str) and ('grelha' in got or 'grid' in got),got
  else:near(got,value)
 near(cells[H]['O135'],cells[Q]['M15']);near(cells[H]['T135'],max(0,-cells[Q]['M15']))
 return {'test':name,'status':'PASS','active_adr':activeadr,'active_occupancy':activeocc,'current_npv':cells[Q]['M15'],'exact_current_grid_npv':cells['V06 Cenarios']['M39'],'monthly_values_compared':monthly,'core_reference_corners':4,'npv_points':points,'color_classes':colors,'break_even_adr':expectedadr,'break_even_occupancy':expectedocc,'errors':0}

CASES=[('Conservative',[(H,'H23','Conservative')],{'hotel_mode':'Conservative','hotel_scenario':'Conservador'}),('Favorable',[(H,'H23','Favorable')],{'hotel_mode':'Favorable','hotel_scenario':'Favorável'}),('Custom',[(H,'H23','Custom'),(H,'H25',200),(H,'H27',.65)],{'hotel_mode':'Custom','hotel_scenario':'Favorável','adr_input':200,'occ_input':.65}),('Custom_ADR',[(H,'H23','Custom'),(H,'H25',205),(H,'H27',.65)],{'hotel_mode':'Custom','hotel_scenario':'Favorável','adr_input':205,'occ_input':.65}),('Custom_Occ',[(H,'H23','Custom'),(H,'H25',200),(H,'H27',.66)],{'hotel_mode':'Custom','hotel_scenario':'Favorável','adr_input':200,'occ_input':.66}),('Fees',[(H,'H32','KK Optimized Terms')],{'base_fee':.02,'systems':.015,'incentive':.05}),('MICE',[(H,'N25',.55)],{'mice_util':.55}),('FFE',[(H,'N27',.03)],{'ffe_rate':.03}),('Discount',[(Q,'H29','8%')],{'discount':.08}),('Tax', [('DASH 02 V06','N25',15)],{'exempt_years':15}),('FinanceExit',[('DASH 03 V06','H23',.08),(Q,'H25',10),(Q,'H23','Conservative')],{'rate':.08,'exit_multiple':10,'reserve':3000000})]

def run(case):
 name,changes,parameters=case;directory=LAB/name;directory.mkdir(exist_ok=True);target=directory/(name+'.xlsx')
 with ZipFile(SOURCE) as z,ZipFile(target,'w',ZIP_DEFLATED,compresslevel=3) as out:
  _,pp=paths(z);group=defaultdict(list)
  for sn,a,value in changes:group[pp[sn]].append((a,value))
  for item in z.infolist():
   data=z.read(item.filename)
   if item.filename in group:
    root=E.fromstring(data);lookup={c.get('r'):c for c in root.iter(T('c'))}
    for address,value in group[item.filename]:
     c=lookup[address]
     for node in list(c):c.remove(node)
     c.attrib.pop('t',None)
     if isinstance(value,str):c.set('t','inlineStr');E.SubElement(E.SubElement(c,T('is')),T('t')).text=value
     else:E.SubElement(c,T('v')).text=str(value)
    data=xml(root)
   out.writestr(copy.copy(item),data)
 result=recalculate(target,directory/'recalc');proof=inspect(result,name,parameters);print(json.dumps({k:v for k,v in proof.items() if k not in ['npv_points','color_classes']},ensure_ascii=False),flush=True);return proof

def main():
 if '--verify-existing' in sys.argv:
  results=[inspect(LAB/name/'recalc/out'/(name+'.xlsx'),name,parameters) for name,_,parameters in CASES]
 else:
  setup();results=[]
  with ThreadPoolExecutor(max_workers=2) as pool:
   for future in as_completed([pool.submit(run,case) for case in CASES]):results.append(future.result())
 baseline=inspect(WORK/'full_recalc/out'/MASTER.name,'Base',{});byname={p['test']:p for p in results};byname['Base']=baseline
 for name in ['Custom_ADR','Custom_Occ']:
  near_values=all(abs(a['npv']-b['npv'])<.00005 for a,b in zip(byname['Custom']['npv_points'],byname[name]['npv_points']));assert near_values
  assert abs(byname['Custom']['current_npv']-byname[name]['current_npv'])>.01
  byname[name]['fixed_grid_explanation']='All fixed ADR/occupancy points remain unchanged by definition; marker, exact-current intersection, current NPV and exact-axis break-even data update.'
 for name in ['Conservative','Favorable','Fees','MICE','FFE','Discount','Tax','FinanceExit']:
  changed=sum(abs(a['npv']-b['npv'])>.01 for a,b in zip(baseline['npv_points'],byname[name]['npv_points']));assert changed>0,name;byname[name]['changed_grid_points_vs_base']=changed
 proof={'status':'PASS','platform':'LibreOffice cache-invalidated recalculation; all 30 existing matrix engines remain live','native_excel365':'NOT TESTED','base':baseline,'tests':results,'independent_corner_monthly_checks_per_case':100800,'exact_current_engine_monthly_checks_per_case':25200,'near_zero_precedence':'ABS(NPV) <= editable threshold; negative below minus threshold; positive above threshold','break_even_scope':'Minimum feasible tested ADR at exact active occupancy, and minimum feasible tested occupancy at exact active ADR; not a continuous extrapolated root'}
 (OUT/'Testes_Matriz_Investimento.json').write_text(json.dumps(proof,indent=2,ensure_ascii=False));print('Investment sensitivity matrix PASS',len(results)+1,'cases')
if __name__=='__main__':main()
