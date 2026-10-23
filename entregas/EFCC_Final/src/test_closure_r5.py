"""R5 acceptance: true input mutations, external recalculation, financial invariants.
Private labs freeze unrelated sensitivity branches only; delivered workbook stays live.
"""
from zipfile import ZipFile,ZIP_DEFLATED
from lxml import etree as E
from pathlib import Path
from concurrent.futures import ThreadPoolExecutor,as_completed
import json,re,copy,sys,math,hashlib
from assemble import paths,T,xml,inject_caches
from validate_final import read
from closure_r5 import BASE,MASTER,OUT,WORK,CC,INPUTS,REF,ACT,IMP,CON,CASH,INV,SALES,col
sys.path.insert(0,str(Path(__file__).resolve().parents[2]/'V06/src'))
from recalculate import recalculate

def load(file):
 with ZipFile(file) as z:return read(z)
def near(a,b,tol=.00005):assert isinstance(a,(float,int)) and isinstance(b,(float,int)) and abs(a-b)<=max(tol,abs(b)*1e-9),(a,b)
def baseline(calculated=None):
 calculated=calculated or WORK/'full_recalc2/out'/MASTER.name;c,errors,stats=load(calculated);assert not errors,errors[:10]
 old,errors,_=load(BASE);assert not errors
 comparisons=0;maxdiff=0
 names=[n for n in old if n=='V06 Motor' or re.fullmatch(r'V06 [SMO]\d+',n) or n=='V06 Referencia']
 for sn in names:
  for rr in range(5,185):
   for k in range(140):
    addr=col(k)+str(rr);x,y=old[sn].get(addr),c[sn].get(addr)
    if isinstance(x,(float,int)) and isinstance(y,(float,int)):near(x,y);maxdiff=max(maxdiff,abs(x-y));comparisons+=1
 near(c['V06 Cenarios']['M5'],c[INV]['C30']);near(c[INV]['B30'],c[INV]['C30']);near(c[CON]['B7'],63168000)
 tmp=WORK/'Cached.xlsx';inject_caches(MASTER,calculated,tmp);tmp.replace(MASTER)
 proof={'status':'PASS','engine_sheets':len(names),'monthly_numeric_comparisons':comparisons,'maximum_difference':maxdiff,'npv':c['V06 Cenarios']['M5'],'irr':c['V06 Cenarios']['K5'],'new_investor_selected_npv':c[INV]['C30'],'native_excel365':'NOT TESTED'}
 (OUT/'Baseline_Reconciliation.json').write_text(json.dumps(proof,indent=2));print('Baseline',proof,flush=True);return c

def make_lab(matrix=False,suffix=""):
 target=WORK/(('Matrix_Lab' if matrix else 'Active_Lab')+suffix+'.xlsx')
 with ZipFile(MASTER) as z,ZipFile(target,'w',ZIP_DEFLATED,compresslevel=3) as out:
  _,pp=paths(z);lookup={p:n for n,p in pp.items()}
  for item in z.infolist():
   data=z.read(item.filename);sn=lookup.get(item.filename,'');freeze=sn=='V06 Referencia' or bool(re.fullmatch(r'V06 [SO]\d+',sn)) or (bool(re.fullmatch(r'V06 M\d+',sn)) and not matrix) or (bool(re.fullmatch(r'V06 F\d+',sn)) and sn!='V06 F01' and not(matrix and int(sn[5:])>=26))
   if freeze or sn=='V06 Cenarios':
    r=E.fromstring(data)
    for row in r.find(T('sheetData')):
     if sn=='V06 Cenarios' and (int(row.get('r'))==5 or matrix and 30<=int(row.get('r'))<=59):continue
     for c in row:
      f=c.find(T('f'))
      if f is not None:c.remove(f)
    data=xml(r)
   out.writestr(copy.copy(item),data)
 return target

def mutate(source,target,changes):
 with ZipFile(source) as z,ZipFile(target,'w',ZIP_DEFLATED,compresslevel=3) as out:
  _,pp=paths(z);roots={}
  for sn,addr,value in changes:
   if sn not in roots:roots[sn]=E.fromstring(z.read(pp[sn]))
   c=roots[sn].find('.//'+T('c')+'[@r="'+addr+'"]');assert c is not None,(sn,addr)
   for n in list(c):c.remove(n)
   c.attrib.pop('t',None)
   if isinstance(value,str):c.set('t','inlineStr');E.SubElement(E.SubElement(c,T('is')),T('t')).text=value
   else:E.SubElement(c,T('v')).text=str(value)
  for item in z.infolist():out.writestr(copy.copy(item),xml(roots[next(sn for sn in roots if pp[sn]==item.filename)]) if item.filename in [pp[sn] for sn in roots] else z.read(item.filename))

def inspect(c,base,changes,name,matrix=False):
 m=c['V06 Motor'];vc=c['V06 Cenarios'];effective=c['V06 Registo Ativo'];flow=c['V06 F01'];rate=effective['B65'];npv=sum(flow.get('B'+str(k+5),0)/(1+rate)**(k/12) for k in range(181));near(npv,vc['M5']);near(c[IMP]['C9'],npv)
 for j,cell in enumerate(['C5','C6','C7','C8','C9','C10']):
  for sn,_,_ in __import__('premium_ui').DASHES:near(c[sn][('V'+str(26+j))],c[IMP][cell])
 for addr in ['B5','B6','B7','B8','B9','B10','B11','B12','B13']:near(c[IMP][addr],base[IMP][addr])
 for rr in range(5,185):
  near(m[col(CC['cash_check'])+str(rr)],0);near(m[col(CC['debt_check'])+str(rr)],0);near(m[col(CC['ffe_fund_check'])+str(rr)],0);near(m[col(CC['inventory_check'])+str(rr)],0)
 capex=sum(m[col(CC['capex'])+str(r)] for r in range(5,185));near(capex,c[CON]['B7']);near(capex,c[IMP]['C5']);near(c[INV]['C30'],vc['M5']) if c['DASH 06 V06']['H25']=='All' and c['DASH 06 V06']['N25']=='All' and c['DASH 06 V06']['H27']=='New' else None
 if name.startswith('Discount'):
  near(vc['K5'],base['V06 Cenarios']['K5']);assert abs(vc['M5']-base['V06 Cenarios']['M5'])>1000
 if name=='Chapter_10':assert capex>base[IMP]['C5']
 if name.startswith('Tax_'):
  years=int(c['DASH 02 V06']['N25']);opening=int(m['D5'])
  for month in range(1,opening):near(m[col(CC['effective_tax_rate'])+str(month+4)],effective['B9'])
  for month in range(opening,min(181,opening+years*12)):near(m[col(CC['effective_tax_rate'])+str(month+4)],0)
 if name.startswith('Collections_'):
  lag=int(c['DASH 04A V06']['N23'])
  for month in range(1,181):near(m[col(CC['receipts'])+str(month+4)],0 if month<=lag else base['V06 Motor'][col(CC['receipts'])+str(month-lag+4)])
 if name=='Promoter':
  for addr in ['C26','C29','C30','C31','C32','C33','C36']:assert c[INV][addr]=='N/A',(addr,c[INV][addr])
 if name=='Gold':near(c[INV]['C27'],.2);near(c[INV]['C26'],10000000);near(c[INV]['C30'],c['V06 Blocos']['K13']+c['V06 Blocos']['K14'])
 if name=='Block_P01':near(c[INV]['C30'],c['V06 Blocos']['K12'])
 if name.startswith('View_'):near(vc['M5'],base['V06 Cenarios']['M5']);assert c[SALES]['H5'] in [0,1]
 if c['EXEC V06']['H23']=='Reference':
  for rr in range(5,185):
   for k in range(140):near(m[col(k)+str(rr)],base['V06 Motor'][col(k)+str(rr)])
 points=[]
 if matrix:
  for ri,idx in enumerate([0,2,3,4,5]):
   for ci in range(4):
    rr=5+idx*5+ci;value=c['V06 Matriz Valores']['C'+str(rr)];cf=c['V06 F'+str(26+idx*5+ci).zfill(2)];expected=sum(cf.get('B'+str(k+5),0)/(1+rate)**(k/12) for k in range(181));near(value,expected);near(c['DASH 05 V06'][col(8+ci*4)+str(120+ri*2)],value);near(c['DASH 10 V06'][col(2+ci*2)+str(508+ri*2)],value);points.append(value)
  near(c['V06 Cenarios']['M39'],vc['M5'])
 return {'test':name,'status':'PASS','changes':changes,'npv':vc['M5'],'irr':vc['K5'],'moic':vc['L5'],'capex':capex,'peak_debt':c[IMP]['C6'],'liquidity':c[IMP]['C7'],'rate':rate,'investor_npv':c[INV]['C30'],'matrix_points_verified':len(points),'errors':0,'base_reference_unchanged':True}

def run_case(case,base,lab):
 name,changes,matrix=case;folder=WORK/'tests'/name;folder.mkdir(parents=True,exist_ok=True);target=folder/(name+'.xlsx')
 try:
  mutate(lab[matrix],target,changes);file=recalculate(target,folder/'recalc');c,err,stats=load(file);assert not err,err[:10];result=inspect(c,base,changes,name,matrix)
 except Exception as ex:result={'test':name,'status':'FAIL','error':repr(ex),'changes':changes}
 print(result,flush=True);return result

def main():
 base=baseline();labs={False:make_lab(False),True:make_lab(True)}
 H='DASH 05 V06';Q='DASH 07 V06';C='DASH 01 V06';TAX='DASH 02 V06';S='DASH 04A V06';R='DASH 04B V06'
 cases=[('Delay_3',[(C,'H23',3)],False),('Delay_6',[(C,'H23',6)],False),('Overrun_10',[(C,'N23',.1)],False),('ETC_2M',[(C,'H25',2000000)],False),('Chapter_10',[(C,'N194',.1)],False),('Chapter_ETC',[(C,'T195',250000)],False),('VAT_short',[(TAX,'H25',1)],False),('VAT_long',[(TAX,'H25',18)],False),('VAT_partial',[(TAX,'H27',.6)],False),('Tax_3',[(TAX,'N25',3)],True),('Tax_5',[(TAX,'N25',5)],False),('Insurance',[(TAX,'N27',1.25)],False),('Pricing',[(S,'H25','Custom'),(S,'N25',1.1)],False),('Sales_delay',[(S,'H23',6)],False),('Collections_3',[(S,'N23',3)],False),('Collections_6',[(S,'N23',6)],False),('View_Type',[(S,'H27','T1')],False),('View_Area',[(S,'N29',100)],False),('Res_Profile',[(R,'H33','Conservative')],False),('Res_Maintenance',[(R,'H31',.01)],False),('ADR',[(H,'H23','Custom'),(H,'H25',205)],False),('Occupancy',[(H,'H23','Custom'),(H,'H27',.68)],False),('MICE',[(H,'N25',.55)],True),('Fees',[(H,'H32','KK Optimized Terms')],True),('FFE',[(H,'N27',.03)],True),('Conservative',[(H,'H23','Conservative')],False),('Favorable',[(H,'H23','Favorable')],False),('Reserve',[(Q,'H23','Conservative')],False),('Discount_8',[(Q,'H29','8%')],True),('Discount_10',[(Q,'H29','10%')],False),('Discount_Custom',[(Q,'H29','Custom'),(Q,'N25',.09)],False),('Exit_multiple',[(Q,'H25',10)],False),('Exit_120',[(Q,'N29',120)],False),('Dividend',[(Q,'N23','Suspended')],False),('Gold',[('DASH 06 V06','H25','Gold')],False),('Block_P01',[('DASH 06 V06','N25','P01')],False),('Promoter',[('DASH 06 V06','H27','Promoters')],False),('EN',[('DASH 10 Centro','H7','EN')],False),('Reference_restore',[(C,'H23',6),(C,'N194',.1),(S,'N23',6),(H,'H23','Custom'),(H,'H25',200),(Q,'H29','8%'),('EXEC V06','H23','Reference')],False)]
 results=[]
 with ThreadPoolExecutor(max_workers=3) as ex:
  for future in as_completed([ex.submit(run_case,case,base,labs) for case in cases]):results.append(future.result())
 proof={'status':'PASS' if all(x['status']=='PASS' for x in results) else 'FAIL','tests':results,'platform':'LibreOffice, cache-invalidated private labs','native_excel365':'NOT TESTED'};(OUT/'Testes_R5.json').write_text(json.dumps(proof,indent=2,ensure_ascii=False));assert proof['status']=='PASS'
if __name__=='__main__':main()
