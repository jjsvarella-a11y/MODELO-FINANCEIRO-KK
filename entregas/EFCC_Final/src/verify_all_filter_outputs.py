"""Verify lower tables and every main chart projection after each dashboard mutation."""
import json,re
from zipfile import ZipFile
from xlsxwriter.utility import xl_col_to_name as col
from concurrent.futures import ThreadPoolExecutor
import test_all_dashboard_filters as lab
ROOT,WORK=lab.ROOT,lab.WORK
old=lab.old
RECORDS={0:['capex','equity','debt_service'],1:['construction','supervision','contingency_paid','extra_etc'],2:['tax_paid','vat_refund','construction_insurance','operating_insurance'],3:['draw','interest','principal','sweep'],4:['receipts','deed_sales','sales_cost','sales_profit'],5:['res_revenue','res_opex','res_fees','owner_distribution','owner_shortfall','res_ffe'],7:['equity','investor_outlay','dividend'],8:['dividend','interest','tax_paid']}
KEYS=['cash','reserve','capex','receipts','vat_refund','tax_paid','debt','deed_sales','res_revenue','owner_distribution','owner_target','equity','investor_outlay','dividend','room_revenue','fb_revenue','mice_revenue','other_revenue','ebitda','revenue','ebitda_margin']
def inspect(test,cells=None):
 file=WORK/'mutations'/test['test']/'recalc/out'/(test['test']+'.xlsx')
 if cells is None:
  with ZipFile(file) as z:cells,errors,_=old.read(z)
  assert not errors
 cc=old.REF['main_columns'];m=cells['V06 Motor'];n=0
 def get(key,r):return m[col(cc[key])+str(r)]
 def check(sn,addr,value):
  nonlocal n
  old.near(cells[sn][addr],value);n+=1
 for i,keys in RECORDS.items():
  sn=lab.DASHES[i][0]
  for j,key in enumerate(keys):check(sn,'M'+str(66+j*2),sum(get(key,r) for r in range(5,185)))
 for year in range(1,16):
  first,last=(year-1)*12+5,year*12+4
  for j,key in enumerate(KEYS,1):
   if key=='ebitda_margin':
    revenue=sum(get('revenue',r) for r in range(first,last+1));value=sum(get('ebitda',r) for r in range(first,last+1))/revenue if revenue>0 else 0
   else:value=get(key,last) if key in ['cash','reserve','debt'] else sum(get(key,r) for r in range(first,last+1))
   check('V06 Visual',col(j)+str(year+4),value)
 selected=[r for r in range(5,185) if get('op_age',r)==cells[lab.H]['N23']-1 and get('available_room_nights',r)>0]
 for addr,key in [('K44','room_revenue'),('N44','fb_revenue'),('E50','mice_revenue'),('H50','revenue'),('N50','ebitda'),('M72','gop'),('M74','operator_fees'),('M76','ffe')]:check(lab.H,addr,sum(get(key,r) for r in selected))
 # Additional visible R2 presentation cells and native chart projections.
 design_path=ROOT/'Design_Spec.json'
 if design_path.exists():
  design=json.loads(design_path.read_text())
  for item in design['display_bindings']:
   got=cells[item['sheet']][item['cell']];expected=cells[item['source_sheet']][item['source_cell']]
   if old.number(expected):check(item['sheet'],item['cell'],expected)
   else:assert got==expected,(item,got,expected);n+=1
  for j,key in enumerate(design.get('chart_keys',[]),240):
   for year in range(1,16):
    first,last=(year-1)*12+5,year*12+4
    if key=='ebitda_margin':
     revenue=sum(get('revenue',r) for r in range(first,last+1));value=sum(get('ebitda',r) for r in range(first,last+1))/revenue if revenue>0 else 0
    else:value=get(key,last) if key in ['cash','debt','reserve','units_not_deeded','res_fund'] else sum(get(key,r) for r in range(first,last+1))
    check('V06 Visual',col(j)+str(year+4),value)
 if design_path.exists():
  for index,key in enumerate(['room_revenue','fb_revenue','mice_revenue','other_revenue']):check('V06 Visual',col(341)+str(201+index),sum(get(key,r) for r in selected))
 # Active/Reference is an explicit report display switch, independent from the active engine.
 if cells[lab.C]['N11']=='Active Scenario':check('DASH 10 V06','G17',cells[lab.Q]['M15'])
 return {'test':test['test'],'status':'PASS','bindings_checked':n,'sheet':test['sheet'],'cell':test['cell'],'choice':test['choice']}
def main():
 source=json.loads((ROOT/'Testes_Todos_Filtros.json').read_text());assert source['status']=='PASS'
 if all('correspondences' in test for test in source['tests']):results=[test['correspondences'] for test in source['tests']]
 else:
  with ThreadPoolExecutor(max_workers=3) as executor:results=list(executor.map(inspect,source['tests']))
 proof={'status':'PASS','scope':'Lower financial tables on all nine dashboards, Hotel selected-year tables, 315 main chart projection values per case and active report NPV','tests':results,'total_bindings_checked':sum(r['bindings_checked'] for r in results),'native_excel365':'NOT TESTED'}
 (ROOT/'Testes_Correspondencias_Dashboards.json').write_text(json.dumps(proof,indent=2,ensure_ascii=False));print('Cross-dashboard lower tables and chart projections',len(results),proof['total_bindings_checked'],'PASS')
if __name__=='__main__':main()
