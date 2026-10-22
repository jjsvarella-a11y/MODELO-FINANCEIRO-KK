"""Scenario-to-engine-to-report/graph verification, with before/after and save/reopen."""
import sys,json,csv,re,shutil,hashlib
from pathlib import Path
from zipfile import ZipFile
from concurrent.futures import ThreadPoolExecutor,as_completed
from lxml import etree as E
from xlsxwriter.utility import xl_col_to_name as col,xl_cell_to_rowcol
import test_premium_controls as p
import test_all_dashboard_filters as lab
from scenario_reference import reference
from premium_ui import ROOT,WORK as BASEWORK
old=p.old;H='DASH 05 V06';Q='DASH 07 V06';REPORT='DASH 10 V06';MODES=['Base Case','Conservative','Favorable','Upside Stress','Custom']

def setup(stage,prepare=True):
 folder=BASEWORK/('scenario_'+stage);folder.mkdir(exist_ok=True);p.WORK=folder;lab.WORK=folder;old.WORK=folder
 old.SOURCE=BASEWORK/'checkpoints/Before_ScenarioSync.xlsx' if stage=='before' else ROOT/'KK01_V06_Executive_Master.xlsx'
 if prepare:old.make_lab();p.trim_private_lab();lab.freeze_unrelated_matrix_lab()
 return folder

def inspect(file,mode,corrected):
 with ZipFile(file) as z:cells,errors,stats=old.read(z)
 assert not errors,errors[:5]
 rows,metrics,flows=reference(mode,corrected);engine=cells['V06 Motor'];cc=old.REF['main_columns'];delta=0;count=0
 for r,d in enumerate(rows,5):
  for key,v in d.items():
   if not old.number(v):continue
   got=engine[col(cc[key])+str(r)];old.near(got,v);delta=max(delta,abs(got-v));count+=1
 for r,v in enumerate(flows,5):old.near(cells['V06 F01']['B'+str(r)],v)
 selected=[r for r in rows if r['op_age']==4 and r['available_room_nights']>0]
 checks=[(H,'E15',sum(r['ebitda'] for r in rows[-12:])),(H,'I15',sum(r['operator_fees'] for r in rows)),(H,'M15',metrics['cfads_total']),(Q,'E9',rows[53]['dividend']),(H,'E9',metrics['adr']),(H,'I9',metrics['occupancy']),(H,'M9',metrics['adr']*metrics['occupancy']),(H,'K44',metrics['room_revenue']),(H,'N44',metrics['fb_revenue']),(H,'E50',metrics['mice_revenue']),(H,'H50',metrics['revenue']),(H,'K50',metrics['revpar']),(H,'N50',metrics['ebitda']),(H,'M72',metrics['gop']),(H,'M74',metrics['operator_fees']),(H,'M76',metrics['ffe']),(H,'M80',metrics['margin']),(Q,'I9',metrics['dividends']),(Q,'M9',metrics['exit']),(Q,'E15',metrics['irr']),(Q,'I15',metrics['moic']),(Q,'M15',metrics['npv']),(REPORT,'G17',metrics['npv']),(REPORT,'A312',sum(r['adr'] for r in selected)/len(selected)),(REPORT,'G312',sum(r['occupancy'] for r in selected)/len(selected)),(REPORT,'A317',metrics['revpar']),(REPORT,'G317',metrics['ebitda']),(REPORT,'I338',metrics['revenue']),(REPORT,'I340',metrics['gop']),(REPORT,'I342',metrics['margin']),(REPORT,'I344',metrics['operator_fees']),(REPORT,'I346',metrics['ffe']),('V06 PDF 07','G17',metrics['revenue']),('V06 PDF 07','J17',metrics['ebitda'])]
 for sn,a,v in checks:old.near(cells[sn][a],v)
 if corrected:
  assert cells['V06 Controlos']['B67']==metrics['profile'];assert cells['V06 Registo Ativo']['B67']==metrics['profile'];assert cells['EXEC V06']['H27']==mode
  old.near(cells['EXEC V06']['H29'],metrics['adr']);old.near(cells['EXEC V06']['N29'],metrics['occupancy'])
 visual=cells['V06 Visual'];stocks={'cash','debt','reserve'};visual_count=0
 for address,key in visual.items():
  if not re.fullmatch('[A-Z]+4',address) or key not in cc:continue
  _,c=xl_cell_to_rowcol(address)
  if c>21 and not 150<=c<169:continue
  for year in range(1,16):
   rr=rows[(year-1)*12:year*12]
   expected=rr[-1][key] if key in stocks else sum(r[key] for r in rr)
   if key=='ebitda_margin':expected=sum(r['ebitda'] for r in rr)/sum(r['revenue'] for r in rr) if sum(r['revenue'] for r in rr)>0 else 0
   old.near(visual[col(c)+str(year+4)],expected);visual_count+=1
 # Verify chart values and chart data references after the external recalculation.
 chart_points=0;ns={'c':'http://schemas.openxmlformats.org/drawingml/2006/chart'}
 with ZipFile(file) as z:
  for part in z.namelist():
   if not re.fullmatch('xl/charts/chart[0-9]+.xml',part):continue
   root=E.fromstring(z.read(part))
   for val in root.findall('.//c:val/c:numRef',ns):
    formula=val.findtext('c:f',namespaces=ns) or '';match=re.fullmatch(r"(?:'([^']+)'|([^!]+))!\$?([A-Z]+)\$?(\d+):\$?([A-Z]+)\$?(\d+)",formula)
    assert match,formula
    sn=match[1] or match[2];assert sn=='V06 Visual',formula;assert match[3]==match[5],formula
    cache=val.find('c:numCache',ns);assert cache is not None,part
    for point in cache.findall('c:pt',ns):
     row=int(match[4])+int(point.get('idx'));expected=cells[sn][match[3]+str(row)];old.near(float(point.findtext('c:v',namespaces=ns)),expected);chart_points+=1
 assert chart_points>100,chart_points
 return {'mode':mode,'metrics':metrics,'monthly_values_compared':count,'max_absolute_monthly_difference':delta,'dashboard_report_bindings':len(checks),'graph_projection_values':visual_count,'chart_cached_points_verified':chart_points,'errors':0,'status':'PASS'},cells

def stage(name):
 existing='--verify-existing' in sys.argv;folder=setup(name,not existing);corrected=name=='after';results=[]
 def run(i,mode):
  test=f'Sync_{name}_{i}';changes=[(H,'H23',mode)]
  if mode=='Custom':changes += [(H,'H25',200),(H,'H27',.65)]
  if not existing:
   result=old.run_case(test,changes);assert result['status']=='PASS',result
  file=folder/'mutations'/test/'recalc/out'/(test+'.xlsx');proof,_=inspect(file,mode,corrected);print(json.dumps(proof,ensure_ascii=False),flush=True);return proof
 with ThreadPoolExecutor(max_workers=3) as ex:
  for f in as_completed([ex.submit(run,i,mode) for i,mode in enumerate(MODES)]):results.append(f.result())
 (BASEWORK/f'Sync_{name}.json').write_text(json.dumps(results,ensure_ascii=False,indent=2))

def sequence():
 folder=setup('sequence');steps=[];base=None;baseline_file=None
 for i,mode in enumerate(['Base Case','Conservative','Favorable','Base Case','Base Case']):
  name=f'Sync_Sequence_{i}';result=old.run_case(name,[(H,'H23',mode)]);assert result['status']=='PASS',result
  file=folder/'mutations'/name/'recalc/out'/(name+'.xlsx');proof,cells=inspect(file,mode,True)
  if i==0:base=cells;baseline_file=file
  if i>=3:
   compared=0;maximum=0;unequal=0
   for addr,v in base['V06 Motor'].items():
    if old.number(v):
     got=cells['V06 Motor'][addr];old.near(got,v);maximum=max(maximum,abs(got-v));unequal+=got!=v;compared+=1
   for sn,addr in [(Q,'M15'),(Q,'E15'),(Q,'I15'),(Q,'M9'),(REPORT,'G17')]:assert cells[sn][addr]==base[sn][addr],(sn,addr)
   proof['roundtrip']={'monthly_values':compared,'maximum_difference':maximum,'non_bit_identical_values':unequal,'financial_kpis_exactly_equal':True,'operation':'Save/reopen recalculation' if i==4 else 'Base after Conservative and Favorable'}
  steps.append(proof);print(json.dumps(proof,ensure_ascii=False),flush=True)
  shutil.copy2(file,folder/'Mutation_Lab.xlsx')
 (BASEWORK/'Sync_sequence.json').write_text(json.dumps(steps,ensure_ascii=False,indent=2))
 # Fresh EN workbook for the final PDFs, using the current template and baseline.
 folder=setup('english');r=old.run_case('Sync_EN',[('DASH 10 Centro','H7','EN')]);assert r['status']=='PASS',r
 file=folder/'mutations/Sync_EN/recalc/out/Sync_EN.xlsx'
 with ZipFile(file) as z:cells,errors,_=old.read(z)
 assert not errors and 'Executive Dashboard' in cells['EXEC V06']['B8'];assert cells['EXEC V06']['E26']=='Active Scenario'
 dest=BASEWORK/'mutations/Premium_EN/recalc/out/Premium_EN.xlsx';shutil.copy2(file,dest)
 (BASEWORK/'Sync_english.json').write_text(json.dumps({'status':'PASS','active_scenario_english':cells['EXEC V06']['E26'],'errors':0},indent=2))

def final():
 before=json.loads((BASEWORK/'Sync_before.json').read_text());after=json.loads((BASEWORK/'Sync_after.json').read_text());seq=json.loads((BASEWORK/'Sync_sequence.json').read_text());en=json.loads((BASEWORK/'Sync_english.json').read_text());assert len(before)==len(after)==len(seq)==5
 proof={'status':'PASS','platform':'LibreOffice cache-invalidated recalculation, existing core executed numerically in Python as independent test reference','native_excel365':'NOT TESTED','before':before,'after':after,'sequence':seq,'english':en,'source':'V06 Controlos!B67','source_sha256':hashlib.sha256((BASEWORK/'checkpoints/Before_ScenarioSync.xlsx').read_bytes()).hexdigest(),'result_sha256':hashlib.sha256((ROOT/'KK01_V06_Executive_Master.xlsx').read_bytes()).hexdigest()}
 (ROOT/'Sincronizacao_Cenarios.json').write_text(json.dumps(proof,ensure_ascii=False,indent=2))
 fields=['stage','mode','profile','adr','occupancy','room_revenue','fb_revenue','mice_revenue','gop','ebitda','margin','operator_fees','ffe','cash_end','cfads_total','dividends','irr','moic','npv','exit']
 with (ROOT/'Comparacao_Cenarios.csv').open('w',encoding='utf-8-sig',newline='') as f:
  w=csv.DictWriter(f,fieldnames=fields,delimiter=';',lineterminator='\n',extrasaction='ignore');w.writeheader()
  for label,records in [('before',before),('after',after)]:
   for r in records:w.writerow({'stage':label,'mode':r['mode'],**r['metrics']})
 print('Scenario synchronisation before/after and exact roundtrip PASS')
if __name__=='__main__':
 name=sys.argv[1]
 if name in ['before','after']:stage(name)
 elif name=='sequence':sequence()
 else:final()
