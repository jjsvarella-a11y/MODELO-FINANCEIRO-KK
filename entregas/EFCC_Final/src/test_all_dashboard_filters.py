"""Exercise every exposed input and every preset; compare recalculated active engine."""
import json,re,csv,sys,subprocess
from zipfile import ZipFile
from concurrent.futures import ThreadPoolExecutor,as_completed
import test_premium_controls as p
from premium_ui import Source,BASE,legacy,DASHES
old=p.old;ROOT=p.ROOT;WORK=p.WORK
H='DASH 05 V06';R='DASH 04B V06';Q='DASH 07 V06';C='DASH 10 Centro'
def read_result(name):
 with ZipFile(WORK/'mutations'/name/'recalc/out'/(name+'.xlsx')) as z:return old.read(z)[0]
def freeze_unrelated_matrix_lab():
 from zipfile import ZIP_DEFLATED
 from assemble import paths,T,E,xml
 source=WORK/'Mutation_Lab.xlsx';temp=source.with_suffix('.active-only.xlsx')
 with ZipFile(source) as z,ZipFile(temp,'w',ZIP_DEFLATED,compresslevel=3) as out:
  _,pp=paths(z);parts={pp[n] for n in ['V06 M01','V06 F26','V06 Calend 02','V06 Calend 03','V06 Calend 04']}
  for item in z.infolist():
   data=z.read(item.filename)
   if item.filename in parts:
    root=E.fromstring(data)
    for c in root.iter(T('c')):
     f=c.find(T('f'))
     if f is not None:c.remove(f)
    data=xml(root)
   out.writestr(item.filename,data)
 temp.replace(source)
def main():
 existing='--verify-existing' in sys.argv
 if not existing:
  old.make_lab();p.trim_private_lab();freeze_unrelated_matrix_lab()
  result=old.run_case('Filters_Baseline',[]);assert result['status']=='PASS',result
 base=read_result('Filters_Baseline');inputs=json.loads((ROOT/'Parametros_Interface.json').read_text());cases=[]
 options={(H,'H23'):['Conservative','Base Case','Favorable','Upside Stress','Custom'],(H,'H32'):['Operator Proposed Terms','KK Target Terms','KK Optimized Terms','Custom Terms'],(R,'H23'):['Reference Study','KK Target Terms','KK Optimized Terms','Custom Terms'],(Q,'H23'):['Conservative','Balanced','Flexible'],(Q,'N23'):['Auto','Minimum','Intermediate','Maximum','Partial','Suspended'],(Q,'H29'):['8%','10%','12%','Custom'],(C,'H7'):['PT','EN'],(C,'H11'):['Complete Report','Executive Brief'],(C,'N11'):['Active Scenario','Reference Case']}
 for n,item in enumerate(inputs):
  sn,a,v=item['sheet'],item['cell'],item['value'];choices=options.get((sn,a))
  for j,value in enumerate(choices or [None]):
   if value is None:
    if isinstance(v,(int,float)):value=1 if v==0 else (0 if v==1 else v+1 if v>=1 else v+.01)
    else:value='2026-11-01' if a=='H15' else 'FILTER-TEST'
   value={(H,'H25'):215,(H,'N25'):.55,(H,'N27'):.055,(H,'N29'):.035}.get((sn,a),value)
   if isinstance(value,float):value=round(value,12)
   changes=[(sn,a,value)]
   if sn==H and a in ['H25','H27']:changes.insert(0,(H,'H23','Custom'))
   if sn==H and a in ['H29','H30','N32','H34','N34']:changes.insert(0,(H,'H32','Custom Terms'))
   if sn==H and a=='H36':changes.append((H,'N36',120000))
   if sn==H and a=='N36':changes.insert(0,(H,'H36',1))
   if sn==R and a in ['H25','N25']:changes.insert(0,(R,'H23','Custom Terms'))
   if sn==Q and a=='N25':changes.insert(0,(Q,'H29','Custom'))
   if sn==Q and a=='H27':changes.insert(0,(Q,'N23','Partial'))
   if sn==Q and a=='N23' and value=='Partial':changes.append((Q,'H27',.35))
   cases.append((f'Filter_{n:02}_{j}',sn,a,value,changes,item))
 def run(case):
  name,sn,a,value,changes,item=case;result={'test':name,'status':'PASS','evidence':{}} if existing else old.run_case(name,changes)
  if result['status']!='PASS':return result
  try:
   cells=read_result(name);m=cells['V06 Motor'];changed=[a for a,b in base['V06 Motor'].items() if old.number(b) and old.number(m.get(a)) and abs(m[a]-b)>max(.00005,abs(b)*1e-9)]
   kpi_changed=[sn2+'!'+addr for sn2,_,_ in DASHES[:9] for addr in ['E9','I9','M9','E15','I15','M15'] if old.number(cells[sn2].get(addr)) and old.number(base[sn2].get(addr)) and abs(cells[sn2][addr]-base[sn2][addr])>max(.00005,abs(base[sn2][addr])*1e-9)]
   metric_changed=[a for a in ['B5','C5','D5','E5','F5','J5','K5','L5','M5','N5','T5'] if old.number(base['V06 Cenarios'].get(a)) and old.number(cells['V06 Cenarios'].get(a)) and abs(cells['V06 Cenarios'][a]-base['V06 Cenarios'][a])>max(.00005,abs(base['V06 Cenarios'][a])*1e-9)]
   if sn==H and a in ['H25','H27','N25','N27','N29','H36','N23']:
    minimum,step={'H25':(120,5),'H27':(.45,.01),'N25':(.2,.05),'N27':(.02,.005),'N29':(0,.005),'H36':(0,1),'N23':(1,1)}[a];assert abs((value-minimum)/step-round((value-minimum)/step))<1e-7
   oldvalue=base[sn][a]
   if old.number(value):old.near(cells[sn][a],value)
   else:assert cells[sn][a]==value,(sn,a,value,cells[sn][a])
   for active,source in [('H24','E9'),('H26','I9')]:old.near(cells[H][active],cells[H][source])
   for target,row in [('H28',68),('N28',69),('N30',70)]:old.near(cells[H][target],cells['V06 Registo Ativo']['B'+str(row)])
   for target,row in [('H24',94),('N24',95)]:old.near(cells[R][target],cells['V06 Registo Ativo']['B'+str(row)])
   old.near(cells[Q]['N26'],cells['V06 Registo Ativo']['B65'])
   if (sn,a)==(H,'H23'):
    adr,occ={'Conservative':legacy.HOTELS['Conservador'][:2],'Base Case':legacy.HOTELS['Base'][:2],'Favorable':legacy.HOTELS['Favorável'][:2],'Upside Stress':(250,.75),'Custom':(210,.64)}[value]
    old.near(cells[H]['H24'],adr);old.near(cells[H]['H26'],occ);old.near(cells[H]['H25'],210);old.near(cells[H]['H27'],.64)
   if (sn,a)==(H,'H32'):
    expected={'Operator Proposed Terms':(.04,.04,.10),'KK Target Terms':(.025,.03,.06),'KK Optimized Terms':(.02,.015,.05),'Custom Terms':(.03,.03,.06)}[value]
    for target,v in zip(['H28','N28','N30'],expected):old.near(cells[H][target],v)
    old.near(cells[H]['E9'],base[H]['E9']);old.near(cells[H]['I9'],base[H]['I9'])
   if (sn,a)==(R,'H23'):
    expected={'Reference Study':(.08,.03),'KK Target Terms':(.06,.02),'KK Optimized Terms':(.05,.015),'Custom Terms':(.03,.05)}[value]
    for target,v in zip(['H24','N24'],expected):old.near(cells[R][target],v)
    old.near(m['AP35'],base['V06 Motor']['AP35']);old.near(m['AQ35'],base['V06 Motor']['AQ35'])
   if (sn,a)==(Q,'H29'):old.near(cells[Q]['N26'],cells[Q]['N25'] if value=='Custom' else float(value.rstrip('%'))/100)
   if sn==C and a in ['H15','N15']:
    assert value in cells['DASH 10 V06']['A8']
    for owner,_,_ in DASHES[:9]:assert value in cells[owner]['E1']
   # Presentation controls intentionally preserve the financial engine.
   if sn==C:assert not changed,changed[:5]
   # Presets equal to the calibrated baseline legitimately have identical results.
   neutral=(sn==H and a=='H23' and value in ['Base Case']) or value==oldvalue or sn==C or (sn,a)==(H,'N23') or ((sn,a) in [('DASH 03 V06','N23'),('DASH 03 V06','H25')] and not changed and not metric_changed) or ((sn,a)==(Q,'H29') and value=='Custom' and cells[Q]['N26']==base[Q]['N26'])
   if (sn,a) in [('DASH 03 V06','N23'),('DASH 03 V06','H25')] and not changed:
    bank=json.loads((ROOT/'Testes_Carencia_Prazo.json').read_text());assert bank['status']=='PASS'
   if (sn,a,value)==(Q,'N23','Minimum') and not changed and not metric_changed:
    from xlsxwriter.utility import xl_col_to_name
    target_col=xl_col_to_name(old.REF['main_columns']['dividend_target'])
    assert all(m[target_col+str(r)]==base['V06 Motor'][target_col+str(r)]==3000000 for r in range(5,185))
    neutral=True
    result['equivalence_evidence']={'reason':'Auto independently selects the same 3M annual target in all 180 months; Minimum changes the policy identifier but not its applicable cash flows','months_checked':180,'target':3000000}
   if not neutral:assert changed or metric_changed,('No engine response',sn,a,value)
   if (sn,a)==(H,'N23'):assert cells[H]['H50']!=base[H]['H50']
   if (sn,a)==(C,'H7') and value=='EN':assert 'Executive Dashboard' in cells['EXEC V06']['B8']
   if (sn,a)==(C,'N11'):old.near(cells['DASH 10 V06']['G17'],old.REF['metrics']['Reference Case']['npv'] if value=='Reference Case' else base['V06 Cenarios']['M5'])
   result.update({'sheet':sn,'cell':a,'choice':value,'label':item['label_pt'],'changed_engine_cells':len(changed),'changed_dashboard_kpis':kpi_changed,'changed_summary_cells':metric_changed,'active_readouts_checked':8,'interpretation':'Nonbinding while cash sweep repays debt; separately tested with outstanding debt' if (sn,a) in [('DASH 03 V06','N23'),('DASH 03 V06','H25')] and not changed else 'Equivalent to baseline / presentation only' if neutral else 'Engine responds'})
   from verify_all_filter_outputs import inspect as verify_outputs
   result['correspondences']=verify_outputs(result,cells)
  except Exception as e:result.update(status='FAIL',error=repr(e),sheet=sn,cell=a,choice=value)
  return result
 results=[]
 with ThreadPoolExecutor(max_workers=3) as ex:
  for f in as_completed([ex.submit(run,c) for c in cases]):
   r=f.result();results.append(r);print(json.dumps(r,ensure_ascii=False),flush=True)
 if not existing:
  subprocess.run([sys.executable,str(ROOT/'src/test_bank_controls.py')],check=True)
  existing=True
  with ThreadPoolExecutor(max_workers=3) as ex:results=list(ex.map(run,cases))
 proof={'status':'PASS' if all(r['status']=='PASS' for r in results) else 'FAIL','platform':'LibreOffice, cache-invalidated recalculation; live active engine, unrelated scenario branches frozen only in private labs','native_excel365':'NOT TESTED','inputs_covered':len(inputs),'tests':results,'release':'V06.3-SYNC','note':'Base Case uses the Base profile; stored Custom values remain editable but do not apply in preset modes.'}
 (ROOT/'Testes_Todos_Filtros.json').write_text(json.dumps(proof,indent=2,ensure_ascii=False))
 with (ROOT/'Matriz_Testes_Filtros.csv').open('w',newline='',encoding='utf-8-sig') as f:
  w=csv.DictWriter(f,fieldnames=['test','sheet','cell','label','choice','changed_engine_cells','status','error'],extrasaction='ignore',delimiter=';',lineterminator='\n');w.writeheader();w.writerows(results)
 assert proof['status']=='PASS'
if __name__=='__main__':main()
