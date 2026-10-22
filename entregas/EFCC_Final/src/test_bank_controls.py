import sys,json
from test_all_dashboard_filters import old,read_result,ROOT
from concurrent.futures import ThreadPoolExecutor,as_completed
cases=[('Filter_BankGrace',[('DASH 03 V06','N25',0),('DASH 03 V06','N23',43)]),('Filter_BankTenor',[('DASH 03 V06','N25',0),('DASH 03 V06','H25',121)])]
results=[]
with ThreadPoolExecutor(max_workers=2) as ex:
 for f in as_completed([ex.submit(old.run_case,*case) for case in cases]):
  r=f.result();print(r,flush=True);results.append(r)
base=read_result('Filter_10_0')['V06 Motor']
for r in results:
 if r['status']=='PASS':
  m=read_result(r['test'])['V06 Motor'];changed=sum(1 for a,v in base.items() if old.number(v) and old.number(m.get(a)) and abs(m[a]-v)>max(.00005,abs(v)*1e-9));assert changed
  r['engine_cells_changed_relative_to_same_no_sweep_scenario']=changed
(ROOT/'Testes_Carencia_Prazo.json').write_text(json.dumps({'status':'PASS' if all(r['status']=='PASS' for r in results) else 'FAIL','baseline':'Filter_10_0: cash sweep disabled; only grace/tenor differ','tests':results,'native_excel365':'NOT TESTED'},indent=2))

assert all(r["status"]=="PASS" for r in results),results
