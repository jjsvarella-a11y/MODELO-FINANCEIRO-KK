import sys,json
sys.path.insert(0,'/workspace/MODELO-FINANCEIRO-KK/entregas/EFCC_Final/src')
import test_closure_r5 as t
from concurrent.futures import ThreadPoolExecutor,as_completed
base,errors,_=t.load(t.WORK/'final_recalc/out'/t.MASTER.name);assert not errors
labs={False:t.make_lab(False,'_numeric'),True:t.make_lab(True,'_numeric')}
cases=[('Discount_Numeric8',[('DASH 07 V06','H29',.08)],True),('Discount_Numeric10',[('DASH 07 V06','H29',.1)],False),('Preset_Numeric12',[('DASH 07 V06','H29',.12)],False)]
with ThreadPoolExecutor(max_workers=3) as ex:results=list(ex.map(lambda c:t.run_case(c,base,labs),cases))
assert all(r['status']=='PASS' for r in results),results
for r,rate in zip(results,[.08,.1,.12]):t.near(r['rate'],rate);t.near(r['irr'],base['V06 Cenarios']['K5'])
(t.OUT/'Testes_Presets_Numericos_R5.json').write_text(json.dumps({'status':'PASS','tests':results,'native_excel365':'NOT TESTED'},indent=2))
