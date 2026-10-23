import sys,json
sys.path.insert(0,'/workspace/MODELO-FINANCEIRO-KK/entregas/EFCC_Final/src')
import test_closure_r5 as t
from concurrent.futures import ThreadPoolExecutor
lab=t.make_lab(False,'_boundary')
def run(item):
 name,changes=item;folder=t.WORK/name;folder.mkdir(exist_ok=True);file=folder/(name+'.xlsx');t.mutate(lab,file,changes);file=t.recalculate(file,folder/'recalc');c,errors,_=t.load(file);assert not errors,errors[:5]
 if name=='Promoter_Final':
  assert c[t.INV]['C29']=='N/A';assert c['DASH 06 V06']['V200']=='N/A'
 if name=='ETC_200M':assert c[t.IMP]['C10']=='N/A' and c[t.IMP]['D10']=='N/A' and c['EXEC V06']['X31']=='N/A'
 result={'test':name,'status':'PASS','errors':0,'npv':c['V06 Cenarios']['M5'],'irr':c[t.IMP]['C10'],'changes':changes};print(result,flush=True);return result
cases=[('Promoter_Final',[('DASH 06 V06','H27','Promoters'),('DASH 07 V06','H33','Promoters')]),('ETC_200M',[('DASH 01 V06','H25',200000000)]),('Bank_Zero',[('DASH 03 V06','H31',0)]),('Pricing_Min',[('DASH 04A V06','H25','Custom'),('DASH 04A V06','N25',.01)])]
with ThreadPoolExecutor(max_workers=3) as ex:results=list(ex.map(run,cases))
(t.OUT/'Testes_Limites_R5.json').write_text(json.dumps({'status':'PASS','tests':results,'native_excel365':'NOT TESTED'},indent=2))
