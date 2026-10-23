import sys,json
from concurrent.futures import ThreadPoolExecutor,as_completed
sys.path.insert(0,'/workspace/MODELO-FINANCEIRO-KK/entregas/EFCC_Final/src')
from bank_emit import *
import test_bank_r6 as tb
source=tb.lab();base=evaluate(BM);base['cells']=read_bank(BM)[0]
cases=[('InterestGrace_Excess',[(BN,'N456',24),(BN,'H474','Excess cash')]),('FeesDefinition_0',[(BN,'H464',10000),(BN,'N464',1),(BN,'H27',150),(BN,'N476',0)]),('FeesDefinition_1',[(BN,'H464',10000),(BN,'N464',1),(BN,'H27',150),(BN,'N476',1)]),('Official_R61',[(BN,'H33','Official KK-01 6%')]),('ACT365_R61',[(BN,'H472','ACT/365')])]
results=[]
with ThreadPoolExecutor(max_workers=3) as ex:
 for f in as_completed([ex.submit(tb.case,x,source,base) for x in cases]):results.append(f.result())
assert all(r['status']=='PASS' for r in results),results
by={r['test']:r for r in results};near(by['FeesDefinition_0']['npv'],by['FeesDefinition_1']['npv']);near(by['FeesDefinition_0']['metrics']['Cumulative dividends'],by['FeesDefinition_1']['metrics']['Cumulative dividends']);assert by['FeesDefinition_0']['metrics']['Minimum DSCR']!=by['FeesDefinition_1']['metrics']['Minimum DSCR']
(BOUT/'Testes_R61_Adicionais.json').write_text(json.dumps({'status':'PASS','tests':results,'DSCR_definition_does_not_change_cash_NPV_or_dividends':True,'native_excel365':'NOT TESTED'},indent=2,ensure_ascii=False));print('Additional R6.1 checks PASS',flush=True)
