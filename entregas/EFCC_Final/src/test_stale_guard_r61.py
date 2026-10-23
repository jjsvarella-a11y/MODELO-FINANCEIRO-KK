import sys,json
sys.path.insert(0,'/workspace/MODELO-FINANCEIRO-KK/entregas/EFCC_Final/src')
from bank_emit import *
import test_bank_r6 as tb
snapshot=BOUT/'Emissao_Demonstracao_6_vs_5_5/Snapshot.xlsx';tb.BM=snapshot;lab=tb.lab();results=[]
for name,changes in [('Unchanged',[]),('Changed',[(BN,'H33','Custom'),(BN,'H23',.08)])]:
 folder=BW/('guard_'+name);folder.mkdir(exist_ok=True);target=folder/'Input.xlsx';mutate(lab,target,changes);out=recalculate(target,folder/'recalc');c,e,_=read_bank(out);assert not e,e[:3];status=c[BN]['E588'];assert ('OUTDATED' in status or 'DESATUALIZADO' in status) == (name=='Changed'),status;results.append({'case':name,'status':'PASS','comparison_message':status,'npv':c['V06 Cenarios']['M5']})
(BOUT/'Snapshot_Stale_Guard.json').write_text(json.dumps({'status':'PASS','checks':results,'native_excel365':'NOT TESTED'},indent=2));print('Snapshot stale guard PASS')
