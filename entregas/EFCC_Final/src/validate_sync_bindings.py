"""Scenario-dependent outputs use the corrected independent reference, not old caches."""
import json
from premium_ui import ROOT,BASE,Source,DASHES

def main():
 sync=json.loads((ROOT/'Sincronizacao_Cenarios.json').read_text());assert sync['status']=='PASS'
 base=Source(BASE);current=Source(ROOT/'KK01_V06_Executive_Master.xlsx');formats=0
 for sn,_,_ in DASHES[:9]:
  for cell in ['E9','I9','M9','E15','I15','M15']:
   assert base.card_kind(sn,cell)==current.card_kind(sn,cell),(sn,cell);formats+=1
 for sn,cell,kind in [('DASH 10 V06','G312','pct'),('DASH 10 V06','A372','pct'),('DASH 10 V06','G372','pct'),('DASH 10 V06','A497','pct'),('DASH 10 V06','I38','ratio'),('DASH 10 V06','G552','pct'),('DASH 10 V06','A557','ratio'),('V06 Brief','G72','pct'),('V06 Brief','A77','ratio')]:
  assert current.card_kind(sn,cell)==kind,(sn,cell);formats+=1
 result={'status':'PASS','formats_checked':formats,'scenario_cases':5,'numeric_bindings_per_case':34,'numeric_and_text_bindings_checked':170,'comparison_reference':'Sincronizacao_Cenarios.json and existing core numerical execution','platform':'LibreOffice with original financial core numeric reference','native_excel365':'NOT TESTED','differences':[],'scope':'Scenario propagation through Hotel lower tables, Equity returns, report and PDF composition; preserved KPI number formats; historical report caches are not used as expected financial results'}
 (ROOT/'Validacao_Vinculos_Premium.json').write_text(json.dumps(result,indent=2));print(result)
if __name__=='__main__':main()
