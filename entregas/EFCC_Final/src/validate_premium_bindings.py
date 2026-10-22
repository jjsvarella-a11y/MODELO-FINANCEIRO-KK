"""Compare every generated report binding and preserved dashboard KPI to recalc."""
import json,sys
from premium_ui import Source,ROOT,WORK,BASE,DASHES

def validate():
 spec=json.loads((WORK/'Premium_Spec.json').read_text());source=Source(WORK/'premium_final_recalc/out/KK01_V06_Executive_Master.xlsx');base=Source(BASE);diff=[];n=0
 def check(sn,addr,expected):
  nonlocal n
  _,got=source.get(sn,addr);n+=1
  if isinstance(expected,(float,int)):
   if not isinstance(got,(float,int)) or abs(got-expected)>max(.00005,abs(expected)*1e-9):diff.append([sn,addr,expected,got])
  elif got!=expected:diff.append([sn,addr,expected,got])
 for b in spec['report_bindings']:check(b['sheet'],b['cell'],b['value'])
 for sn,_,_ in DASHES[:9]:
  for addr in ['E9','I9','M9','E15','I15','M15']:check(sn,addr,base.get(sn,addr)[1])
 formats=0
 for sn,_,_ in DASHES[:9]:
  for addr in ['E9','I9','M9','E15','I15','M15']:
   expected=base.card_kind(sn,addr);got=source.card_kind(sn,addr);formats+=1
   if expected!=got:diff.append([sn,addr,'format',expected,got])
 for sn,addr,expected in [('DASH 10 V06','G312','pct'),('DASH 10 V06','A372','pct'),('DASH 10 V06','G372','pct'),('DASH 10 V06','A497','pct'),('DASH 10 V06','I38','ratio'),('DASH 10 V06','G552','pct'),('DASH 10 V06','A557','ratio'),('V06 Brief','G72','pct'),('V06 Brief','A77','ratio')]:
  got=source.card_kind(sn,addr);formats+=1
  if got!=expected:diff.append([sn,addr,'format',expected,got])
 result={'formats_checked':formats,'status':'PASS' if not diff else 'FAIL','platform':'LibreOffice independent full recalculation and Python source values','native_excel365':'NOT TESTED','numeric_and_text_bindings_checked':n,'differences':diff};(ROOT/'Validacao_Vinculos_Premium.json').write_text(json.dumps(result,indent=2,ensure_ascii=False));print(result);assert not diff,diff[:10]
if __name__=='__main__':validate()
