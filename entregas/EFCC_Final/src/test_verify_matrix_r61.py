import sys,json
sys.path.insert(0,'/workspace/MODELO-FINANCEIRO-KK/entregas/EFCC_Final/src')
from bank_emit import *
src=Source(BW/'emissions/KK01-BANK-20261023T175546Z/selected/recalc/out/Input.xlsx');rate=src.get('V06 Registo Ativo','B65')[1];points=[]
for ri,idx in enumerate([0,2,3,4,5]):
 for ci in range(4):
  num=26+idx*5+ci;v=sum((src.get(f'V06 F{num:02}','B'+str(k+5))[1] or 0)/(1+rate)**(k/12) for k in range(181));row=5+idx*5+ci
  near(v,src.get('V06 Matriz Valores','C'+str(row))[1]);near(v,src.get('DASH 05 V06',col(8+ci*4)+str(120+ri*2))[1]);near(v,src.get('DASH 10 V06',col(2+ci*2)+str(508+ri*2))[1]);points.append({'adr':[150,175,200,225][ci],'occupancy':[.5,.6,.65,.7,.75][ri],'npv':v})
near(src.get('V06 Cenarios','M39')[1],src.get('V06 Cenarios','M5')[1]);(BOUT/'Bank_Matrix_20.json').write_text(json.dumps({'status':'PASS','scenario':'Proposal Bank A TAN 5.5% / grace42','npv_points':points,'current_case_reconciled':True,'native_excel365':'NOT TESTED'},indent=2));print('20 independent bank sensitivity points PASS')
