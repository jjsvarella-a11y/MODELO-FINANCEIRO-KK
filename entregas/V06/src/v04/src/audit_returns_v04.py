"""Independent horizons10/15, selected waterfall, recoverability and completion checks."""
from engine import *
from selected_policy import selected
from waterfall import constant_policy,BASES,SHARES,COUNTS,TICKETS
from completion import completion
import numpy as np,openpyxl

def tenyear(r):
 p=r['p'];sel=selected(r);legacy=p['capital_scenario']!='KK-E25' and p['equity_mode']!='Combinação tickets';mode='Proporcional' if legacy else p['waterfall_mode'];cp=constant_policy(r);bal=BASES.copy();rates=np.array([p['target_platinum'],p['target_gold'],p['target_silver']])
 for i,row in enumerate(sel['rows'][:120]):
  if i:bal*=np.power(1+rates,1/12)
  bal=np.maximum(0,bal-np.array(row['dividends']))
 sm=summary(r,120);assets=max(0,sm['ev']*(1-p['exit_cost'])-r['flow'][119]['debt']+r['flow'][119]['cash']-r['flow'][119]['hothold']-r['flow'][119]['conthold']);credit=min(r['flow'][119]['shareloan'],assets);exit=max(0,assets-r['flow'][119]['shareloan'])
 if mode=='Otimizado' and cp['feasible']:alloc=exit*np.array(cp['exit_weights'])
 elif mode in ['Saída','Combinado','Otimizado']:
  avail=exit;alloc=np.zeros(3)
  for k in range(3):alloc[k]=min(avail,max(0,bal[k]));avail-=alloc[k]
  alloc+=avail*SHARES
 else:alloc=exit*SHARES
 counts=np.array(sel['rows'][119]['new_counts']);newex=exit*economic_share(p,120) if legacy else sum(alloc/COUNTS*counts);newcf=[0]+[x['new_cf']+(newex if x['m']==120 else 0) for x in sel['rows'][:120]];paid=sum(x['new_outlay'] for x in sel['rows'][:120]);received=sum(x['new_received'] for x in sel['rows'][:120])+newex
 cats=[]
 for k in range(3):cats.append(irr([0]+[x['dividends'][k]/COUNTS[k]-(TICKETS[k] if x['m']==1 else 0)+(alloc[k]/COUNTS[k] if x['m']==120 else 0) for x in sel['rows'][:120]]))
 return {'B52':exit,'B53':credit,'B55':irr(newcf),'B56':received/paid if paid else None,'B57':sum(v/(1+p['equity_discount'])**(m/12) for m,v in enumerate(newcf)),'F47':cats[0],'F48':cats[1],'F49':cats[2]}

def audit(path,params):
 r=run(config(**params));sel=selected(r);w=openpyxl.load_workbook(path,read_only=True,data_only=True);values={c.coordinate:c.value for row in w['V04 Retornos'].iter_rows(max_row=60,max_col=11) for c in row if hasattr(c,'coordinate')};expected=tenyear(r);expected.update({'B27':sel['metrics']['new_irr'],'B28':sel['metrics']['new_moic'],'B29':sel['metrics']['new_npv'],'B30':sum(x['jscv_received'] for x in sel['rows']),'B31':sum(x['jals_received'] for x in sel['rows']),'B34':sum(x['promoter_cf'] for x in sel['rows'])});fail=[]
 for addr,v in expected.items():
  a=values.get(addr)
  if v is None:
   if isinstance(a,(int,float)):fail.append((addr,a,v))
  elif not isinstance(a,(int,float)) or abs(a-v)>max(1e-7,abs(v)*1e-8):fail.append((addr,a,v))
 cv=completion(r);cells={c.coordinate:c.value for row in w['V04 Completion'].iter_rows(max_row=21,max_col=3) for c in row if hasattr(c,'coordinate')}
 for addr,key in [('B8','projected_bank'),('B17','projected_coverage'),('B18','projected_gap')]:
  a=cells[addr];v=cv[key]
  if not isinstance(a,(int,float)) or abs(a-v)>max(.01,abs(v)*1e-9):fail.append((addr,a,v))
 w.close();return {'case':path.stem,'ok':not fail,'comparisons':len(expected)+3,'failures':fail}
if __name__=='__main__':
 cases=json.loads((ROOT/'entrega/Testes_Seletores.json').read_text());out=[audit(Path(x['output']),x['params']) for x in cases if x['ok']];(ROOT/'entrega/Testes_Retornos_Completion.json').write_text(json.dumps(out,ensure_ascii=False,indent=2));print(json.dumps(out,ensure_ascii=False,indent=2));raise SystemExit(not all(x['ok'] for x in out))
