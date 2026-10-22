from engine import *
from waterfall import constant_policy,priority,COUNTS,TICKETS

def selected(result):
 p=result['p'];legacy=p['capital_scenario']!='KK-E25' and p['equity_mode']!='Combinação tickets';mode='Proporcional' if legacy else p['waterfall_mode'];constant=constant_policy(result)
 w=constant if mode=='Otimizado' and constant['feasible'] else priority(result,'Combinado' if mode=='Otimizado' else mode)
 rows=[];plans=pipeline(p);scale=(p['equity']-p['conversion'])/25e6 if p['equity_mode']!='Combinação tickets' else 1;ss=summary(result);assets=max(0,ss['ev']*(1-p['exit_cost'])-result['flow'][-1]['debt']+result['flow'][-1]['cash']-result['flow'][-1]['hothold']-result['flow'][-1]['conthold']);recoverable=min(result['flow'][-1]['shareloan'],assets)
 for i,row in enumerate(w['rows']):
  m=i+1;d=row['dividends'];e=row['exit'];counts=[]
  for cat in ['Platinum','Gold','Silver']:counts.append(sum(min(1,x['planned_received']/(x['ticket']*scale)) for x in plans if x['category']==cat and x['pay_month']<=m) if scale>0 else 0)
  amount=[(d[k]+e[k])/COUNTS[k] for k in range(3)];total=sum(d)+sum(e)
  new=total*economic_share(p,m) if legacy else sum(amount[k]*counts[k] for k in range(3));jscv=total*p['legacy_jscv'] if legacy else amount[0]*(1+.55*(1-counts[0]))+2*amount[1];jals=total-new-jscv;paid=investor_outlay(p,m);f=result['flow'][i];credit=f['loan_interest']+f['loan_principal']+(recoverable if m==180 else 0);promoter=jscv+jals+paid*(1-p['treasury_ratio'])+credit
  rows.append({'m':m,'dividends':d,'exit':e,'new_counts':counts,'new_received':new,'jscv_received':jscv,'jals_received':jals,'new_outlay':paid,'new_cf':new-paid,'promoter_cf':promoter,'check':total-new-jscv-jals})
 vals={}
 for who,key,initial in [('new','new_cf',0),('promoter','promoter_cf',-p['historic'])]:
  cf=[initial]+[r[key] for r in rows];vals[who+'_irr']=irr(cf);vals[who+'_npv']=sum(x/(1+p['equity_discount'])**(m/12) for m,x in enumerate(cf))
 vals['new_moic']=sum(x['new_received'] for x in rows)/sum(x['new_outlay'] for x in rows) if sum(x['new_outlay'] for x in rows)>0 else None
 return {'rows':rows,'method':'Pesos exatos' if mode=='Otimizado' and constant['feasible'] else mode+' / prioridade','constant_feasible':constant['feasible'],'metrics':vals,'categories':w['categories'],'terminal_credit_recoverable':recoverable,'terminal_credit_gap':result['flow'][-1]['shareloan']-recoverable}
