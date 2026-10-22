"""Constrained payout allocation. No additional project cash is created.
LP variables are distributions per category/event; targets are conventional bases,
not fabricated promoter contributions. Uniform cash per block within each class.
"""
import numpy as np
from scipy.optimize import linprog
from engine import run,config,summary,irr
CATEGORIES=['Platinum','Gold','Silver'];COUNTS=np.array([2.,3.,15.]);TICKETS=np.array([1e7,5e6,1e6]);BASES=COUNTS*TICKETS;TARGETS=np.array([.10,.06,.055]);SHARES=BASES/BASES.sum()
def solve(result,mode='Combinado',target_scale=1):
 p=result['p'];s=summary(result);div=np.array([x['dividend'] for x in result['flow']]);ex=s['exit_equity'];times=np.r_[np.flatnonzero(div>.000001)+1,180] if ex>0 else np.flatnonzero(div>.000001)+1;pool=np.r_[div[div>.000001],ex] if ex>0 else div[div>.000001];kinds=['dividend']*sum(div>.000001)+(['exit'] if ex>0 else []);N=len(times);rates=np.array([p['target_platinum'],p['target_gold'],p['target_silver']])*target_scale
 factors=np.array([[(1+r)**(-(m-1)/12) for m in times] for r in rates]);nv=3*N
 eq=np.zeros((N,nv))
 for j in range(N):eq[j,[j,N+j,2*N+j]]=1
 bounds=[(0,None)]*nv
 if mode=='Proporcional':a=np.outer(SHARES,pool);status='Distribuição proporcional';ok=False;exact_ok=False
 else:
  # A target IRR is a linear discounted-cash constraint for a conventional outlay M1.
  pv=np.zeros((3,nv))
  for k in range(3):pv[k,k*N:(k+1)*N]=factors[k]
  for k in range(3):
   for j,m in enumerate(times):
    fixed=(mode=='Só saída' and kinds[j]=='dividend') or (mode=='Só dividendos' and kinds[j]=='exit')
    if fixed:bounds[k*N+j]=(pool[j]*SHARES[k],pool[j]*SHARES[k])
  exact=linprog(np.zeros(nv),A_eq=np.vstack([eq,pv]),b_eq=np.r_[pool,BASES],bounds=bounds,method='highs')
  exact_ok=exact.success
  floor=linprog(np.zeros(nv),A_ub=-pv,b_ub=-BASES,A_eq=eq,b_eq=pool,bounds=bounds,method='highs');ok=floor.success
  if exact_ok:a=exact.x.reshape(3,N);status='Metas simultâneas matematicamente alcançáveis (simulação)'
  elif ok:
   weights=np.array([1e6,1e3,1.]);cost=sum(weights[k]*pv[k]/BASES[k] for k in range(3));near=linprog(cost,A_ub=-pv,b_ub=-BASES,A_eq=eq,b_eq=pool,bounds=bounds,method='highs');a=near.x.reshape(3,N);status='Metas mínimas alcançáveis; excedentes não são incumprimento'
  else:
   # Common max-min attainment, then lexicographic requested priority P,G,S.
   # Scaled PV shortfalls preserve meaningful comparison of different ticket sizes.
   A=np.hstack([-pv/BASES[:,None],-np.eye(3)]);b=-np.ones(3)
   E=np.hstack([eq,np.zeros((N,3))]);obj=np.r_[np.zeros(nv),[1e6,1e3,1.]]
   opt=linprog(obj,A_ub=A,b_ub=b,A_eq=E,b_eq=pool,bounds=bounds+[(0,None)]*3,method='highs')
   if not opt.success:raise ValueError(opt.message)
   a=opt.x[:nv].reshape(3,N);status='Metas não simultâneas; solução de menor desvio com prioridade Platinum/Gold/Silver'
  a=np.maximum(a,0)
 out=[]
 for k in range(3):
  cf=np.zeros(181);cf[1]=-TICKETS[k]
  np.add.at(cf,times,a[k]/COUNTS[k])
  value=irr(cf.tolist());out.append({'category':CATEGORIES[k],'target':float(rates[k]),'irr':value,'deviation':value-float(rates[k]) if value is not None else None,'conventional_basis':float(TICKETS[k]),'dividends':float(sum(a[k,j] for j in range(N) if kinds[j]=='dividend'))/COUNTS[k],'received':float(sum(a[k]))/COUNTS[k]})
 if mode=='Proporcional':ok=all(x['irr'] is not None and x['irr']+1e-8>=x['target'] for x in out);exact_ok=all(x['irr'] is not None and abs(x['irr']-x['target'])<1e-8 for x in out)
 return {'mode':mode,'feasible_targets':bool(ok),'exact_targets':bool(exact_ok),'kinds':kinds,'status':status,'events':times.tolist(),'pools':pool.tolist(),'allocations':a.tolist(),'categories':out,'reconciliation':float(max(abs(a.sum(axis=0)-pool))) if N else 0}
if __name__=='__main__':
 import json
 r=run(config());results=[solve(r,m) for m in ['Proporcional','Só dividendos','Só saída','Combinado']];print(json.dumps(results,ensure_ascii=False,indent=2))

def priority(result,mode='Combinado'):
 """Native Excel reproducible accrued preference; minimal target is not a guarantee."""
 p=result['p'];rates=np.array([p['target_platinum'],p['target_gold'],p['target_silver']]);balances=BASES.copy();rows=[];alloc=np.zeros((3,180));ex=summary(result)['exit_equity']
 for i,x in enumerate(result['flow']):
  if i>0:balances*=np.power(1+rates,1/12)
  opening=balances.copy();cash=x['dividend'];paid=np.zeros(3)
  if mode in ['Combinado','Dividendos']:
   for k in range(3):paid[k]=min(max(0,cash),max(0,balances[k]));cash-=paid[k];balances[k]-=paid[k]
   paid+=max(0,cash)*SHARES
  else:paid=cash*SHARES;balances=np.maximum(0,balances-paid)
  exitpaid=np.zeros(3)
  if i==179:
   cash=ex
   if mode in ['Combinado','Saída']:
    for k in range(3):exitpaid[k]=min(max(0,cash),max(0,balances[k]));cash-=exitpaid[k];balances[k]-=exitpaid[k]
    exitpaid+=max(0,cash)*SHARES
   else:exitpaid=cash*SHARES;balances=np.maximum(0,balances-exitpaid)
  alloc[:,i]=paid+exitpaid;rows.append({'m':i+1,'opening':opening.tolist(),'dividends':paid.tolist(),'exit':exitpaid.tolist(),'closing':balances.tolist(),'check':float(sum(paid)-x['dividend']+sum(exitpaid)-(ex if i==179 else 0))})
 cats=[]
 for k in range(3):
  f=[0]+[float(alloc[k,i]/COUNTS[k]-(TICKETS[k] if i==0 else 0)) for i in range(180)];value=irr(f);cats.append({'category':CATEGORIES[k],'target':float(rates[k]),'irr':value,'deviation':value-float(rates[k]) if value is not None else None})
 return {'mode':mode,'rows':rows,'categories':cats,'minimum_targets_met':all(x['irr'] is not None and x['deviation']>=-1e-8 for x in cats),'check':max(abs(x['check']) for x in rows)}

def constant_policy(result):
 """Six weights solve3target PV identities and2cash conservation identities.
 Feasible interval in one scalar; can be reproduced with ordinary Excel formulas.
 """
 p=result['p'];rates=np.array([p['target_platinum'],p['target_gold'],p['target_silver']]);div=np.array([x['dividend'] for x in result['flow']]);ex=summary(result)['exit_equity'];t=np.arange(180)/12
 F=np.array([sum(div/(1+r)**t) for r in rates]);G=ex/np.power(1+rates,179/12)
 if min(G)<=0:return {'feasible':False,'reason':'Sem saída positiva para método de pesos constantes'}
 R=F/G;den=R[1]-R[2]
 if abs(den)<1e-12:return {'feasible':False,'reason':'Metas degeneradas; recorrer à otimizaçãoLP'}
 alpha=-(R[0]-R[2])/den;beta=(sum(BASES/G)-1-R[2])/den
 a=np.array([1,alpha,-1-alpha,-R[0],-alpha*R[1],(1+alpha)*R[2]]);b=np.array([0,beta,1-beta,BASES[0]/G[0],BASES[1]/G[1]-beta*R[1],BASES[2]/G[2]-(1-beta)*R[2]])
 low=0.;high=1.;bad=False
 for aa,bb in zip(a,b):
  if abs(aa)<1e-12:bad|=bool(bb< -1e-8 or bb>1+1e-8)
  elif aa>0:low=max(low,-bb/aa);high=min(high,(1-bb)/aa)
  else:low=max(low,(1-bb)/aa);high=min(high,-bb/aa)
 feasible=not bad and low<=high+1e-10
 if not feasible:return {'feasible':False,'reason':'Sem pesos constantes que atinjam simultaneamente as metas','low':float(low),'high':float(high),'F':F.tolist(),'G':G.tolist(),'a':a.tolist(),'b':b.tolist()}
 x=low+(high-low)*p.get('allocation_preference',.5);x=min(1,max(0,x));weights=np.clip(a*x+b,0,1);rows=[];cats=[]
 for i,row in enumerate(result['flow']):rows.append({'m':i+1,'dividends':(row['dividend']*weights[:3]).tolist(),'exit':(ex*weights[3:] if i==179 else np.zeros(3)).tolist(),'check':float(row['dividend']*(sum(weights[:3])-1)+(ex*(sum(weights[3:])-1) if i==179 else 0))})
 for k in range(3):
  cf=[0]+[(row['dividends'][k]+row['exit'][k])/COUNTS[k]-(TICKETS[k] if i==0 else 0) for i,row in enumerate(rows)];value=irr(cf);cats.append({'category':CATEGORIES[k],'irr':value,'target':float(rates[k]),'deviation':value-float(rates[k]) if value is not None else None})
 return {'feasible':True,'dividend_weights':weights[:3].tolist(),'exit_weights':weights[3:].tolist(),'low':float(low),'high':float(high),'F':F.tolist(),'G':G.tolist(),'a':a.tolist(),'b':b.tolist(),'rows':rows,'categories':cats,'check':max(abs(x['check']) for x in rows)}
