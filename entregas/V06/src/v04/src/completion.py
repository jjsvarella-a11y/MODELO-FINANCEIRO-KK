"""Projected completion coverage is explicitly separated from secured evidence."""
from engine import *
def completion(result,milestone=24):
 p=result['p'];f=result['flow'];m=int(milestone);end=min(180,int(p['open']));future=f[m:end];last=f[end-1]
 # Existing cash includes protected amounts; protected opening cash is a use once.
 capex=sum(x['construction']+x['supervision']+x['contingency'] for x in future)
 charges=sum(x['interest']+x['principal']+x['openingfee']+x['broker']+x['mobil']+x['loan_interest']+x['loan_principal']+x['commitment']+x['hotpay'] for x in future)
 operating=max(0,-sum(x['hotel_cf']+x['res_cf'] for x in future));reserve=last['reserve'];cost=capex+charges+operating+reserve
 projected_cash=f[m-1]['cash'];planned_receipts=sum(x['receipts']-x['transfer']-x['marketing']-x['kyenu'] for x in future);planned_equity=sum(x['equity'] for x in future)
 # Only forecast-drawable capacity, not the30M€strategic gross limit.
 window=min(end,int(p['draw_end']));bank_ceiling=min(p['limit']*p['bank_available'],sum(x['total'] for x in result['cap'][:window])*p['bank_capex_share']) if p['pari_passu'] else p['limit']*p['bank_available']
 eligible=sum(x['equity'] for x in f[:window])*(p['subordinated_eligible'] if p['funding_kind']=='Empréstimo' else 1)+p['eligible_history']
 if eligible<sum(x['total'] for x in result['cap'][:window])*p['min_equity_ratio']:bank_ceiling=0
 projected_bank=max(0,bank_ceiling-f[m-1]['cumulative']) if p['bank_cp']==1 and m<p['draw_end'] and p['bank_start']<=p['draw_end'] else 0
 projected=projected_cash+planned_receipts+planned_equity+projected_bank
 keys=['completion_cash','completion_bank','completion_receipts','completion_equity'];known=all(p[k] is not None for k in keys)
 secured=sum(p[k] for k in keys) if known else None;ratio=secured/cost if secured is not None and cost>0 else None
 status='Pendente de Validação Externa — recursos não documentados' if not known else 'Pendente de Validação Externa — custo não validado' if not p['completion_cost_validated'] else 'Pendente de Validação Externa — calendário não validado' if not p['completion_timing_validated'] else 'Conforme — cobertura documentada no cenário' if ratio>=1 and secured>=0 else 'Alerta — défice de conclusão'
 return {'milestone':m,'cost_to_complete':capex,'charges_to_open':charges,'opening_operating_need':operating,'opening_reserve':reserve,'total_uses':cost,'projected_cash':projected_cash,'projected_receipts':planned_receipts,'projected_equity':planned_equity,'projected_bank':projected_bank,'projected_liquidity':projected,'projected_coverage':projected/cost if cost>0 else None,'projected_gap':max(0,cost-projected),'secured_liquidity':secured,'secured_coverage':ratio,'secured_gap':max(0,cost-secured) if secured is not None else None,'status':status}
