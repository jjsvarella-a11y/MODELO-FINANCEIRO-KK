"""Test-only numeric execution of the existing core; never embedded in dashboards."""
from functools import lru_cache
from collections import defaultdict
import importlib.util
from pathlib import Path
_spec=importlib.util.spec_from_file_location('kk_sync_core',Path(__file__).with_name('core.py'));_core=importlib.util.module_from_spec(_spec);_spec.loader.exec_module(_core)
Model,asx,TICKETS,legacy=_core.Model,_core.asx,_core.TICKETS,_core.legacy
from reference_io import load_reference
from premium_ui import ROOT
class SinkSheet:
 def __init__(self,name):self.name=name
 def __getattr__(self,name):return lambda *a,**k:None
class SinkWorkbook:
 def add_worksheet(self,name):return SinkSheet(name)
@lru_cache(None)
def reference(mode,corrected=True):
 ref=load_reference(ROOT);p=ref['EFCC']['effective'].copy();profile={'Conservative':'Conservador','Base Case':'Base'}.get(mode,'Favorável') if corrected else 'Favorável'
 p.update(hotel_mode=mode,hotel_scenario=profile)
 if mode=='Custom':p.update(adr_input=200,occ_input=.65)
 units=[{k:asx(v) for k,v in u.items()} for u in legacy.UNITS]
 model=Model(SinkWorkbook(),defaultdict(lambda:None),{k:asx(v) for k,v in p.items()},units,legacy.CHAPTERS)
 model.tickets=[(code,cat,*map(asx,[amount,share,month,1,1,0,0])) for code,cat,amount,share,month in TICKETS]
 data=model.calculate('ReferenceTest')['rows'];rows=[{k:v.v for k,v in d.items()} for d in data]
 last=int(p['exit_month']);end=rows[last-1];ev=sum(r['ebitda'] for r in rows[last-12:last])*p['exit_multiple'];exit=max(0,ev*(1-p['exit_cost'])-end['debt']+end['cash']-end['hotusa_reserve']-end['tax_payable'])
 flows=[0]+[(d['dividend']*(1900000/4150000 if d['dividend_target']>=4150000 else 1650000/3575000 if d['dividend_target']>=3575000 else 1400000/3000000)-d['investor_outlay']+(exit*.5 if m==last else 0)) if m<=last else 0 for m,d in enumerate(rows,1)]
 paid=sum(r['investor_outlay'] for r in rows[:last]);selected=[r for r in rows if r['op_age']==4 and r['available_room_nights']>0];opening=next(r for r in rows if r['available_room_nights']>0)
 metrics={'profile':profile,'adr':opening['adr'],'occupancy':opening['occupancy'],'revpar':sum(r['room_revenue'] for r in selected)/sum(r['available_room_nights'] for r in selected),'room_revenue':sum(r['room_revenue'] for r in selected),'fb_revenue':sum(r['fb_revenue'] for r in selected),'mice_revenue':sum(r['mice_revenue'] for r in selected),'gop':sum(r['gop'] for r in selected),'ebitda':sum(r['ebitda'] for r in selected),'revenue':sum(r['revenue'] for r in selected),'operator_fees':sum(r['operator_fees'] for r in selected),'ffe':sum(r['ffe'] for r in selected),'cash_end':end['cash'],'cfads_total':sum(r['cfads'] for r in rows),'dividends':sum(r['dividend'] for r in rows[:last]),'irr':legacy.irr(flows),'moic':(sum(flows)+paid)/paid,'npv':sum(cf/(1+p['discount'])**(m/12) for m,cf in enumerate(flows)),'exit':exit}
 metrics['margin']=metrics['ebitda']/metrics['revenue']
 return rows,metrics,flows
if __name__=='__main__':
 import json
 for mode in ['Base Case','Conservative','Favorable','Upside Stress','Custom']:
  _,metrics,_=reference(mode);print(json.dumps({'mode':mode,**metrics},ensure_ascii=False),flush=True)
