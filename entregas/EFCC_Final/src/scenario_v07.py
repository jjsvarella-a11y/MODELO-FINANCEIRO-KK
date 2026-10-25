"""Versioned scenario operations without mandatory VBA. Never overwrite history.
Usage: python scenario_v07.py save-base|save-comparative|restore [--file ...]
Save means: recalculate ALL formulas, validate 180 months, retain prior version,
commit the effective inputs and record an immutable parameter/result snapshot.
"""
from editable_base_v07 import *
from test_closure_r5 import recalculate,near
from assemble import inject_caches
import argparse,math,shutil
from bank_domains import METRIC_KEYS
VERSIONS=OUT7/'Historico';VERSIONS.mkdir(exist_ok=True)
def sha(p):return hashlib.sha256(Path(p).read_bytes()).hexdigest()
def values(source,sn):
 out={}
 for addr in source.cells(sn):out[addr]=source.get(sn,addr)[1]
 return out

def evaluate7(file,scan=True):
 source=Source(file);errors=[]
 if scan:
  for sn,p in source.pp.items():
   data=source.z.read(p)
   if b't="e"' in data or b"t='e'" in data:
    for c in E.fromstring(data).iter(T('c')):
     if c.get('t')=='e':errors.append([sn,c.get('r'),c.findtext(T('v'))])
 if errors:raise ValueError('Formula errors: '+repr(errors[:30]))
 result={'source_sha256':sha(file),'model':'V07','calculation_platform':'LibreOffice full recalculation','native_excel365':'NOT TESTED','contexts':{},'formula_errors':errors}
 for mode,mp in MAPS.items():
  cells=values(source,mp['V06 Motor']);p={key:source.get(mp[DAT],'B'+str(P[key]))[1] for key,*_ in PARAMS};p['effective_tan']=source.get(mp[DAT],'B42')[1];effective={d['key']:source.get(store(mode),'D'+str(ROWS[d['key']]))[1] for d in DEFS}
  if any(not isinstance(p[key],(int,float)) or not math.isfinite(p[key]) for key in ['tan','grace','term','facility','draw_start','draw_end','period']):raise ValueError('Missing numeric bank input')
  if p['term']<=0 or p['period']<=0 or p['draw_start']>p['draw_end'] or p['draw_end']>180 or p['draw_end']>p['term']+p['grace']:raise ValueError('Invalid financing horizon')
  if p['status']=='Approved' and p['evidence'] in ['',None,'Not supplied','Not saved']:raise ValueError('Approved terms require documentary evidence')
  if not 0<=effective['occ']<=1 or effective['adr']<=0 or effective['discount']<=-1:raise ValueError('Invalid operating/discount parameters')
  monthly=[];checks={k:0. for k in ['cash_check','debt_check','ffe_fund_check','inventory_check']}
  for rr in range(5,185):
   row={key:cells.get(col(index)+str(rr)) for key,index in CC.items()};row.update(interest_due=cells['EK'+str(rr)],capitalized_interest=cells['EL'+str(rr)],early_fee=cells['EM'+str(rr)],closing_fee=cells['EN'+str(rr)],all_fees=cells['EO'+str(rr)],contractual_service=cells['EP'+str(rr)],bank_payments=cells['EQ'+str(rr)])
   for key in checks:near(row[key],0,.005);checks[key]=max(checks[key],abs(row[key]))
   near(row['debt_open']+row['draw']+row['capitalized_interest']-row['principal']-row['sweep'],row['debt'],.005)
   if row['cumulative_draw']>p['facility']+.005:raise ValueError('Facility exceeded')
   near(row['all_fees'],row['commitment_fee']+row['bank_fees']+row['early_fee']+row['closing_fee'],.005)
   near(row['bank_payments'],row['interest']+row['principal']+row['sweep']+row['all_fees'],.005)
   row['headroom']=row['cash']-row['reserve'];monthly.append(row)
  cc='B' if mode=='Base' else 'C';metrics={source.get(MET7,'A'+str(r))[1]:source.get(MET7,cc+str(r))[1] for r in range(6,20)}
  cf=[source.get(mp['V06 F01'],'B'+str(r))[1] or 0 for r in range(5,186)];npv=sum(v/(1+effective['discount'])**(m/12) for m,v in enumerate(cf));near(npv,metrics['Investor NPV'])
  bankmetrics={METRIC_KEYS[r-210]:source.get(mp[SCH],'B'+str(r))[1] for r in range(210,227)}
  result['contexts'][mode]={'bank_metrics':bankmetrics,'id':source.get(STATUS7,'B5' if mode=='Base' else 'B6')[1],'params':p,'effective_inputs':effective,'metrics':metrics,'monthly':monthly,'investor_cashflows':cf,'checks':checks,'saved':all(all(source.get(store(mode),a+str(r))[1]==source.get(store(mode),b+str(r))[1] for a,b in [('D','E'),('B','J'),('C','K')]) for r in range(6,len(DEFS)+6))}
 return result

def patch_cells(file,edits,target=None):
 target=Path(target or file);source=Source(file);changes={};roots={}
 for sn,addr,value in edits:
  if sn not in roots:roots[sn]=Sheet(E.fromstring(source.z.read(source.pp[sn])),{})
  roots[sn].put(addr,value)
 for sn,s in roots.items():
  dimension=s.r.find(T('dimension'));maximum_row=max(rc(a)[0]+1 for a in s.cells);maximum_col=max(rc(a)[1] for a in s.cells)
  if dimension is not None:dimension.set('ref','A1:'+col(maximum_col)+str(maximum_row))
  changes[source.pp[sn]]=s.finish()
 tmp=WORK7/('Mutation_'+datetime.now().strftime('%H%M%S%f')+'.xlsx');rewrite(file,tmp,changes);shutil.move(tmp,target)
 return target

def full_recalc(file,folder):
 calculated=recalculate(Path(file),Path(folder));data=evaluate7(calculated);cached=WORK7/('Cached_'+datetime.now().strftime('%H%M%S%f')+'.xlsx');inject_caches(file,calculated,cached);shutil.move(cached,file);data['_validated_file_sha256']=sha(file);data['calculation_file_sha256']=data['source_sha256'];data['source_sha256']=sha(file);data['calculated_at']=datetime.now(timezone.utc).isoformat();return data

def selection_location(source,mode,key):
 # V08 additionally exposes existing Base inputs directly on their dashboards.
 # Discover the canonical location from its simple B-column reference so old
 # versions remain compatible and restore never overwrites a binding formula.
 f,_=source.get(store(mode),'B'+str(ROWS[key]))
 match=re.fullmatch(r"'([^']+)'!(\$?[A-Z]+\$?\d+)",f or '')
 if match and match[1] in [x[0] for x in DASHES]:return match[1],match[2].replace('$','')
 return PANELREF.get((key,mode),(store(mode),'B'+str(ROWS[key])))

def raw_snapshot(source,mode):
 data={}
 for d in DEFS:
  key=d['key'];rr=ROWS[key];panel=selection_location(source,mode,key);data[key]={'selection':source.get(*panel)[1],'custom':source.get(store(mode),'C'+str(rr))[1]}
 return data

def save(file,mode,description='',data=None):
 file=Path(file);before=Source(file);priorid=before.get(STATUS7,'B5' if mode=='Base' else 'B6')[1];priorvalues={d['key']:before.get(store(mode),'E'+str(ROWS[d['key']]))[1] for d in DEFS};previousraw={d['key']:{'selection':before.get(store(mode),'J'+str(ROWS[d['key']]))[1],'custom':before.get(store(mode),'K'+str(ROWS[d['key']]))[1]} for d in DEFS};before.z.close()
 data=data if data is not None and data.get('_validated_file_sha256')==sha(file) else full_recalc(file,WORK7/('Save_'+datetime.now().strftime('%H%M%S%f')));source=Source(file);ctx=data['contexts'][mode];timestamp=datetime.now(timezone.utc).isoformat();identifier=('BASE-' if mode=='Base' else 'CMP-')+datetime.now(timezone.utc).strftime('%Y%m%dT%H%M%S%fZ');path=VERSIONS/(identifier+'.json')
 ctx['id']=identifier;raw=raw_snapshot(source,mode);snapshot={'id':identifier,'previous_id':priorid,'classification':mode,'saved_at':timestamp,'description':description,'raw_inputs':raw,'context':ctx,'current_base_id':data['contexts']['Base']['id'],'source_sha256':sha(file),'native_excel365':'NOT TESTED'}
 # Seed the previously imported Base if not already registered; never lose it.
 priorpath=VERSIONS/(str(priorid)+'.json')
 if not priorpath.exists():
  seed={d['key']:{'selection':initial_choice(d),'custom':d['initial']} for d in DEFS};priorpath.write_text(json.dumps({'id':priorid,'classification':mode,'description':'Initial adopted R6.1 / R4 historical import','raw_inputs':seed,'effective_inputs':priorvalues},ensure_ascii=False,indent=2))
 path.write_text(json.dumps(snapshot,ensure_ascii=False,indent=2));edits=[(STATUS7,'B5' if mode=='Base' else 'B6',identifier),(STATUS7,'B8',timestamp)]
 for d in DEFS:
  rr=ROWS[d['key']];edits.extend([(store(mode),'E'+str(rr),ctx['effective_inputs'][d['key']]),(store(mode),'J'+str(rr),raw[d['key']]['selection']),(store(mode),'K'+str(rr),raw[d['key']]['custom'])])
 hist=values(source,HIST7);last=max([int(re.sub('[A-Z]','',a)) for a in hist if a.startswith('A')]+[5]);changed=[(key,priorvalues[key],v) for key,v in ctx['effective_inputs'].items() if v!=priorvalues[key]]
 for key,entry in raw.items():
  for field in ['selection','custom']:
   if entry[field]!=previousraw[key][field]:changed.append((key+'.'+field,previousraw[key][field],entry[field]))
 changed=changed or [('(saved unchanged)','','')]
 for index,(key,priorvalue,newvalue) in enumerate(changed,last+1):
  row=[identifier,mode,timestamp,key,priorvalue,newvalue,ctx['metrics']['Investor NPV'],ctx['metrics']['Investor IRR'],description,sha(path)]
  edits +=[(HIST7,col(j)+str(index),v) for j,v in enumerate(row)]
 saved_ids=[v for a,v in values(source,STATUS7).items() if a.startswith('D') and rc(a)[1]==3 and rc(a)[0]>=19 and v];edits.append((STATUS7,'D'+str(len(saved_ids)+20),identifier))
 source.z.close();patch_cells(file,edits)
 # Only the save-status caches changed; financial caches came from full_recalc.
 s=Source(file);sn=STATUS7;sh=Sheet(E.fromstring(s.z.read(s.pp[sn])),{});address='B9' if mode=='Base' else 'B10';formula=sh.cells[address].findtext(T('f'));sh.put(address,'Saved',formula);tmp=WORK7/'Saved.xlsx';rewrite(file,tmp,{s.pp[sn]:sh.finish()});s.z.close();shutil.move(tmp,file)
 ctx['saved']=True;data['_validated_file_sha256']=sha(file);data['source_sha256']=sha(file)
 return identifier,data

def restore(file,identifier,mode=None):
 path=VERSIONS/(identifier+'.json');snapshot=json.loads(path.read_text(encoding='utf-8-sig'));mode=mode or snapshot['classification'];edits=[];source=Source(file)
 for key,raw in snapshot['raw_inputs'].items():
  rr=ROWS[key];sn,addr=selection_location(source,mode,key)
  if key not in ALIASES:edits.append((sn,addr,raw['selection']))
  edits.append((store(mode),'C'+str(rr),raw['custom']))
 source.z.close();patch_cells(file,edits);return save(file,mode,'Restored '+identifier)
if __name__=='__main__':
 p=argparse.ArgumentParser();p.add_argument('action',choices=['save-base','save-comparative','restore','validate']);p.add_argument('--file',type=Path,default=FINAL7);p.add_argument('--id');p.add_argument('--description',default='');a=p.parse_args()
 if a.action=='restore':print(restore(a.file,a.id)[0])
 elif a.action=='validate':print(json.dumps(full_recalc(a.file,WORK7/'Validate'),indent=2))
 else:print(save(a.file,'Base' if a.action=='save-base' else 'Comparative',a.description)[0])
