"""Financial acceptance for independently editable Base and Comparative contexts.
Private labs freeze unrelated historic sensitivity engines. Public file retains all.
"""
from scenario_v07 import *
from test_closure_r5 import mutate
from concurrent.futures import ThreadPoolExecutor,as_completed

def lab(source,target):
 s=Source(source);changes={}
 for sn,p in s.pp.items():
  freeze=sn=='V06 Referencia' or bool(re.fullmatch(r'V06 [SMOF]\d+',sn)) and sn!='V06 F01'
  if not freeze and sn!='V06 Cenarios':continue
  root=E.fromstring(s.z.read(p))
  for row in root.find(T('sheetData')):
   if sn=='V06 Cenarios' and row.get('r')=='5':continue
   for c in row:
    f=c.find(T('f'))
    if f is not None:c.remove(f)
  changes[p]=xml(root)
 rewrite(source,target,changes);return target

def edits_for(mode,key,selection,custom=None):
 sn,addr=PANELREF.get((key,mode),(store(mode),'B'+str(ROWS[key])));out=[(sn,addr,selection)]
 if custom is not None:out.append((store(mode),'C'+str(ROWS[key]),custom))
 return out

def run(name,edits,expected,source,initial):
 folder=WORK7/'Tests'/name;folder.mkdir(parents=True,exist_ok=True);file=folder/'Input.xlsx';mutate(source,file,edits);calculated=recalculate(file,folder/'Recalc');data=evaluate7(calculated);base=data['contexts']['Base'];cmp=data['contexts']['Comparative'];oldb=initial['contexts']['Base'];oldc=initial['contexts']['Comparative'];s=Source(calculated)
 if expected in ['Base','Comparative']:
  changed=data['contexts'][expected];unchanged=data['contexts']['Comparative' if expected=='Base' else 'Base'];original=oldc if expected=='Base' else oldb
  for key,v in original['metrics'].items():near(unchanged['metrics'][key],v)
  assert abs(changed['metrics']['Investor NPV']-oldb['metrics']['Investor NPV'])>100,(name,changed['metrics'])
 if name=='Discount':
  near(base['metrics']['Investor IRR'],oldb['metrics']['Investor IRR']);near(base['metrics']['Planned Dividends'],oldb['metrics']['Planned Dividends']);assert base['metrics']['Investor NPV']>oldb['metrics']['Investor NPV']
 if name=='Custom':near(base['params']['effective_tan'],.0575);near(base['effective_inputs']['adr'],207)
 if name=='Operator_Profile':near(base['effective_inputs']['basefee'],.025);near(base['effective_inputs']['systems'],.03);near(base['effective_inputs']['incentive'],.06)
 if name=='Operating_Profile':near(base['effective_inputs']['adr'],160);near(base['effective_inputs']['occ'],.52)
 if name=='Comparative_View':
  for rr in range(5,185):
   for cc in ['T','U','V','AH','CK','DG','DE']:near(s.get('V06 Motor',cc+str(rr))[1],s.get(MAPS['Comparative']['V06 Motor'],cc+str(rr))[1])
 for rr in range(5,14):
  for colname in ['B','C']:near(s.get(IMP,colname+str(rr))[1],s.get(MET7,colname+str(rr+1))[1])
 for sn,*_ in DASHES:
  near(s.get(sn,'T30')[1],base['metrics']['Investor NPV']);near(s.get(sn,'V30')[1],cmp['metrics']['Investor NPV'])
 return {'test':name,'status':'PASS','Base':base['metrics'],'Comparative':cmp['metrics'],'checks':{mode:d['checks'] for mode,d in data['contexts'].items()},'platform':'LibreOffice; private lab; historic branches frozen'}

def main():
 initial=evaluate7(FINAL7);old=Source(BM);new=Source(FINAL7);count=0;maximum=0
 for rr in range(5,185):
  for cc in range(147):
   addr=col(cc)+str(rr);x=old.get('V06 Motor',addr)[1]
   for sn in ['V06 Motor',MAPS['Base']['V06 Motor'],MAPS['Comparative']['V06 Motor']]:
    y=new.get(sn,addr)[1]
    if isinstance(x,(int,float)) and isinstance(y,(int,float)):near(y,x,.005);count+=1;maximum=max(maximum,abs(x-y))
 baseline={'status':'PASS','numeric_cells_compared':count,'max_difference':maximum,'Base':initial['contexts']['Base']['metrics']};(WORK7/'Baseline_Reconciliation.json').write_text(json.dumps(baseline,indent=2));print('Baseline',baseline,flush=True)
 private=lab(FINAL7,WORK7/'Private_Lab.xlsx');cases=[('Base_TAN',edits_for('Base','bank_tan','5.0%'),'Base'),('Comparative_TAN',edits_for('Comparative','bank_tan','7.0%'),'Comparative'),('Base_ADR',edits_for('Base','adr','200'),'Base'),('Comparative_ADR',edits_for('Comparative','adr','225'),'Comparative'),('Custom',edits_for('Base','bank_tan','Custom',.0575)+edits_for('Base','adr','Custom',207),'Base'),('Operator_Profile',edits_for('Base',IRO[('DASH 05 V06','H32')],'KK Target Terms'),'Base'),('Operating_Profile',edits_for('Base',IRO[('DASH 05 V06','H23')],'Conservative'),'Base'),('Discount',edits_for('Base','discount','8%'),'Base'),('Comparative_View',edits_for('Comparative','adr','225')+[('EXEC V06','H23','Comparative')],'Comparative'),('Base_Return',edits_for('Base','bank_tan','6.0%')+edits_for('Base','adr','Profile'),None)]
 tests=[]
 # Independent calculation processes, no agents. Two labs bound memory use.
 with ThreadPoolExecutor(max_workers=2) as ex:
  fs=[ex.submit(run,name,edits,expected,private,initial) for name,edits,expected in cases]
  for f in as_completed(fs):r=f.result();tests.append(r);print(r['test'],'PASS',flush=True)
 result={'status':'PASS','baseline':baseline,'tests':tests,'native_excel365':'NOT TESTED'};(OUT7/'Testes_Financeiros.json').write_text(json.dumps(result,indent=2));print('ALL PASS',flush=True)
if __name__=='__main__':main()
