"""Compare restoration with the published workbook, not only the historic model."""
import json,re,hashlib
from zipfile import ZipFile
from lxml import etree as E
from premium_ui import Source,ROOT,WORK,T
from assemble import paths

def verify():
 old=WORK/'checkpoints/Before_Navigation_Restoration.xlsx';new=ROOT/'KK01_V06_Executive_Master.xlsx';count=0;changed=[]
 with ZipFile(old) as a,ZipFile(new) as b:
  _,ap=paths(a);_,bp=paths(b)
  engines=[n for n in ap if re.fullmatch(r'V06 Motor|V06 Referencia|V06 [SMO]\d+',n)]
  def values(data):
   r=E.fromstring(data);return {c.get('r'):(c.findtext(T('f')),c.get('t'),c.findtext(T('v'))) for c in r.iter(T('c'))}
  for sn in engines:
   left,right=values(a.read(ap[sn])),values(b.read(bp[sn]));count+=sum(1 for x in left.values() if x[0] is not None)
   if left!=right:changed.append(sn)
 assert len(engines)==69 and not changed,changed
 previous=Source(old);current=Source(new);inputs=json.loads((ROOT/'Parametros_Interface.json').read_text());controls=[]
 for p in inputs:
  sn,addr=p['sheet'],p['cell'];assert previous.get(sn,addr)==current.get(sn,addr),(sn,addr);controls.append(sn+'!'+addr)
  for source in [previous,current]:
   r=E.fromstring(source.z.read(source.pp[sn]));rules=[E.tostring(x) for x in r.findall(T('dataValidations')+'/'+T('dataValidation')) if x.get('sqref')==addr]
   if source is previous:oldrules=rules
   else:assert rules==oldrules,(sn,addr,'validation changed')
 proof={'status':'PASS','source_sha256':hashlib.sha256(old.read_bytes()).hexdigest(),'result_sha256':hashlib.sha256(new.read_bytes()).hexdigest(),'financial_engines_compared':len(engines),'engine_formulas_compared':count,'all_engine_formulas_and_cached_values_identical':True,'editable_inputs_and_validation_rules_unchanged':len(controls),'changes':['Menu opens the dashboard menu rather than report chapters','Ten frozen numbered routes on every dashboard','Hotel annual KPI includes RevPAR, GOP retained in operating table','Decision report identifies developer, investors, lenders, buyers and operator'],'native_excel365':'NOT TESTED','official_logo':'PENDING insertion of original image; existing identity area retained'}
 (ROOT/'Restauro_Mandato.json').write_text(json.dumps(proof,indent=2,ensure_ascii=False));print(proof)
if __name__=='__main__':verify()
