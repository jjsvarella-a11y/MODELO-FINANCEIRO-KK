"""Check all 69 recalculated engines, allowing the nine exact-axis variants."""
import sys,json,gzip,hashlib,re
from pathlib import Path
from collections import defaultdict
from zipfile import ZipFile
from lxml import etree as E
from xlsxwriter.utility import xl_col_to_name as col
import scenario_reference as sr
import validate_final as vf
from assemble import paths,T
from premium_refine import ROOT,OUT,WORK,BASE,MASTER,H

def main():
 calc=WORK/'full_recalc/out'/MASTER.name
 with gzip.open(ROOT/'Referencia_Python_V06_SYNC.json.gz','rt') as f:ref=json.load(f)
 params=ref['EFCC']['effective'].copy();model=sr.Model(sr.SinkWorkbook(),defaultdict(lambda:None),{k:sr.asx(v) for k,v in params.items()},[{k:sr.asx(v) for k,v in u.items()} for u in sr.legacy.UNITS],sr.legacy.CHAPTERS);model.tickets=[(code,cat,*map(sr.asx,[amount,share,month,1,1,0,0])) for code,cat,amount,share,month in sr.TICKETS]
 labels=list(ref['scenarios']);new=[]
 for ri,occ in enumerate([.50,.58,.60,.65,.70,.75]):
  for ci,adr in enumerate([150,175,200,225,185]):
   if ri!=1 and ci!=4:continue
   label=labels[25+ri*5+ci];data=model.calculate('NumericExactAxis',{'hotel_mode':'Custom','adr_input':adr,'occ_input':occ})['rows'];ref['scenarios'][label]=[{k:v.v for k,v in row.items()} for row in data];new.append({'sheet':f'V06 M{ri*5+ci+1:02}','adr':adr,'occupancy':occ})
 vf.ROOT=OUT;vf.load_reference=lambda _:ref;proof=vf.validate(calc);proof.update(exact_axis_cases=new,reference='Original existing-core numeric reference; only nine reused matrix variants recalculated for exact current axes')
 with ZipFile(BASE) as old,ZipFile(MASTER) as newzip:
  _,po=paths(old);_,pn=paths(newzip);count=0
  for sn,p in po.items():
   if not re.fullmatch(r'V06 Motor|V06 Referencia|V06 [SMO]\d+|V06 F\d+',sn):continue
   def formulas(z,p):return [(c.get('r'),c.findtext(T('f'))) for c in E.fromstring(z.read(p)).iter(T('c')) if c.find(T('f')) is not None]
   a,b=formulas(old,p),formulas(newzip,pn[sn]);assert a==b,sn;count+=len(a)
 proof['original_engine_and_cashflow_formulas_unchanged']=count;proof['financial_model_rebuilt']=False
 (OUT/'Validacao_Financeira.json').write_text(json.dumps(proof,indent=2,ensure_ascii=False));print('R2 financial reference PASS; exact axes and original formulas verified')
if __name__=='__main__':main()
