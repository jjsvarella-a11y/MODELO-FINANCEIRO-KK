"""Compare every engine formula with the published pre-correction checkpoint."""
import hashlib,json,re
from zipfile import ZipFile
from lxml import etree as E
from assemble import paths,T
from premium_ui import ROOT,WORK

def main():
 before=WORK/'checkpoints/Before_ScenarioSync.xlsx';after=ROOT/'KK01_V06_Executive_Master.xlsx';checked=0;engines=[]
 with ZipFile(before) as a,ZipFile(after) as b:
  _,pa=paths(a);_,pb=paths(b)
  for sn in pa:
   if not re.fullmatch(r'V06 Motor|V06 Referencia|V06 [SMO]\d+',sn):continue
   def formulas(z,p):
    root=E.fromstring(z.read(p));return [(c.get('r'),c.findtext(T('f')),dict(c.find(T('f')).attrib)) for c in root.iter(T('c')) if c.find(T('f')) is not None]
   old,new=formulas(a,pa[sn]),formulas(b,pb[sn]);assert old==new,sn;checked+=len(old);engines.append(sn)
 assert len(engines)==69
 result={'status':'PASS','before_sha256':hashlib.sha256(before.read_bytes()).hexdigest(),'after_sha256':hashlib.sha256(after.read_bytes()).hexdigest(),'engine_sheets':len(engines),'formula_records_unchanged':checked,'change_location':'V06 Controlos!B67 is derived from canonical operational selector; no engine formula changes versus previous published premium version','native_excel365':'NOT TESTED'}
 (ROOT/'Preservacao_Motor_SYNC.json').write_text(json.dumps(result,indent=2));print(result)
if __name__=='__main__':main()
