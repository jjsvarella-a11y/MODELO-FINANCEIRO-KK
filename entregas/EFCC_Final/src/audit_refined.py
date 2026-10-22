"""Compare delivered financial expressions and resolve native navigation destinations."""
from zipfile import ZipFile
from lxml import etree as E
import json,re,csv,hashlib
from assemble import paths,T
from premium_refine import MASTER,OUT,WORK

def main():
 rows=[];count=0
 with ZipFile(WORK/'BASE_R1.xlsx') as baseline,ZipFile(MASTER) as current:
  _,bp=paths(baseline);_,cp=paths(current);assert list(bp)==list(cp)
  for sn,path in cp.items():
   root=E.fromstring(current.read(path))
   if sn=='V06 Motor' or re.fullmatch(r'V06 [SMO]\d+',sn) or sn.startswith('V06 F') or sn=='V06 Ref Motor':
    def formulas(z,p):return {c.get('r'):c.find(T('f')).text for c in E.fromstring(z.read(p)).iter(T('c')) if c.find(T('f')) is not None}
    assert formulas(baseline,bp[sn])==formulas(current,path),sn;count+=1
   for c in root.iter(T('c')):
    f=c.find(T('f'))
    if f is not None and f.text and 'HYPERLINK(' in f.text:
     m=re.search(r'HYPERLINK\("#\x27([^\x27]+)\x27!([^\"]+)"',f.text)
     if m:
      assert m[1] in cp,(sn,c.get('r'),m[1]);rows.append([sn,c.get('r'),m[1],m[2],'native HYPERLINK','PASS (structural; native click NOT TESTED)'])
  assert len(cp)==190
 with (OUT/'Matriz_Navegacao.csv').open('w',newline='') as f:
  w=csv.writer(f);w.writerow(['origin','cell','destination','anchor','type','result']);w.writerows(rows)
 result={'status':'PASS','sheets':len(cp),'financial_formula_sheets_compared':count,'native_links_resolved':len(rows),'native_excel365':'NOT TESTED','sha256':hashlib.sha256(MASTER.read_bytes()).hexdigest()}
 (OUT/'Auditoria_Preservacao.json').write_text(json.dumps(result,indent=2));print(result)
if __name__=='__main__':main()
