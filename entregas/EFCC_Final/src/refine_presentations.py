"""Recalculate PT/EN presentation from verified monthly caches, preserving delivery formulas."""
from zipfile import ZipFile,ZIP_DEFLATED
from pathlib import Path
from lxml import etree as E
from concurrent.futures import ThreadPoolExecutor
import copy,sys,json
from assemble import paths,T,xml,inject_caches
from premium_refine import MASTER,OUT,WORK,DASHES
sys.path.insert(0,str(OUT.parents[1]/'V06/src'))
from recalculate import recalculate
LIVE={x[0] for x in DASHES}|{'V06 Visual','DASH 10 V06','V06 Brief','V06 Controlos','V06 Hipoteses','V06 Traducoes','V06 Registo Ativo'}|{f'V06 PDF {i:02}' for i in range(1,10)}
def make(language):
 target=WORK/('Presentation_'+language+'.xlsx')
 with ZipFile(MASTER) as z,ZipFile(target,'w',ZIP_DEFLATED,compresslevel=3) as out:
  _,pp=paths(z);snames={p:sn for sn,p in pp.items()}
  for item in z.infolist():
   data=z.read(item.filename);sn=snames.get(item.filename)
   if sn is not None:
    root=E.fromstring(data)
    if sn not in LIVE:
     for c in root.iter(T('c')):
      f=c.find(T('f'))
      if f is not None:c.remove(f)
    if sn=='DASH 10 Centro':
     c=root.find('.//'+T('c')+'[@r="H7"]')
     for n in list(c):c.remove(n)
     c.set('t','inlineStr');E.SubElement(E.SubElement(c,T('is')),T('t')).text=language
    data=xml(root)
   out.writestr(copy.copy(item),data)
 return recalculate(target,WORK/('presentation_'+language))
def main():
 proof=json.loads((OUT/'Validacao_Financeira.json').read_text());assert proof['status']=='PASS'
 temp=MASTER.with_suffix('.cached.xlsx');count=inject_caches(MASTER,WORK/'full_recalc/out'/MASTER.name,temp);temp.replace(MASTER)
 with ThreadPoolExecutor(max_workers=2) as pool:results=dict(zip(['PT','EN'],pool.map(make,['PT','EN'])))
 temp=MASTER.with_suffix('.presentation.xlsx');inject_caches(MASTER,results['PT'],temp);temp.replace(MASTER)
 (WORK/'Presentations.json').write_text(json.dumps({k:str(v) for k,v in results.items()},indent=2));print('Verified financial caches',count,'; PT/EN presentation recalculated; original delivery formulas retained')
if __name__=='__main__':main()
