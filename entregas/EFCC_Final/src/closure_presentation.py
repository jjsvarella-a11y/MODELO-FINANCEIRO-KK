"""Recalculate bilingual presentation against verified financial caches.
Only private copies freeze backend cells. Original delivered formulas stay live.
"""
from zipfile import ZipFile,ZIP_DEFLATED
from lxml import etree as E
import copy,json,re,sys
from concurrent.futures import ThreadPoolExecutor,as_completed
from closure_r5 import *
from assemble import inject_caches
sys.path.insert(0,str(Path(__file__).resolve().parents[2]/'V06/src'))
from recalculate import recalculate

def chart_cache(path):
 src=Source(path);changes={};C='{http://schemas.openxmlformats.org/drawingml/2006/chart}'
 with ZipFile(path) as z:
  for part in z.namelist():
   if not part.startswith('xl/charts/chart') or not part.endswith('.xml'):continue
   if part in ['xl/charts/chart42.xml','xl/charts/chart44.xml','xl/charts/chart45.xml']:continue
   root=E.fromstring(z.read(part));dirty=False
   for ref in root.findall('.//'+C+'numRef')+root.findall('.//'+C+'strRef'):
    f=ref.findtext(C+'f');m=re.fullmatch(r"'([^']+)'!\$?([A-Z]+)\$?(\d+)(?::\$?([A-Z]+)\$?(\d+))?",f or '')
    if not m:
     named=re.fullmatch(r"'([^']+)'!(Q_Labels|Q_BaseCAP|Q_TestCAP|Q_BaseCum|Q_TestCum)",f or '')
     if not named:continue
     n=int(__import__('math').ceil((30+src.get('V06 Controlos','B38')[1])/3));letters={'Q_Labels':'A','Q_BaseCAP':'B','Q_TestCAP':'C','Q_BaseCum':'D','Q_TestCum':'E'};letter=letters[named[2]];m=re.fullmatch(r"'([^']+)'!\$?([A-Z]+)\$?(\d+)(?::\$?([A-Z]+)\$?(\d+))?","'"+named[1]+"'!"+letter+'46:'+letter+str(45+n))
    sn,c1,r1,c2,r2=m.groups();r1=int(r1);r2=int(r2 or r1);c2=c2 or c1
    if c1!=c2:continue
    values=[src.get(sn,c1+str(r))[1] for r in range(r1,r2+1)];numeric=E.QName(ref).localname=='numRef';cache=ref.find(C+('numCache' if numeric else 'strCache'))
    if cache is not None:ref.remove(cache)
    cache=E.SubElement(ref,C+('numCache' if numeric else 'strCache'))
    if numeric:E.SubElement(cache,C+'formatCode').text='General'
    E.SubElement(cache,C+'ptCount',val=str(len(values)))
    for i,value in enumerate(values):
     if value is None:continue
     if numeric and not isinstance(value,(float,int)):continue
     p=E.SubElement(cache,C+'pt',idx=str(i));E.SubElement(p,C+'v').text=str(value)
    dirty=True
   if dirty:changes[part]=xml(root)
  temp=path.with_suffix('.charts.xlsx')
  with ZipFile(temp,'w',ZIP_DEFLATED,compresslevel=6) as out:
   for item in z.infolist():out.writestr(copy.copy(item),changes.get(item.filename,z.read(item.filename)))
 temp.replace(path)

def lab(language):
 dest=WORK/('Presentation_'+language+'.xlsx')
 with ZipFile(MASTER) as z,ZipFile(dest,'w',ZIP_DEFLATED,compresslevel=3) as out:
  _,pp=paths(z);lookup={p:n for n,p in pp.items()}
  for item in z.infolist():
   data=z.read(item.filename);sn=lookup.get(item.filename)
   freeze=sn=='V06 Motor' or sn=='V06 Referencia' or bool(re.fullmatch(r'V06 [SMOF]\d+',sn or '')) or sn in ['V06 Cenarios','V06 Fluxos Invest','V06 Blocos']
   if freeze or sn=='DASH 10 Centro':
    root=E.fromstring(data)
    if freeze:
     for c in root.iter(T('c')):
      f=c.find(T('f'))
      if f is not None:c.remove(f)
    if sn=='DASH 10 Centro':
     c=root.find('.//'+T('c')+'[@r="H7"]')
     for n in list(c):c.remove(n)
     c.set('t','inlineStr');E.SubElement(E.SubElement(c,T('is')),T('t')).text=language
    data=xml(root)
   out.writestr(copy.copy(item),data)
 return recalculate(dest,WORK/('presentation_'+language))

def main():
 results={}
 with ThreadPoolExecutor(max_workers=2) as ex:
  for f in as_completed({ex.submit(lab,x):x for x in ['PT','EN']}):
   file=f.result();language='PT' if '_PT' in file.name else 'EN';chart_cache(file);results[language]=str(file)
 temp=WORK/'Presentation_Cached.xlsx';inject_caches(MASTER,results['PT'],temp);temp.replace(MASTER);chart_cache(MASTER);(WORK/'Presentations.json').write_text(json.dumps(results,indent=2));print('Bilingual presentation caches and chart values refreshed',flush=True)
if __name__=='__main__':main()
