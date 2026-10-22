"""Create flat private visual-review workbooks; delivered workbook remains live."""
from pathlib import Path
from zipfile import ZipFile,ZIP_DEFLATED
from lxml import etree as E
import json,copy
import fitz
from concurrent.futures import ThreadPoolExecutor,as_completed
from assemble import paths,T,xml
from premium_refine import MASTER,OUT,WORK
from premium_ui import DASHES
from export_all import export

def viewport(source,language):
 target=WORK/('Visual_'+language+'.xlsx')
 with ZipFile(source) as z,ZipFile(target,'w',ZIP_DEFLATED,compresslevel=3) as result:
  wb,pp=paths(z);dn=wb.find(T('definedNames'));snames=[s.get('name') for s in wb.find(T('sheets'))]
  for name in dn:
   if name.get('name')=='_xlnm.Print_Area' and name.get('localSheetId') is not None:
    sn=snames[int(name.get('localSheetId'))]
    if sn in [s[0] for s in DASHES]:name.text=f"'{sn}'!$B$1:$Y${147 if sn=='DASH 05 V06' else 117 if sn!='DASH 10 Centro' else 85}"
  for item in z.infolist():
   data=z.read(item.filename)
   if item.filename in pp.values():
    root=E.fromstring(data)
    for c in root.iter(T('c')):
     formula=c.find(T('f'))
     if formula is not None:c.remove(formula)
    data=xml(root)
   elif item.filename=='xl/workbook.xml':data=xml(wb)
   result.writestr(copy.copy(item),data)
 return target

def main():
 sources={'PT':MASTER,'EN':WORK/'tests/mutations/Filter_44_1/recalc/out/Filter_44_1.xlsx'}
 # EN source is supplied explicitly by the runner after real translation recalc.
 import sys
 if len(sys.argv)>1:sources['EN']=Path(sys.argv[1])
 sources={k:v for k,v in sources.items() if v.exists()}
 def task(item):
  lang,source=item;flat=viewport(source,lang);names=[x[0] for x in DASHES];pdf=export(flat,names,'Views_'+lang,WORK/('views_'+lang));doc=fitz.open(pdf);assert len(doc)==10,len(doc);proof=[]
  directory=WORK/('png_'+lang);directory.mkdir(exist_ok=True)
  for i,page in enumerate(doc):
   p=directory/f'DASH_{i+1:02}_{lang}.png';page.get_pixmap(matrix=fitz.Matrix(1.5,1.5)).save(p);proof.append({'dashboard':i+1,'language':lang,'page':i+1,'image':str(p),'page_width':page.rect.width,'page_height':page.rect.height,'errors':any(e in page.get_text() for e in ['#VALUE!','#REF!','#NAME?','#N/A'])})
  print(lang,'10 views generated',flush=True);return proof
 results=[]
 with ThreadPoolExecutor(max_workers=2) as pool:
  for future in as_completed([pool.submit(task,x) for x in sources.items()]):results+=future.result()
 (WORK/'Visual_Previews.json').write_text(json.dumps(results,indent=2,ensure_ascii=False))
if __name__=='__main__':main()
