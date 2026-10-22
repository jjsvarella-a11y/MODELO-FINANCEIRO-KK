import sys,json,shutil
from pathlib import Path
from concurrent.futures import ThreadPoolExecutor
sys.path.insert(0,str(Path(__file__).resolve().parent))
from export_all import export
from premium_refine import MASTER,OUT,WORK
import fitz

def run(lang):
 source=WORK/f'presentation_{lang}/out/Presentation_{lang}.xlsx'
 dest=OUT/'PDF'/lang;dest.mkdir(parents=True,exist_ok=True);proof=[]
 for names,label,count in [([f'V06 PDF {i:02}' for i in range(1,10)],'Dashboards',9),(['DASH 10 V06'],'DASH_10',10),(['V06 Brief'],'Executive_Brief',2)]:
  # Technical print names are obtained from workbook, not guessed.
  if label=='Dashboards':
   from zipfile import ZipFile
   from assemble import paths
   with ZipFile(source) as z:
    _,pp=paths(z);names=[n for n in pp if n.startswith('V06 PDF')]
   assert len(names)==9,names
  pdf=export(source,names,label,WORK/f'final_pdf_{lang}'/label);doc=fitz.open(pdf);assert len(doc)==count,(label,len(doc))
  if label=='Dashboards':
   for i in range(9):
    part=fitz.open();part.insert_pdf(doc,from_page=i,to_page=i);part.save(dest/f'DASH_{i+1:02}_{lang}.pdf')
  else:shutil.copy2(pdf,dest/f'{label}_{lang}.pdf')
  for i,p in enumerate(doc):
   text=p.get_text();assert not any(e in text for e in ['#VALUE!','#REF!','#N/A','#NAME?']), (label,i)
   proof.append({'language':lang,'document':label,'page':i+1,'width':p.rect.width,'height':p.rect.height,'errors':False})
  print(lang,label,len(doc),flush=True)
 return proof
with ThreadPoolExecutor(max_workers=2) as pool: result=list(pool.map(run,['PT','EN']))
(OUT/'PDF_Validacao.json').write_text(json.dumps({'platform':'LibreOffice PDF export','native_excel365':'NOT TESTED','pages':sum(result,[])},indent=2))
