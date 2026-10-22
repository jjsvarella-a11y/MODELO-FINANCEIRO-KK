"""Export native print compositions with LibreOffice, with page/paper assertions."""
from pathlib import Path
from concurrent.futures import ThreadPoolExecutor,as_completed
import sys,json,fitz,shutil
from export_all import export
from premium_ui import ROOT,WORK

def all_pdfs():
 sources={'PT':WORK/'premium_final_recalc/out/KK01_V06_Executive_Master.xlsx','EN':WORK/'mutations/Premium_EN/recalc/out/Premium_EN.xlsx'}
 jobs=[]
 for lang,source in sources.items():
  jobs += [(lang,source,[f'V06 PDF {i:02}' for i in range(1,10)],'Dashboards',9),(lang,source,['DASH 10 V06'],'DASH_10_Integral',10),(lang,source,['V06 Brief'],'Executive_Brief',2)]
 def task(job):
  lang,source,targets,label,count=job;file=export(source,targets,label,WORK/f'premium_pdf_{lang}'/label);doc=fitz.open(file);assert len(doc)==count,(lang,label,len(doc),count);folder=ROOT/'PDF'/lang;folder.mkdir(exist_ok=True);proof=[]
  for i,page in enumerate(doc):
   a3=label=='Dashboards';assert abs(page.rect.width-(1190.55 if a3 else 595.3))<2 and abs(page.rect.height-841.89)<2,(label,page.rect)
   assert '#VALUE!' not in page.get_text() and '#VALOR!' not in page.get_text()
  if label=='Dashboards':
   for i in range(9):
    out=fitz.open();out.insert_pdf(doc,from_page=i,to_page=i);p=folder/f'DASH_{i+1:02}_{lang}.pdf';out.save(p);out.close();proof.append({'file':str(p.relative_to(ROOT)),'pages':1,'paper':'A3 landscape'})
  else:
   p=folder/f'{label}_{lang}.pdf';shutil.copy2(file,p);proof.append({'file':str(p.relative_to(ROOT)),'pages':count,'paper':'A4 portrait'})
  print(lang,label,count,'PASS',flush=True);return proof
 proof=[]
 with ThreadPoolExecutor(max_workers=3) as ex:
  for future in as_completed([ex.submit(task,j) for j in jobs]):proof+=future.result()
 assert len(proof)==22
 (ROOT/'PDF_Premium.json').write_text(json.dumps({'status':'PASS','platform':'LibreOffice','native_excel365':'NOT TESTED','files':proof},indent=2,ensure_ascii=False))
if __name__=='__main__':all_pdfs()
