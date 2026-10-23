"""Export verified PT/EN views. Native Excel printing remains NOT TESTED."""
from pathlib import Path
from concurrent.futures import ThreadPoolExecutor,as_completed
from zipfile import ZipFile,ZIP_DEFLATED
from lxml import etree as E
import json,shutil,copy,fitz
from closure_r5 import *
from export_all import export

def area_copy(source,language):
 target=WORK/('Appendix_'+language+'.xlsx')
 with ZipFile(source) as z,ZipFile(target,'w',ZIP_DEFLATED,compresslevel=3) as out:
  wb,pp=paths(z);names=[n.get('name') for n in wb.find(T('sheets'))];idx=names.index(CON);row=45+int(__import__('math').ceil((30+Source(source).get('V06 Controlos','B38')[1])/3))
  for n in wb.find(T('definedNames')):
   if n.get('name')=='_xlnm.Print_Area' and n.get('localSheetId')==str(idx):n.text=q(CON,'$A$1:$H$'+str(row))
  for item in z.infolist():out.writestr(copy.copy(item),xml(wb) if item.filename=='xl/workbook.xml' else z.read(item.filename))
 return target

def run(language):
 source=Path(json.loads((WORK/'Presentations.json').read_text())[language]);dest=OUT/'PDF'/language;dest.mkdir(parents=True,exist_ok=True);proof=[]
 tasks=[([f'V06 PDF {i:02}' for i in range(1,10)],'Dashboards',9),(['DASH 10 V06'],'DASH_10',10),(['V06 Brief'],'Executive_Brief',2),([CON],'Budget_Appendix',None)]
 for targets,label,count in tasks:
  file=area_copy(source,language) if label=='Budget_Appendix' else source;pdf=export(file,targets,label,WORK/('pdf_'+language)/label);d=fitz.open(pdf)
  if count is not None:assert len(d)==count,(language,label,len(d))
  for i,p in enumerate(d):
   text=p.get_text();assert not any(x in text for x in ['#VALUE!','#REF!','#NAME?','#N/A','###']),(language,label,i)
   proof.append({'document':label,'page':i+1,'width':p.rect.width,'height':p.rect.height,'text_chars':len(text)})
  if label=='Dashboards':
   for i in range(9):
    doc=fitz.open();doc.insert_pdf(d,from_page=i,to_page=i);doc.save(dest/f'DASH_{i+1:02}_{language}.pdf')
  else:shutil.copy2(pdf,dest/(label+'_'+language+'.pdf'))
  print(language,label,len(d),flush=True)
 # The second dashboard includes its chapter/quarter appendix in the delivered PDF.
 file=dest/f'DASH_02_{language}.pdf';d=fitz.open(file);appendix=fitz.open(dest/f'Budget_Appendix_{language}.pdf');d.insert_pdf(appendix);d.save(WORK/('Budget_Complete_'+language+'.pdf'));d.close();shutil.copy2(WORK/('Budget_Complete_'+language+'.pdf'),file)
 # All 20 real NPV points must be present on consolidated sensitivity page 9.
 original=json.loads((ROOT/'Premium_Checkpoint/Testes_Matriz_Investimento.json').read_text())['base']['npv_points'];d=fitz.open(dest/f'DASH_10_{language}.pdf');txt=d[8].get_text()
 for point in original:assert f"{abs(point['npv'])/1e6:.2f}" in txt,(language,point)
 return {'language':language,'platform':'LibreOffice export','native_excel365':'NOT TESTED','pages':proof,'budget_includes_appendix':True,'matrix_values_verified':20}

def main():
 with ThreadPoolExecutor(max_workers=2) as ex:results=list(ex.map(run,['PT','EN']))
 (OUT/'PDF_Validation.json').write_text(json.dumps(results,indent=2));print('PDF exports complete',flush=True)
if __name__=='__main__':main()
