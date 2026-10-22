"""Export configured A4 areas with independent LibreOffice; never label as native Excel."""
from pathlib import Path
from zipfile import ZipFile,ZIP_DEFLATED
from lxml import etree as E
import subprocess,sys,json,shutil
import fitz
from assemble import paths,T,xml
DASHES=['EXEC V06','DASH 01 V06','DASH 02 V06','DASH 03 V06','DASH 04A V06','DASH 04B V06','DASH 05 V06','DASH 06 V06','DASH 07 V06']
ROOT=Path(__file__).resolve().parents[1];WORK=Path('/workspace/kk01_efcc')
def export(source,targets,label,folder):
 folder=Path(folder);folder.mkdir(parents=True,exist_ok=True);dest=folder/(label+'.xlsx');pdf=dest.with_suffix('.pdf')
 with ZipFile(source) as z,ZipFile(dest,'w',ZIP_DEFLATED,compresslevel=3) as zz:
  w,pp=paths(z);first=0
  for i,s in enumerate(w.find(T('sheets'))):
   s.set('state','visible' if s.get('name') in targets else 'hidden')
   if s.get('name')==targets[0]:first=i
  dn=w.find(T('definedNames'))
  if dn is not None:
   allnames=[s.get('name') for s in w.find(T('sheets'))]
   for n in list(dn):
    if n.get('name') in ['_xlnm.Print_Area','_xlnm.Print_Titles'] and (n.get('localSheetId') is None or allnames[int(n.get('localSheetId'))] not in targets):dn.remove(n)
  for v in w.find(T('bookViews')):v.set('activeTab',str(first));v.set('firstSheet',str(first))
  cp=w.find(T('calcPr'))
  if cp is not None:cp.set('fullCalcOnLoad','0');cp.set('forceFullCalc','0')
  changes={'xl/workbook.xml':xml(w)}
  for name in targets:
   r=E.fromstring(z.read(pp[name]))
   for v in r.findall(T('sheetViews')+'/'+T('sheetView')):
    for n in list(v):
     if n.tag in [T('pane'),T('selection')]:v.remove(n)
    v.set('topLeftCell','A1');v.set('tabSelected','1' if name==targets[0] else '0')
   changes[pp[name]]=xml(r)
  for p in z.infolist():zz.writestr(p.filename,changes.get(p.filename,z.read(p.filename)))
 if pdf.exists():pdf.unlink()
 proc=subprocess.run(['soffice','-env:UserInstallation='+str((folder/'profile').resolve().as_uri()),'--headless','--convert-to','pdf:calc_pdf_Export','--outdir',str(folder),str(dest)],capture_output=True,text=True,timeout=300)
 if proc.returncode or not pdf.exists():raise RuntimeError(proc.stdout+proc.stderr)
 return pdf

def all_pdfs(source,language='PT'):
 dest=ROOT/'PDF'/language;dest.mkdir(parents=True,exist_ok=True);work=WORK/('pdf_'+language);proof=[]
 pdf=export(source,DASHES,'Dashboards_01_09',work/'dashboards')
 doc=fitz.open(pdf);assert len(doc)==9,len(doc)
 for i,page in enumerate(doc):
  out=fitz.open();out.insert_pdf(doc,from_page=i,to_page=i);file=dest/f'DASH_{i+1:02}_{language}.pdf';out.save(file);out.close();proof.append({'file':str(file.relative_to(ROOT)),'pages':1,'paper':'A4 landscape','width':page.rect.width,'height':page.rect.height})
 for names,label,count in [(['DASH 10 V06'],'DASH_10_Integral',10),(['V06 Brief'],'Executive_Brief',2)]:
  pdf=export(source,names,label,work/label);doc=fitz.open(pdf);assert len(doc)==count,(label,len(doc));target=dest/(label+'_'+language+'.pdf');shutil.copy2(pdf,target);proof.append({'file':str(target.relative_to(ROOT)),'pages':len(doc),'paper':'A4 portrait','width':doc[0].rect.width,'height':doc[0].rect.height})
 (ROOT/('PDF_Validacao_'+language+'.json')).write_text(json.dumps({'platform':'LibreOffice','native_excel':'NOT TESTED','files':proof},indent=2));print('PDFs verified',language,len(proof));return proof
if __name__=='__main__':all_pdfs(Path(sys.argv[1]),sys.argv[2] if len(sys.argv)>2 else 'PT')
