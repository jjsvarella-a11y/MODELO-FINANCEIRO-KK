"""Export the entire report: private print copy, no page-range truncation."""
from pathlib import Path
from zipfile import ZipFile,ZIP_DEFLATED
from lxml import etree as E
import subprocess,sys
from extend import paths,T

def export_pdf(source,folder):
 folder=Path(folder);folder.mkdir(parents=True,exist_ok=True);dest=folder/'DASH10_integral.xlsx'
 with ZipFile(source) as z,ZipFile(dest,'w',ZIP_DEFLATED) as zz:
  w,lookup=paths(z);index=0
  for i,s in enumerate(w.find(T('sheets'))):
   s.set('state','visible' if s.get('name')=='DASH 10 V06' else 'hidden')
   if s.get('name')=='DASH 10 V06':index=i
  defined=w.find(T('definedNames'))
  if defined is not None:
   for n in list(defined):
    if n.get('name') in ['_xlnm.Print_Area','_xlnm.Print_Titles'] and int(n.get('localSheetId','-1'))!=index:defined.remove(n)
  for view in w.find(T('bookViews')):view.set('activeTab',str(index));view.set('firstSheet',str(index))
  for part in z.infolist():
   data=E.tostring(w,encoding='UTF-8',xml_declaration=True) if part.filename=='xl/workbook.xml' else z.read(part.filename)
   if part.filename==lookup['DASH 10 V06']:
    root=E.fromstring(data)
    for view in root.findall(T('sheetViews')+'/'+T('sheetView')):
     for node in list(view):
      if node.tag in [T('pane'),T('selection')]:view.remove(node)
     view.set('topLeftCell','A1')
    data=E.tostring(root,encoding='UTF-8',xml_declaration=True)
   zz.writestr(part,data)
 pdf=folder/'DASH10_integral.pdf'
 if pdf.exists():pdf.unlink()
 proc=subprocess.run(['soffice','-env:UserInstallation='+str((folder/'profile').resolve().as_uri()),'--headless','--convert-to','pdf:calc_pdf_Export','--outdir',str(folder),str(dest)],capture_output=True,text=True,timeout=240)
 if proc.returncode or not pdf.exists():raise RuntimeError(proc.stdout+proc.stderr)
 print(pdf);return pdf
if __name__=='__main__':export_pdf(Path(sys.argv[1]),Path(sys.argv[2]))
