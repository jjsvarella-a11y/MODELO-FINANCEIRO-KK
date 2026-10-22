"""Apply final print layout to new modules only; does not change financial engines."""
from pathlib import Path
from zipfile import ZipFile,ZIP_DEFLATED
from lxml import etree as E
import json,sys,posixpath
N='http://schemas.openxmlformats.org/spreadsheetml/2006/main';R='http://schemas.openxmlformats.org/officeDocument/2006/relationships';T=lambda n:'{'+N+'}'+n
def layout(source,output):
 with ZipFile(source) as z,ZipFile(output,'w',ZIP_DEFLATED) as zz:
  w=E.fromstring(z.read('xl/workbook.xml'));rels={r.get('Id'):r.get('Target') for r in E.fromstring(z.read('xl/_rels/workbook.xml.rels'))};names={}
  for s in w.find(T('sheets')):
   t=rels[s.get('{'+R+'}id')];names[t.lstrip('/') if t.startswith('/') else posixpath.normpath('xl/'+t)]=s.get('name')
  for p in z.infolist():
   data=z.read(p.filename)
   if names.get(p.filename,'') in ['EXEC V06','DASH 01 V06','DASH 02 V06','DASH 03 V06','DASH 04A V06','DASH 04B V06','DASH 05 V06','DASH 06 V06','DASH 07 V06']:
    root=E.fromstring(data);setup=root.find(T('pageSetup'));assert setup is not None;setup.set('orientation','portrait');setup.set('paperSize','8');data=E.tostring(root,encoding='UTF-8',xml_declaration=True)
   zz.writestr(p,data)
if __name__=='__main__':layout(Path(sys.argv[1]),Path(sys.argv[2]))
