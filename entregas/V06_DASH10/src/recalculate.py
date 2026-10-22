"""Invalidate formula caches in a private copy, then force external recalculation.
The source remains unchanged. LibreOffice is not an Excel compatibility certificate.
"""
from pathlib import Path
from zipfile import ZipFile,ZIP_DEFLATED
from lxml import etree as E
import subprocess,sys,json
N='http://schemas.openxmlformats.org/spreadsheetml/2006/main';T=lambda n:'{'+N+'}'+n
def recalculate(source,folder):
 folder=Path(folder);folder.mkdir(parents=True,exist_ok=True);(folder/'out').mkdir(exist_ok=True);profile=folder/'profile/user';profile.mkdir(parents=True,exist_ok=True)
 (profile/'registrymodifications.xcu').write_text('<oor:items xmlns:oor="http://openoffice.org/2001/registry"><item oor:path="/org.openoffice.Office.Calc/Formula/Load"><prop oor:name="OOXMLRecalcMode" oor:op="fuse"><value>0</value></prop></item></oor:items>')
 dest=folder/Path(source).name
 with ZipFile(source) as z,ZipFile(dest,'w',ZIP_DEFLATED) as zz:
  for p in z.infolist():
   data=z.read(p.filename)
   if p.filename.startswith('xl/worksheets/') and p.filename.endswith('.xml'):
    root=E.fromstring(data)
    for c in root.iter(T('c')):
     if c.find(T('f')) is not None:
      c.attrib.pop('t',None);v=c.find(T('v'))
      if v is not None:c.remove(v)
    data=E.tostring(root,encoding='UTF-8',xml_declaration=True)
   zz.writestr(p,data)
 output=folder/'out'/dest.name
 if output.exists():output.unlink()
 cmd=['soffice','-env:UserInstallation='+str((folder/'profile').resolve().as_uri()),'--headless','--convert-to','xlsx','--outdir',str(folder/'out'),str(dest)]
 proc=subprocess.run(cmd,capture_output=True,text=True,timeout=300)
 if proc.returncode or not output.exists():raise RuntimeError(proc.stdout+proc.stderr)
 print(output,flush=True);return output
if __name__=='__main__':recalculate(Path(sys.argv[1]),Path(sys.argv[2]))
