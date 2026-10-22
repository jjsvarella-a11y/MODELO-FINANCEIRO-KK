"""Widen navigation hit areas in a recalculated private report copy; values unchanged."""
from zipfile import ZipFile,ZIP_DEFLATED
from lxml import etree as E
from pathlib import Path
import sys
from extend import paths,T

def polish(source,target):
 with ZipFile(source) as z,ZipFile(target,'w',ZIP_DEFLATED) as zz:
  _,lookup=paths(z)
  for p in z.infolist():
   data=z.read(p.filename)
   if p.filename==lookup['DASH 10 V06']:
    r=E.fromstring(data);merges=r.find(T('mergeCells'));assert merges is not None;existing={m.get('ref') for m in merges}
    for i in range(10):
     row=i*50+11
     for ref in [f'A{row}:C{row}',f'E{row}:H{row}',f'J{row}:L{row}']:
      if ref not in existing:E.SubElement(merges,T('mergeCell'),ref=ref)
    merges.set('count',str(len(merges)));data=E.tostring(r,encoding='UTF-8',xml_declaration=True)
   zz.writestr(p,data)
if __name__=='__main__':polish(*sys.argv[1:3])
