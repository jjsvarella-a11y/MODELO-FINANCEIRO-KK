"""Keep the hotel warning clear of the page edge and record scenario on all views."""
from pathlib import Path
from zipfile import ZipFile,ZIP_DEFLATED
from lxml import etree as E
import copy,sys,re
from assemble import paths,T,xml
DASHES=['EXEC V06','DASH 01 V06','DASH 02 V06','DASH 03 V06','DASH 04A V06','DASH 04B V06','DASH 05 V06','DASH 06 V06','DASH 07 V06']
def finish(source,dest):
 changes={}
 with ZipFile(source) as z:
  _,pp=paths(z)
  for sn in DASHES:
   root=E.fromstring(z.read(pp[sn]));sd=root.find(T('sheetData'));cells={c.get('r'):c for c in root.iter(T('c'))};row=sd.find(T('row')+'[@r="1"]');style=cells['E5'].get('s','0')
   if row is None:row=E.Element(T('row'),r='1',ht='12',customHeight='1');sd.insert(0,row)
   row.set('ht','12');row.set('customHeight','1');c=cells.get('E1')
   if c is None:c=E.SubElement(row,T('c'),r='E1')
   for n in list(c):c.remove(n)
   c.set('s',style);c.set('t','str');E.SubElement(c,T('f')).text='\'V06 Hipoteses\'!$D$29&" · "&\'V06 Hipoteses\'!$D$28&" · "&\'V06 Controlos\'!$B$5';E.SubElement(c,T('v')).text='EFCC-20261022-01 · 2026-10-22 · PT'
   # Correct snapshot language without changing the language control.
   strings=[''.join(s.itertext()) for s in E.fromstring(z.read('xl/sharedStrings.xml'))];lr=E.fromstring(z.read(pp['V06 Controlos'])).find('.//'+T('c')+'[@r="B5"]');lv=lr.findtext(T('v'));language=strings[int(lv)] if lr.get('t')=='s' else lv;c.find(T('v')).text='EFCC-20261022-01 · 2026-10-22 · '+str(language)
   merges=root.find(T('mergeCells'));E.SubElement(merges,T('mergeCell'),ref='E1:P1');merges.set('count',str(len(merges)))
   if sn=='DASH 05 V06':
    old=cells['E61'];new=copy.deepcopy(old);new.set('r','E59');rr=sd.find(T('row')+'[@r="59"]');existing=rr.find(T('c')+'[@r="E59"]')
    if existing is not None:rr.remove(existing)
    rr.append(new)
    for n in list(old):old.remove(n)
    old.attrib.pop('t',None)
    for m in merges:
     if m.get('ref')=='E61:P62':m.set('ref','E59:P60')
   def ci(c):
    v=0
    for ch in re.match('[A-Z]+',c.get('r'))[0]:v=v*26+ord(ch)-64
    return v
   for r in sd:r[:]=sorted(list(r),key=ci)
   changes[pp[sn]]=xml(root)
  with ZipFile(dest,'w',ZIP_DEFLATED,compresslevel=6) as zz:
   for part in z.infolist():zz.writestr(part.filename,changes.get(part.filename,z.read(part.filename)))
 print('Print edge / scenario labels',dest)
if __name__=='__main__':finish(Path(sys.argv[1]),Path(sys.argv[2]))
