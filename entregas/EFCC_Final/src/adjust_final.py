"""Final print-fit and bilingual report labels; no financial formula changes."""
from pathlib import Path
from zipfile import ZipFile,ZIP_DEFLATED
from lxml import etree as E
import sys,re
from assemble import paths,T,xml
DASHES=['EXEC V06','DASH 01 V06','DASH 02 V06','DASH 03 V06','DASH 04A V06','DASH 04B V06','DASH 05 V06','DASH 06 V06','DASH 07 V06']
LANG="'V06 Controlos'!$B$5";case="'V06 Hipoteses'!$D$27";mode="'V06 Hipoteses'!$D$26"
def adjust(source,dest):
 changes={}
 with ZipFile(source) as z:
  _,pp=paths(z);ss=E.fromstring(z.read('xl/sharedStrings.xml'));strings=[''.join(x.itertext()) for x in ss]
  cr=E.fromstring(z.read(pp['V06 Controlos']));cl=cr.find('.//'+T('c')+'[@r="B5"]');v=cl.findtext(T('v'));language=strings[int(v)] if cl.get('t')=='s' else v
  hr=E.fromstring(z.read(pp['V06 Hipoteses']));fields={}
  for row in hr.find(T('sheetData')):
   d={re.sub(r'\d','',c.get('r')):c for c in row}
   if 'A' in d and int(row.get('r'))>=5:
    c=d['A'];v=c.findtext(T('v'));key=strings[int(v)] if c.get('t')=='s' else v;fields[key]="'V06 Hipoteses'!$D$"+row.get('r')
  mode=fields['report_mode'];case=fields['report_case'];pt=language=='PT'
  friendlycase=f'IF({case}="Reference Case",IF({LANG}="PT","Caso de referência","Reference Case"),IF({LANG}="PT","Cenário ativo","Active Scenario"))'
  friendlymode=f'IF({LANG}="PT",IF({mode}="Executive Brief","Resumo Executivo","Relatório Integral"),{mode})'
  for sn in DASHES:
   r=E.fromstring(z.read(pp[sn]))
   for row in r.find(T('sheetData')):
    if int(row.get('r')) in [2,3,4]:row.set('ht','16');row.set('customHeight','1')
   changes[pp[sn]]=xml(r)
  # Older LO exports include absolute chart extents in addition to cell anchors.
  for name in z.namelist():
   if name.startswith('xl/drawings/drawing') and name.endswith('.xml'):
    r=E.fromstring(z.read(name));ns='http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing';a='http://schemas.openxmlformats.org/drawingml/2006/main';changed=False
    for anchor in r:
     fr=anchor.find('{'+ns+'}from');chart=anchor.find('.//{http://schemas.openxmlformats.org/drawingml/2006/chart}chart')
     if fr is None or chart is None:continue
     row=int(fr.find('{'+ns+'}row').text);col=int(fr.find('{'+ns+'}col').text)
     if row in [34,46] and col==4:
      ext=anchor.find('{'+ns+'}graphicFrame/{'+ns+'}xfrm/{'+a+'}ext')
      if ext is not None:ext.set('cy',str(126*12700));changed=True
    if changed:changes[name]=xml(r)
  for sn in ['DASH 10 V06','V06 Brief']:
   r=E.fromstring(z.read(pp[sn]))
   for row in r.find(T('sheetData')):
    for c in row:
     address=c.get('r');rownum=int(re.search(r'\d+',address)[0]);col=re.match('[A-Z]+',address)[0];f=c.find(T('f'));v=c.find(T('v'))
     if rownum%50==3 and col=='E' and f is not None:
      f.text=f.text.replace(case,friendlycase);v.text=(v.text or '').replace('Active Scenario','Cenário ativo' if pt else 'Active Scenario').replace('Reference Case','Caso de referência' if pt else 'Reference Case')
     if rownum%50==11 and col=='E' and f is not None:
      # Only hyperlink display name changes; mode still controls the destination.
      old=f.text;prefix,sep,last=old.rpartition(','+mode+')')
      if sep:f.text=prefix+','+friendlymode+')';v.text='Relatório Integral' if pt else 'Complete Report'
     if rownum%50==11 and col in ['A','J']:
      for n in list(c):c.remove(n)
      c.set('t','str');a,b=('← Executivo','← Executive') if col=='A' else ('Imprimir / PDF','Print / PDF');E.SubElement(c,T('f')).text=f'IF({LANG}="PT","{a}","{b}")';E.SubElement(c,T('v')).text=a if pt else b
   changes[pp[sn]]=xml(r)
  # Correct an over-broad caption: fund balance is separately reported on page 7.
  oldnew={'FF&E dotação / fundo final':'Dotações FF&E','FF&E contribution / closing fund':'FF&E contributions'}
  for node in ss.iter(T('t')):
   if node.text in oldnew:node.text=oldnew[node.text]
  changes['xl/sharedStrings.xml']=xml(ss)
  for sn in ['DASH 10 V06','V06 Traducoes']:
   r=E.fromstring(changes.get(pp[sn],z.read(pp[sn])))
   for node in r.iter(T('v')):
    if node.text in oldnew:node.text=oldnew[node.text]
   changes[pp[sn]]=xml(r)
  with ZipFile(dest,'w',ZIP_DEFLATED,compresslevel=6) as zz:
   for part in z.infolist():zz.writestr(part.filename,changes.get(part.filename,z.read(part.filename)))
 print('Final print adjustments',dest)
if __name__=='__main__':adjust(Path(sys.argv[1]),Path(sys.argv[2]))
