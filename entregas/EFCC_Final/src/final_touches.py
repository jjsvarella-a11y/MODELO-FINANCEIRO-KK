from reference_io import load_reference
"""Small OOXML presentation corrections after generation, without rewriting engines."""
from pathlib import Path
from zipfile import ZipFile,ZIP_DEFLATED
from lxml import etree as E
import sys,copy,json
from assemble import paths,T,xml
ROOT=Path(__file__).resolve().parents[1]
def touches(source):
 ref=load_reference(ROOT);changes={}
 with ZipFile(source) as z:
  w,pp=paths(z)
  # Dynamic opening month in full report: construction delays must propagate.
  root=E.fromstring(z.read(pp['DASH 10 V06']))
  for c in root.iter(T('c')):
   f=c.find(T('f'))
   if f is None:continue
   for key in ['adr','occupancy']:
    from xlsxwriter.utility import xl_col_to_name
    col=xl_col_to_name(ref['main_columns'][key]);old=f"'V06 Motor'!{col}35";new=f"INDEX('V06 Motor'!${col}$5:${col}$184,31+'V06 Controlos'!$B$38)"
    if old in (f.text or ''):f.text=f.text.replace(old,new)
  changes[pp['DASH 10 V06']]=xml(root)
  # Selected-year EBITDA must explicitly flag horizons beyond 180 months.
  root=E.fromstring(z.read(pp['V06 Matriz Valores']))
  for c in root.iter(T('c')):
   if c.get('r','').startswith('E') and c.find(T('f')) is not None:
    f=c.find(T('f'));old=f.text
    if old.startswith('SUM(INDEX('):
     # 30 + delay + operational year*12 <= 180.
     f.text='IF(30+\'V06 Controlos\'!$B$38+\'DASH 05 V06\'!N23*12<=180,'+old+',"Selected year exceeds 180-month horizon")'
  changes[pp['V06 Matriz Valores']]=xml(root)
  # Explicit link to the existing three remaining custom fee inputs.
  root=E.fromstring(z.read(pp['DASH 05 V06']));sd=root.find(T('sheetData'));row=sd.find(T('row')+'[@r="32"]')
  if row is None:row=E.SubElement(sd,T('row'),r='32')
  cs=row.find(T('c')+'[@r="N32"]')
  if cs is None:cs=E.SubElement(row,T('c'),r='N32')
  cs.set('t','str');E.SubElement(cs,T('f')).text='HYPERLINK("#\'V06 Controlos\'!B68",IF(\'V06 Controlos\'!B5="PT","Base / sistemas / incentivo →","Base / systems / incentive →"))';E.SubElement(cs,T('v')).text='Base / sistemas / incentivo →'
  # Preserve monotonically increasing cell order.
  for rr in sd:
   def ci(c):
    import re
    a=re.match('[A-Z]+',c.get('r','A1'))[0];v=0
    for ch in a:v=v*26+ord(ch)-64
    return v
   rr[:]=sorted(list(rr),key=ci)
  changes[pp['DASH 05 V06']]=xml(root)
  # Registry covers every setting, including report IDs, date and selector states.
  root=E.fromstring(z.read(pp['V06 Registo Ativo']));sd=root.find(T('sheetData'));hyp=E.fromstring(z.read(pp['V06 Hipoteses']));ss=[''.join(x.itertext()) for x in E.fromstring(z.read('xl/sharedStrings.xml'))]
  def get(c):
   v=c.findtext(T('v'));return ss[int(v)] if c.get('t')=='s' else v
  last=max(int(r.get('r')) for r in sd)
  for hr in hyp.find(T('sheetData')):
   rowid=int(hr.get('r'))
   if rowid<5:continue
   cells={c.get('r').rstrip('0123456789'):c for c in hr};key=get(cells['A']);val=cells['D'];unit=get(cells['E']);last+=1;nr=E.SubElement(sd,T('row'),r=str(last))
   for c,text in [('A','selector.'+key),('C',unit or ''),('D',f"'V06 Hipoteses'!D{rowid}")]:
    nc=E.SubElement(nr,T('c'),r=c+str(last),t='inlineStr');E.SubElement(E.SubElement(nc,T('is')),T('t')).text=text
   nc=E.SubElement(nr,T('c'),r='B'+str(last));E.SubElement(nc,T('f')).text=f"'V06 Hipoteses'!D{rowid}"
   typ=val.get('t');v=get(val)
   if typ in ['s','str','inlineStr']:nc.set('t','str')
   E.SubElement(nc,T('v')).text=v
   # A B C D ordering
   nr[:]=sorted(list(nr),key=lambda c:c.get('r')[0])
  root.find(T('dimension')).set('ref','A4:E'+str(last));changes[pp['V06 Registo Ativo']]=xml(root)
  tmp=Path(source).with_suffix('.touched.xlsx')
  with ZipFile(tmp,'w',ZIP_DEFLATED,compresslevel=6) as zz:
   for part in z.infolist():zz.writestr(part.filename,changes.get(part.filename,z.read(part.filename)))
 tmp.replace(source)
 print('Presentation refinements',len(changes))
if __name__=='__main__':touches(Path(sys.argv[1]))
