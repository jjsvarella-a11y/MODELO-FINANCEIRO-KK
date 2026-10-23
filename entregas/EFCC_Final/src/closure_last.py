"""Active inventory dates and native dynamic quarter chart ranges."""
from zipfile import ZipFile,ZIP_DEFLATED
from lxml import etree as E
import copy,json
from closure_r5 import *
from assemble import resolve,relpath
C='http://schemas.openxmlformats.org/drawingml/2006/chart'

def apply(path):
 source=Source(path);changes={}
 with ZipFile(path) as z:
  wb,pp=paths(z);names=[s.get('name') for s in wb.find(T('sheets'))];text=[''.join(n.itertext()) for n in E.fromstring(z.read('xl/sharedStrings.xml'))];styles={}
  for c in E.fromstring(z.read(pp[ACT])).find(T('sheetData'))[0]:
   if c.get('t')=='s':styles[text[int(c.findtext(T('v')))]]=int(c.get('s'))
  def sheet(sn):return Sheet(E.fromstring(changes.get(pp[sn],z.read(pp[sn]))),styles)
  def alias(s,area,sn,addr,kind):s.merge(area,source.get(sn,addr)[1],q(sn,addr),kind)
  s=sheet(SALES)
  for rr in range(5,99):s.put('G'+str(rr),'Available','IF(\'DASH 04A V06\'!H35>=\'V06 Calend 01\'!D'+str(rr)+',"Deeded",IF(\'DASH 04A V06\'!H35>=\'V06 Calend 01\'!C'+str(rr)+',"Reserved","Available"))')
  changes[pp[SALES]]=s.finish()
  # Compact type analysis is also printed; all four unit types fit without truncation.
  for sn,first in [('DASH 10 V06',216),('V06 PDF 05',38)]:
   s=sheet(sn);s.merge(f'A{first}:L{first+6}',None,style='note')
   for area,pt,en in [(f'A{first}:B{first}','Tipo','Type'),(f'C{first}:C{first}','Nº','Units'),(f'D{first}:E{first}','Área m²','Area m²'),(f'F{first}:G{first}','€/m²','€/m²'),(f'H{first}:J{first}','VGV','GDV'),(f'K{first}:K{first}','Disp.','Avail.'),(f'L{first}:L{first}','Com.','Comm.')]:s.text(area,pt,en,'head')
   for j in range(4):
    rr=first+1+j;mr=105+j
    for area,c,kind in [(f'A{rr}:B{rr}','A','label'),(f'C{rr}:C{rr}','B','number'),(f'D{rr}:E{rr}','C','number'),(f'F{rr}:G{rr}','E','money'),(f'H{rr}:J{rr}','G','base'),(f'K{rr}:K{rr}','H','number'),(f'L{rr}:L{rr}','I','number')]:alias(s,area,SALES,c+str(mr),kind)
   changes[pp[sn]]=s.finish()
  s=sheet('DASH 06 V06');s.text('E288:Y291','Participações definem a alocação na saída. Os dividendos por classe seguem a política existente e podem não ser pró-rata. Objetivos de cenário e facilidades não equivalem a direitos contratuais verificados.','Ownership defines exit allocation. Class dividends follow existing policy and may differ from pro-rata. Scenario targets and proposed benefits are not verified contractual rights.','note');changes[pp['DASH 06 V06']]=s.finish()
  for sn,off in [('V06 PDF 08',0),('V06 PDF 09',12)]:
   s=sheet(sn);s.text('G22:I23','Impacto no Investidor','Investor Impact','section');s.text('J22:J23','Base','Ref.','head');s.text('K22:K23','Ativo','Active','head');s.text('L22:L23','Delta','Delta','head');ref=q(INV,col(off+3)+'29');s.merge('L27:L27',0,'IF(ISNUMBER('+ref+'),'+ref+'*100,"N/A")','pp');changes[pp[sn]]=s.finish()
  # Optional two-page brief retains a real reference/active comparison.
  s=sheet('V06 Brief');s.merge('A36:L48',None,style='note');s.text('A36:L37','Referência R4 vs Ativo','R4 Reference vs Active','section')
  for j in range(6):
   rr=39+j;mr=5+j;alias(s,f'A{rr}:F{rr}',IMP,'A'+str(mr),'label');alias(s,f'G{rr}:H{rr}',IMP,'B'+str(mr),'pct' if j==5 else 'base');alias(s,f'I{rr}:J{rr}',IMP,'C'+str(mr),'pct' if j==5 else 'test');alias(s,f'K{rr}:L{rr}',IMP,'D'+str(mr),'pct' if j==5 else 'delta')
  s.text('G38:H38','Base','Ref.','head');s.text('I38:J38','Ativo','Active','head');s.text('K38:L38','Delta','Delta','head');changes[pp['V06 Brief']]=s.finish()
  # Sheet-local names keep the chart to 10 quarters at Base, extending with delay.
  dn=wb.find(T('definedNames'));mapping={'Q_Labels':'A','Q_BaseCAP':'B','Q_TestCAP':'C','Q_BaseCum':'D','Q_TestCum':'E'};idx=names.index(CON)
  for name,c in mapping.items():
   for n in list(dn):
    if n.get('name')==name and n.get('localSheetId')==str(idx):dn.remove(n)
   E.SubElement(dn,T('definedName'),name=name,localSheetId=str(idx)).text='OFFSET(\''+CON+'\'!$'+c+'$46,0,0,ROUNDUP((30+\'V06 Controlos\'!$B$38)/3,0),1)'
  for part in z.namelist():
   if part.startswith('xl/charts/chart') and part.endswith('.xml'):
    r=E.fromstring(z.read(part));dirty=False
    for f in r.findall('.//{'+C+'}f'):
     for name,c in mapping.items():
      if f.text==q(CON,'$'+c+'$46:$'+c+'$65'):f.text=q(CON,name);dirty=True
    if dirty:changes[part]=xml(r)
  changes['xl/workbook.xml']=xml(wb);tmp=path.with_suffix('.last.xlsx')
  with ZipFile(tmp,'w',ZIP_DEFLATED,compresslevel=6) as out:
   for item in z.infolist():out.writestr(copy.copy(item),changes.get(item.filename,z.read(item.filename)))
 tmp.replace(path);print('Active calendar and quarter ranges',path.name,flush=True)
if __name__=='__main__':apply(MASTER)
