"""Print comparisons, topic details, native print areas and scalar delta colours."""
from zipfile import ZipFile,ZIP_DEFLATED
from lxml import etree as E
import copy,json,posixpath
from closure_r5 import *
from assemble import relpath,resolve
from closure_topics import D,C,R,P,CT

def apply():
 source=Source(WORK/'full_recalc2/out'/MASTER.name);changes={};changed=[]
 with ZipFile(MASTER) as z:
  wb,pp=paths(z);names=[n.get('name') for n in wb.find(T('sheets'))];stylesroot=E.fromstring(z.read('xl/styles.xml'));ss=[''.join(n.itertext()) for n in E.fromstring(z.read('xl/sharedStrings.xml'))];style={}
  for c in E.fromstring(z.read(pp[ACT])).find(T('sheetData'))[0]:
   if c.get('t')=='s':style[ss[int(c.findtext(T('v')))]]=int(c.get('s'))
  def sheet(sn):return Sheet(E.fromstring(changes.get(pp[sn],z.read(pp[sn]))),style)
  def v(sn,addr):
   value=source.get(sn,addr)[1];return value if value is not None else ''
  def alias(s,area,sn,addr,kind='money'):s.merge(area,v(sn,addr),q(sn,addr),kind)
  def draw(sn):
   r=E.fromstring(z.read(pp[sn]));rels=E.fromstring(z.read(relpath(pp[sn])));d=next(x for x in rels if x.get('Type').endswith('/drawing'));p=resolve(pp[sn],d.get('Target'));dr=E.fromstring(z.read(p));rp=relpath(p);rr=E.fromstring(z.read(rp));cp=[resolve(p,x.get('Target')) for x in rr if x.get('Type').endswith('/chart')];return p,dr,cp
  # Project-wide state and finance eligibility used by conditional formatting.
  s=sheet(IMP);s.put('C20',1,'AND(MAX('+motor('funding_gap')+')<=0.005,MIN(\''+CASH+'\'!O5:O184)>=-0.005)')
  old=s.cells['B16'];f=old.findtext(T('f'));gate='IF('+q('V06 Motor',col(CC['debt'])+'58')+'>0.005,'+tr('M54: dívida por liquidar','M54: debt outstanding')+',IF(\'V06 Controlos\'!B57="Suspended",'+tr('M54: dividendos suspensos','M54: dividends suspended')+','+f+'))';s.put('B16',v(IMP,'B16'),gate,int(old.get('s')));changes[pp[IMP]]=s.finish()
  dxfs=stylesroot.find(T('dxfs'));dxf={}
  for key,bg,fg in [('green','E0F1E5','166534'),('red','FBE5E7','991B1B')]:
   dxf[key]=len(dxfs);d=E.SubElement(dxfs,T('dxf'));font=E.SubElement(d,T('font'));E.SubElement(font,T('color'),rgb='FF'+fg);p=E.SubElement(E.SubElement(d,T('fill')),T('patternFill'),patternType='solid');E.SubElement(p,T('fgColor'),rgb='FF'+bg);E.SubElement(p,T('bgColor'),rgb='FF'+bg)
  dxfs.set('count',str(len(dxfs)));changes['xl/styles.xml']=xml(stylesroot)
  dn=wb.find(T('definedNames'))
  if not any(n.get('name')=='KK5_Finance_OK' for n in dn):E.SubElement(dn,T('definedName'),name='KK5_Finance_OK').text=q(IMP,'$C$20')
  for index,(sn,_,_) in enumerate(DASHES):
   s=sheet(sn);s.text('Q22:Y23','Impacto Global do Cenário no Projeto','Project-Wide Scenario Impact','section');s.merge('Q35:Y36','Modo ativo: Test →','HYPERLINK("#\'EXEC V06\'!H23",'+tr('Modo ativo: ','Active mode: ')+'&\'EXEC V06\'!H23&" →")','input')
   for j in range(9):
    direction=-1 if j in [0,1] else 1;rr=154+j*2;areas=[(f'S{rr}:U{rr+1}','S'+str(rr))]
    if j<6:areas.append((f'X{26+j}:Y{26+j}','X'+str(26+j)))
    for area,cell in areas:
     cf=E.SubElement(s.r,T('conditionalFormatting'),sqref=area);tol=.00000001 if j in [5,6] else .005
     for key,expr in [('red',f'{cell}*{direction}<-{tol}'),('green',f'AND({cell}*{direction}>{tol},KK5_Finance_OK)')]:
      rule=E.SubElement(cf,T('cfRule'),type='expression',dxfId=str(dxf[key]),priority='1');E.SubElement(rule,T('formula')).text=f'IF(ISNUMBER({cell}),{expr},FALSE)'
   changes[pp[sn]]=s.finish()
  # Source sheet labels clarify that the first model row is the active case.
  s=sheet('V06 Cenarios');s.put('A5','Ativo / Teste','IF(\'EXEC V06\'!H23="Reference","Reference R4","Active / Test")');changes[pp['V06 Cenarios']]=s.finish()
  # Six critical cash months in consolidated page 3 and individual cash page.
  for sn,first in [('DASH 10 V06',156),('V06 PDF 04',38)]:
   s=sheet(sn);s.merge(f'A{first}:L{first+6}',None,style='note')
   for j in range(6):
    rr=first+j;mr=190+j;alias(s,f'A{rr}:F{rr}',CASH,'A'+str(mr),'label');alias(s,f'G{rr}:H{rr}',CASH,'B'+str(mr),'number');alias(s,f'I{rr}:L{rr}',CASH,'C'+str(mr),'ratio' if j==4 else 'money')
   changes[pp[sn]]=s.finish()
  # Selected block comparisons on shareholder PDFs, without removing their chart.
  for sn,off in [('V06 PDF 08',0),('V06 PDF 09',12)]:
   s=sheet(sn);s.merge('G22:L35',None,style='note');s.text('G22:L23','Impacto no Investidor selecionado','Selected Investor Impact','section')
   for j in range(11):
    rr=j+24;mr=j+26;kind='pct' if j in [1,3,6,9,10] else 'ratio' if j==5 else 'number' if j==7 else 'base'
    alias(s,f'G{rr}:I{rr}',INV,col(off)+str(mr),'label');alias(s,f'J{rr}:J{rr}',INV,col(off+1)+str(mr),kind);alias(s,f'K{rr}:K{rr}',INV,col(off+2)+str(mr),kind);alias(s,f'L{rr}:L{rr}',INV,col(off+3)+str(mr),kind)
   changes[pp[sn]]=s.finish();p,dr,cps=draw(sn)
   for anchor in dr:
    end=anchor.find('{'+D+'}to');end.find('{'+D+'}col').text='5';end.find('{'+D+'}colOff').text='0';anchor.set('editAs','twoCell');ext=anchor.find('.//{http://schemas.openxmlformats.org/drawingml/2006/main}ext')
    if ext is not None:ext.set('cx','5137500')
   changes[p]=xml(dr)
   if sn=='V06 PDF 08':changes[cps[0]]=z.read(draw(INV)[2][0])
  # Construction screen and its dedicated PDF share the new quarterly comparison.
  p,dr,cps=draw('V06 PDF 02');changes[cps[0]]=z.read(draw(CON)[2][1])
  # Report comparisons identify their columns and no longer show a stale historical case.
  s=sheet('DASH 10 V06')
  for rr in [38,578]:
   s.text(f'A{rr}:F{rr}','Indicador','KPI','head');s.text(f'G{rr}:H{rr}','Referência R4','Reference R4','head');s.text(f'I{rr}:J{rr}','Ativo','Active','head');s.text(f'K{rr}:L{rr}','Variação','Delta','head')
  changes[pp['DASH 10 V06']]=s.finish()
  # Native construction appendix: reference plus actual active extension of quarters.
  s=sheet(CON);cols=s.r.find(T('cols'));cols[:]=[]
  for j,width in enumerate([8,50,18,10,12,18,18,18],1):E.SubElement(cols,T('col'),min=str(j),max=str(j),width=str(width),customWidth='1')
  setup=s.r.find(T('pageSetup'))
  if setup is None:setup=E.SubElement(s.r,T('pageSetup'))
  setup.set('paperSize','8');setup.set('orientation','landscape');setup.set('fitToWidth','1');setup.set('fitToHeight','0')
  pr=s.r.find(T('sheetPr'))
  if pr is None:pr=E.Element(T('sheetPr'));s.r.insert(0,pr)
  fit=pr.find(T('pageSetUpPr'))
  if fit is None:fit=E.SubElement(pr,T('pageSetUpPr'))
  fit.set('fitToPage','1')
  rb=s.r.find(T('rowBreaks'))
  if rb is None:rb=E.SubElement(s.r,T('rowBreaks'))
  rb.clear();rb.set('count','1');rb.set('manualBreakCount','1');E.SubElement(rb,T('brk'),id='44',max='16383',man='1')
  for rr in range(4,106):s.row(rr).set('ht','15');s.row(rr).set('customHeight','1')
  changes[pp[CON]]=s.finish();idx=names.index(CON)
  for n in list(dn):
   if n.get('name')=='_xlnm.Print_Area' and n.get('localSheetId')==str(idx):dn.remove(n)
  E.SubElement(dn,T('definedName'),name='_xlnm.Print_Area',localSheetId=str(idx)).text='OFFSET(\''+CON+'\'!$A$1,0,0,45+ROUNDUP((30+\'V06 Controlos\'!$B$38)/3,0),8)'
  s=sheet('DASH 01 V06');s.merge('E229:Y229','Imprimir orçamento por capítulos →','HYPERLINK("#\''+CON+'\'!A1",'+tr('Imprimir orçamento por capítulos →','Print chapter budget →')+')','input');changes[pp['DASH 01 V06']]=s.finish()
  changes['xl/workbook.xml']=xml(wb)
  # Release labels only: do not modify financial formulas.
  for sn in [n[0] for n in DASHES]+['DASH 10 V06','V06 Brief']+[f'V06 PDF {i:02}' for i in range(1,10)]:
   r=E.fromstring(changes.get(pp[sn],z.read(pp[sn])))
   for n in r.iter():
    if n.text and ('PREMIUM-R2' in n.text or 'PREMIUM R2' in n.text):n.text=n.text.replace('PREMIUM-R2 CHECKPOINT','R5 CHECKPOINT').replace('PREMIUM R2','R5 CHECKPOINT')
   changes[pp[sn]]=xml(r)
  temp=MASTER.with_suffix('.print.xlsx')
  with ZipFile(temp,'w',ZIP_DEFLATED,compresslevel=6) as out:
   for item in z.infolist():out.writestr(copy.copy(item),changes.get(item.filename,z.read(item.filename)))
 temp.replace(MASTER);print('R5 print areas and comparison colours applied',flush=True)
if __name__=='__main__':apply()
