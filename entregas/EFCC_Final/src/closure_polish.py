"""Final presentation: cumulative quarterly curve, true bridge, readable controls.
All added data cells project existing outputs. No engine formula is changed.
"""
from zipfile import ZipFile,ZIP_DEFLATED
from lxml import etree as E
import copy,json,xlsxwriter,re
from closure_r5 import *
from closure_topics import D,C,R,P,CT
from assemble import relpath,resolve

def apply():
 # Standard generator only; no bespoke drawing/label layouts.
 b=xlsxwriter.Workbook(WORK/'Budget_Charts.xlsx');b.add_worksheet(CON)
 quarterly=b.add_chart({'type':'column'});acc=b.add_chart({'type':'line'})
 for name,c in [('Reference','B'),('Test','C')]:quarterly.add_series({'name':name,'categories':'='+q(CON,'$A$46:$A$65'),'values':'='+q(CON,'$'+c+'$46:$'+c+'$65')})
 for name,c in [('Reference cumulative','D'),('Test cumulative','E')]:acc.add_series({'name':name,'categories':'='+q(CON,'$A$46:$A$65'),'values':'='+q(CON,'$'+c+'$46:$'+c+'$65'),'y2_axis':True})
 quarterly.combine(acc);quarterly.set_y_axis({'num_format':'0,," M€"'});quarterly.set_y2_axis({'num_format':'0,," M€"'});quarterly.set_legend({'position':'bottom','font':{'size':9}})
 bridge=b.add_chart({'type':'column','subtype':'stacked'})
 for name,c,colour in [('Base','C',None),('Increase / total','D','1765A8'),('Decrease','E','C8101E')]:bridge.add_series({'name':name,'categories':'='+q(CON,'$A$110:$A$113'),'values':'='+q(CON,'$'+c+'$110:$'+c+'$113'),'fill':{'color':colour} if colour else {'none':True},'border':{'none':True}})
 bridge.set_y_axis({'num_format':'0,," M€"'});bridge.set_legend({'position':'bottom','font':{'size':9}});dummy=b.get_worksheet_by_name(CON);dummy.insert_chart('J2',quarterly);dummy.insert_chart('J20',bridge);b.close()
 with ZipFile(MASTER) as z,ZipFile(WORK/'Budget_Charts.xlsx') as assets:
  wb,pp=paths(z);ss=[''.join(n.itertext()) for n in E.fromstring(z.read('xl/sharedStrings.xml'))];style={};changes={}
  for c in E.fromstring(z.read(pp[ACT])).find(T('sheetData'))[0]:
   if c.get('t')=='s':style[ss[int(c.findtext(T('v')))]]=int(c.get('s'))
  def sheet(sn):return Sheet(E.fromstring(changes.get(pp[sn],z.read(pp[sn]))),style)
  def chartparts(sn):
   rr=E.fromstring(z.read(relpath(pp[sn])));d=next(x for x in rr if x.get('Type').endswith('/drawing'));p=resolve(pp[sn],d.get('Target'));rels=E.fromstring(z.read(relpath(p)));return [resolve(p,x.get('Target')) for x in rels if x.get('Type').endswith('/chart')]
  for sn in [CON,'DASH 01 V06']:
   targets=chartparts(sn);changes[targets[1]]=assets.read('xl/charts/chart1.xml');changes[targets[2]]=assets.read('xl/charts/chart2.xml')
  changes[chartparts('V06 PDF 02')[0]]=assets.read('xl/charts/chart1.xml')
  s=sheet(CON)
  for rr,f1,f2,f3 in [(110,'0','B110','0'),(111,'MIN(B110,B110+B111)','MAX(0,B111)','MAX(0,-B111)'),(112,'MIN(B110+B111,B110+B111+B112)','MAX(0,B112)','MAX(0,-B112)'),(113,'0','B113','0')]:
   for c,f in [('C',f1),('D',f2),('E',f3)]:s.put(c+str(rr),0,f)
  changes[pp[CON]]=s.finish()
  # Short bilingual status and derived economic equivalents; no fictitious promoter ticket.
  s=sheet(INV)
  for rr in range(5,20):
   s.put('E'+str(rr),10000000 if rr in [5,6,12] else 5000000 if rr in [13,14] else 1000000,'D'+str(rr)+'*SUM($F$12:$F$19)/SUM($D$12:$D$19)',style['money']);s.put('G'+str(rr),'Direito adquirido' if rr<12 else 'Proposta',tr('Direito adquirido','Acquired right') if rr<12 else tr('Proposta','Proposed'))
   if rr>=12:s.put('H'+str(rr),0,'IF(AND(N(\'V06 Capital\'!H'+str(rr-7)+')>0,N(\'V06 Capital\'!I'+str(rr-7)+')>0),MIN(F'+str(rr)+',N(\'V06 Capital\'!H'+str(rr-7)+')),0)',style['money'])
  for c in ['A','M']:s.put(c+'26','Capital aplicado no cenário',tr('Capital aplicado no cenário','Scenario capital invested'),style['label'])
  changes[pp[INV]]=s.finish()
  s=sheet('DASH 06 V06')
  for area,pt,en in [('E226:F227','Bloco','Block'),('G226:H227','Titular','Owner'),('I226:J227','Classe','Class'),('K226:K227','%','%'),('L226:N227','Equivalente €','Equivalent €'),('O226:Q227','Ticket novo €','New ticket €'),('R226:T227','Realizado doc.','Documented paid'),('U226:W227','Por realizar','Unpaid'),('X226:Y227','Estado','Status')]:s.text(area,pt,en,'head')
  s.text('E270:Y271','Legenda de direitos e condições por classe','Class rights and terms legend','section')
  for j,(cat,share,ticket) in enumerate([('Platinum',.2,10000000),('Gold',.1,5000000),('Silver',.02,1000000)]):
   rr=274+j*4;s.merge(f'E{rr}:G{rr+3}',cat,style='head');s.merge(f'H{rr}:L{rr+3}',f'{share:.0%} | {ticket/1e6:.0f} M€',style='money');s.text(f'M{rr}:R{rr+3}','Informação e governação: condições documentais a aprovar','Information and governance: documentary terms pending approval','note');s.text(f'S{rr}:Y{rr+3}','Facilidades propostas: a definir / aprovar','Proposed access benefits: to be defined / approved','note')
  changes[pp['DASH 06 V06']]=s.finish()
  # Saleable area is sourced from units; 47,000 m² is the existing cost-allocation basis.
  s=sheet(SALES)
  for addr,f in [('B114','SUM(D5:D98)'),('B115','SUM(E5:E98)/SUM(D5:D98)'),('B116','SUM(E5:E98)'),('B117','47000')]:s.put(addr,13290 if addr=='B114' else 47000 if addr=='B117' else 0,f)
  changes[pp[SALES]]=s.finish();s=sheet('DASH 04A V06')
  for c1,c2,label,en,addr,kind,cache in [('E','I','Área vendável m²','Saleable area m²','B114','number',13290),('J','N','€/m² ponderado','Weighted €/m²','B115','money',3811.613),('O','S','VGV','GDV','B116','money',50656336),('T','Y','Área global · imputação m²','Global cost basis m²','B117','number',47000)]:s.text(f'{c1}184:{c2}185',label,en,'label');s.merge(f'{c1}186:{c2}187',cache,q(SALES,addr),kind)
  s.text('E210:Y213','Estados comerciais: projeções, não prova documental. 47.000 m² é a área global já usada pelo motor para imputar custos; não é área construída residencial certificada. Os filtros de consulta não mudam os fluxos; pricing e atrasos mudam.','Commercial states are projections, not documentary evidence. 47,000 m² is the existing global cost-allocation basis, not a certified residential built area. View filters do not change cash flows; pricing and delays do.','note');changes[pp['DASH 04A V06']]=s.finish()
  # Grey inactive Custom fields; one authoritative editable source per parameter.
  sr=E.fromstring(z.read('xl/styles.xml'));dx=sr.find(T('dxfs'));idx=len(dx);d=E.SubElement(dx,T('dxf'));font=E.SubElement(d,T('font'));E.SubElement(font,T('color'),rgb='FF687586');pf=E.SubElement(E.SubElement(d,T('fill')),T('patternFill'),patternType='solid');E.SubElement(pf,T('fgColor'),rgb='FFF0F2F5');E.SubElement(pf,T('bgColor'),rgb='FFF0F2F5');dx.set('count',str(len(dx)));changes['xl/styles.xml']=xml(sr)
  for sn,areas,condition in [('DASH 07 V06',['N25:P25'],'$H$29<>"Custom"'),('DASH 05 V06',['H25:J25','H27:J27'],'$H$23<>"Custom"'),('DASH 05 V06',['H29:J29','H30:J30','N32:P32','H34:J34','N34:P34'],'$H$32<>"Custom Terms"'),('DASH 04A V06',['N25:P25'],'$H$25<>"Custom"'),('DASH 04B V06',['H25:J25','N25:P25'],'$H$23<>"Custom Terms"')]:
   s=sheet(sn)
   for area in areas:cf=E.SubElement(s.r,T('conditionalFormatting'),sqref=area);rule=E.SubElement(cf,T('cfRule'),type='expression',dxfId=str(idx),priority='1');E.SubElement(rule,T('formula')).text=condition
   changes[pp[sn]]=s.finish()
  # Delta IRR is shown in percentage points, not a relative percentage.
  for sn,_,_ in DASHES:
   s=sheet(sn);s.merge('X31:Y31',0,q(IMP,'D10')+'*100','pp');changes[pp[sn]]=s.finish()
  for n in range(1,10):
   sn=f'V06 PDF {n:02}';s=sheet(sn);s.merge('I55:J55',0,q(IMP,'D10')+'*100','pp');changes[pp[sn]]=s.finish()
  s=sheet('DASH 10 V06')
  for row in [44,584]:s.merge(f'K{row}:L{row}',0,q(IMP,'D10')+'*100','pp')
  changes[pp['DASH 10 V06']]=s.finish()
  for sn,column in [('DASH 06 V06','D'),('DASH 07 V06','P')]:
   s=sheet(sn);s.merge('V200:Y201',0,q(INV,column+'29')+'*100','pp');changes[pp[sn]]=s.finish()
  # Read-only active states link to their single public control owner.
  s=sheet('DASH 03 V06')
  for area,sn,addr,pt,en,f in [('K27:P28','DASH 07 V06','H23','Reservas: ','Reserves: ',q('V06 Hipoteses','D19')),('K29:P30','DASH 05 V06','H23','Operação: ','Operations: ',q('V06 Registo Ativo','B82')),('K31:P32','DASH 04A V06','H23','Vendas · atraso: ','Sales delay: ',q('V06 Registo Ativo','B37'))]:s.merge(area,pt,'HYPERLINK("#\''+sn+'\'!'+addr+'",'+tr(pt,en)+'&'+f+')','input')
  changes[pp['DASH 03 V06']]=s.finish()
  for sn,_,_ in DASHES[:9]:
   s=sheet(sn);s.text('E39:P39','Azul: teste editável. Cinzento: resultado ou Custom inativo.','Blue: editable test. Grey: result or inactive Custom.','note');changes[pp[sn]]=s.finish()
  # Currency axes are explicit; count charts retain integer axes.
  candidates=set()
  for sn in [CON,CASH,INV,SALES,'DASH 01 V06','DASH 03 V06','DASH 06 V06','DASH 04A V06','V06 PDF 02','V06 PDF 08']:candidates.update(chartparts(sn))
  for p in candidates:
   r=E.fromstring(changes.get(p,z.read(p)));formulas=[n.text or '' for n in r.findall('.//{'+C+'}f')]
   if any("'"+INV+"'!$B$55" in f or "'"+SALES+"'!$H$105" in f for f in formulas):continue
   for ax in r.findall('.//{'+C+'}valAx'):
    n=ax.find('{'+C+'}numFmt')
    if n is None:
     n=E.Element('{'+C+'}numFmt');order=['axId','scaling','delete','axPos','majorGridlines','minorGridlines','title','numFmt','majorTickMark','minorTickMark','tickLblPos','spPr','txPr','crossAx','crosses','crossesAt','crossBetween','majorUnit','minorUnit','dispUnits','extLst'];ax.append(n);ax[:]=sorted(ax,key=lambda e:order.index(E.QName(e).localname) if E.QName(e).localname in order else 999)
    n.set('formatCode','0,," M€"');n.set('sourceLinked','0')
   changes[p]=xml(r)
  temp=MASTER.with_suffix('.polish.xlsx')
  with ZipFile(temp,'w',ZIP_DEFLATED,compresslevel=6) as out:
   for item in z.infolist():out.writestr(copy.copy(item),changes.get(item.filename,z.read(item.filename)))
 temp.replace(MASTER);print('R5 final presentation projections and controls applied',flush=True)
if __name__=='__main__':apply()
