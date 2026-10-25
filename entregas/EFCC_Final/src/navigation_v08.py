"""Safe navigation/report/sensitivity presentation changes on existing sheets.
Retains every financial engine part, canonical input coordinate and chart series.
"""
from robust_filters_v08 import *
from premium_refine import add_styles,resolve,relpath,NSR,NSP

def apply(source,target):
 src=Source(source);changes={};sty=add_styles(src.styles);proof={'navigation':[],'chart_moves':[],'financial_engine_changes':0,'layout_acceptance':'PARTIAL — full ten-dashboard gabarito not certified'}
 def sh(sn):return Sheet(E.fromstring(changes.get(src.pp[sn],src.z.read(src.pp[sn]))),sty)
 for i,(sn,pt,en) in enumerate(DASHES,1):
  s=sh(sn)
  for view in s.r.findall('.//'+T('sheetView')):
   view.set('tabSelected','1' if i==1 else '0');view.set('zoomScale','85');view.set('showGridLines','0')
  # Compact the existing controls without moving canonical input addresses.
  if sn!=CENTER:
   for row in range(22,40):s.row(row).set('ht','16')
  s.link('E5:H5','EXEC V06','A1','Início','Home');s.link('I5:L5',BASE7,'A1','Parâmetros','Assumptions');s.link('M5:P5','V06 KPI Dictionary' if 'V06 KPI Dictionary' in src.pp else 'V06 Dicionario','A1','Dicionário KPI','KPI Dictionary') if 'V06 Dicionario' in src.pp or 'V06 KPI Dictionary' in src.pp else None
  s.link('Q5:T5',REPORT7,'A1','Imprimir / Exportar PDF','Print / Export PDF');s.link('U5:Y5',CENTER,'A1','Relatórios','Reports')
  proof['navigation'].extend([{'source':sn+'!E5','target':'EXEC V06!A1'},{'source':sn+'!I5','target':BASE7+'!A1'},{'source':sn+'!Q5','target':REPORT7+'!A1'}])
  changes[src.pp[sn]]=s.finish()
 # Verified destinations are actual section headings, not A1.
 sn='DASH 01 V06';s=sh(sn)
 for area,addr,pt,en in [('E38:J39','E190','Ver Orçamento Sintético','View Construction Budget'),('K38:P39','E230','Ver Quarterly Disbursement','View Quarterly Disbursement')]:
  assert src.get(sn,addr)[0];s.link(area,sn,addr,pt,en);proof['navigation'].append({'source':sn+'!'+area.split(':')[0],'target':sn+'!'+addr})
 s.link('T190:Y191',sn,'A1','Voltar ao DASH 02','Back to DASH 02');s.link('T230:Y231',sn,'A1','Voltar ao DASH 02','Back to DASH 02');changes[src.pp[sn]]=s.finish()
 # Central actual-NPV matrix: view references existing reconciled sensitivity
 # results; no separate financial calculation and no hard-coded NPV.
 sn='DASH 05 V06';s=sh(sn);s.text('L58:R59','ADR × Ocupação → VAL','ADR × Occupancy → NPV','section');s.text('L60:L62','Ocup.','Occup.','table_head')
 bands=[('M','M','I'),('N','O','M'),('P','Q','Q'),('R','R','U')]
 for start,end,old in bands:
  s.merge(start+'60:'+end+'62',src.get(sn,old+'118')[1],q(sn,old+'118'),'adr')
 for r0,oldrow in zip([63,66,69,72,75],[120,122,124,126,128]):
  s.merge('L'+str(r0)+':L'+str(r0+2),src.get(sn,'E'+str(oldrow))[1],q(sn,'E'+str(oldrow)),'pct')
  for start,end,old in bands:
   addr=start+str(r0);s.merge(addr+':'+end+str(r0+2),src.get(sn,old+str(oldrow))[1],q(sn,old+str(oldrow)),'heat')
   cf=E.SubElement(s.r,T('conditionalFormatting'),sqref=addr+':'+end+str(r0+2))
   for kind,formula in [('negative',addr+'<-$W$116'),('near','ABS('+addr+')<=$W$116'),('positive',addr+'>$W$116')]:
    rule=E.SubElement(cf,T('cfRule'),type='expression',dxfId=str(sty['dxf'][kind]),priority='1');E.SubElement(rule,T('formula')).text=formula
   # Same closest-grid convention as the original matrix, using active axes.
   formula='AND(ABS('+start+'$60-$E$9)=MIN(ABS($M$60-$E$9),ABS($N$60-$E$9),ABS($P$60-$E$9),ABS($R$60-$E$9)),ABS($L'+str(r0)+'-$I$9)=MIN(ABS($L$63-$I$9),ABS($L$66-$I$9),ABS($L$69-$I$9),ABS($L$72-$I$9),ABS($L$75-$I$9)))'
   rule=E.SubElement(cf,T('cfRule'),type='expression',dxfId=str(sty['dxf']['active']),priority='1');E.SubElement(rule,T('formula')).text=formula
 s.text('L79:R81','VAL em M€ · cores: negativo / perto de zero / positivo. Contorno: cenário mais próximo.','NPV in EUR m · negative / near zero / positive. Border: nearest scenario.','note');s.link('L83:R86',sn,'E114','Equilíbrio, escala e metodologia →','Break-even, scale and methodology →');changes[src.pp[sn]]=s.finish()
 # Relocate only drawing anchors, preserving the underlying chart XML/series.
 for sn,kind in [('DASH 05 V06','hotel'),('DASH 03 V06','bank')]:
  root=E.fromstring(src.z.read(src.pp[sn]));drawing=root.find(T('drawing'))
  if drawing is None:continue
  rels=E.fromstring(src.z.read(relpath(src.pp[sn])));rid=drawing.get('{'+NSR+'}id');part=resolve(src.pp[sn],next(x.get('Target') for x in rels if x.get('Id')==rid));dr=E.fromstring(src.z.read(part));anchors=list(dr)
  dn='{http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing}'
  def move(node,c,r,c2,r2):
   for tag,co,ro in [('from',c,r),('to',c2,r2)]:
    m=node.find(dn+tag)
    for t,val in [('col',co),('row',ro),('colOff',0),('rowOff',0)]:m.find(dn+t).text=str(val)
  if kind=='hotel':move(anchors[1],4,330,14,355);proof['chart_moves'].append({'sheet':sn,'action':'Composition chart moved to detail; actual-NPV heatmap displayed centrally'})
  else:
   assert len(anchors)==9
   for idx,node in enumerate(anchors[:3]):move(node,4+idx*7,580,11+idx*7,605)
   for node,coords in zip(anchors[4:8],[(4,57,14,72),(14,57,25,72),(4,73,14,89),(14,73,25,89)]):move(node,*coords)
   s=sh(sn);s.text('E57:N57','Utilização e amortização · Serviço da dívida','Drawdown & repayment · Debt service','section');s.text('O57:Y57','CFADS e dívida · Liquidez crítica','CFADS and debt · Critical liquidity','section');changes[src.pp[sn]]=s.finish();proof['chart_moves'].append({'sheet':sn,'action':'Four existing financing charts promoted to primary chart area'})
  changes[part]=xml(dr)
 # Reporting cards lead to the actual format selector; hyperlinks do not execute
 # software or silently change a scenario. Existing report-generation gate stays.
 sn=CENTER;s=sh(sn)
 for area,pt,en in [('E8:I13','One-Page Executive','One-Page Executive'),('J8:N13','Bank Financing','Bank Financing'),('O8:S13','Equity Investor','Equity Investor'),('T8:Y13','Comprehensive Project Report','Comprehensive Project Report')]:s.link(area,REPORT7,'B6',pt,en)
 s.text('E15:Y17','Selecione o formato na configuração. A emissão usa a Base guardada e exige validação e snapshot.','Choose the format in configuration. Issuance uses the saved Base and requires validation and a snapshot.','note')
 s.link('E19:P20',REPORT7,'B6','Configurar relatório, destinatário e comparação →','Configure report, recipient and comparison →');s.link('Q19:Y20',STATUS7,'B9','Verificar estado de gravação →','Check saved state →')
 # Old chapter index remains reachable as detail, not competing primary cards.
 for rr in range(22,79):s.row(rr).set('hidden','1')
 s.link('E80:O82',REPORT7,'A1','Validação · Snapshot · Pré-visualização · Exportação','Validation · Snapshot · Preview · Export');s.link('P80:Y82',HIST7,'A1','Histórico de cenários','Scenario history')
 changes[src.pp[sn]]=s.finish();changes['xl/styles.xml']=xml(src.styles)
 # One active tab only; do not leave the workbook in grouped editing mode.
 wb=src.w
 for view in wb.findall('.//'+T('workbookView')):view.set('activeTab',str(list(wb.find(T('sheets'))).index(next(n for n in wb.find(T('sheets')) if n.get('name')=='EXEC V06'))))
 changes['xl/workbook.xml']=xml(wb)
 rewrite(source,target,changes);src.z.close();return proof
if __name__=='__main__':
 p=argparse.ArgumentParser();p.add_argument('source',type=Path);p.add_argument('target',type=Path);p.add_argument('--proof',type=Path,required=True);a=p.parse_args();a.proof.write_text(json.dumps(apply(a.source,a.target),ensure_ascii=False,indent=2))
