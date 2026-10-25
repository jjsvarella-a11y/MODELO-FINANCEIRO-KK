"""Resolve issues found by rendering the ten dashboards, not just inspecting XML.
Preserve original input units, remove stale captions, expose all sensitivity rows,
and connect availability/pricing visuals to the actual inventory query mask.
"""
from wiring_v08 import *
def review(source,target,reference,charts_proof=None):
 src=Source(source);ref=Source(reference);styles=add_styles(src.styles);changes={};proof=[]
 # Two-decimal percentage fields retain basis-point precision in editable inputs.
 fmts=src.styles.find(T('numFmts'));xfs=src.styles.find(T('cellXfs'));num=max(int(n.get('numFmtId')) for n in fmts)+1;E.SubElement(fmts,T('numFmt'),numFmtId=str(num),formatCode='0.00%');xf=deepcopy(xfs[styles['input_pct']]);xf.set('numFmtId',str(num));styles['precise_pct']=len(xfs);xfs.append(xf)
 for n in [fmts,xfs]:n.set('count',str(len(n)))
 roots={}
 def sh(sn):
  if sn not in roots:roots[sn]=Sheet(E.fromstring(src.z.read(src.pp[sn])),styles)
  return roots[sn]
 for definition in DEFS:
  key=definition['key'];sn=definition.get('sheet');addr=definition.get('cell')
  if not key.startswith('I') or key in ALIASES or sn in PANELS or sn not in [d[0] for d in DASHES[:9]]:continue
  old=ref.cells(sn).get(addr)
  if old is None:continue
  kind=ref.format_kind(sn,addr);style=styles['precise_pct'] if kind=='percent' or definition.get('kind')=='pct' else int(old.get('s','0'));sh(sn).cells[addr].set('s',str(style));proof.append({'input':sn+'!'+addr,'unit_preserved':kind})
 for addr in ['H29','N25']:sh('DASH 07 V06').cells[addr].set('s',str(styles['precise_pct']))
 captions=[('Fontes e aplicações','Sources and uses','Liquidez e dívida','Liquidity and debt','Desempenho global','Overall performance'),('Orçamento por capítulo','Budget by chapter','Desembolso trimestral','Quarterly disbursement','Curva S acumulada','Cumulative S curve'),('IVA e recuperação','VAT and recovery','IRPC anual','Annual corporate tax','Seguros e liquidez','Insurance and liquidity'),('Utilização e amortização','Drawdown and repayment','Serviço e cobertura','Service and coverage','CFADS e liquidez','CFADS and liquidity'),('Disponibilidade por tipologia','Availability by typology','Pricing por tipologia','Pricing by typology','Recebimentos trimestrais','Quarterly collections'),('Yield dos proprietários','Owner cash yield','Receita e OPEX','Revenue and OPEX','Distribuições e reservas','Distributions and reserves'),('Evolução de receitas','Revenue evolution','ADR × Ocupação / VAL','ADR × Occupancy / NPV','EBITDA e margem','EBITDA and margin'),('Estrutura acionista','Ownership structure','Progresso da captação','Capital raise progress','Composição dos blocos','Block composition'),('Dividendos anuais','Annual dividends','Dividendos e saída','Dividends and exit','Fluxos e waterfall','Cash flows and waterfall')]
 for idx,(sn,*_) in enumerate(DASHES[:9]):
  s=sh(sn)
  if sn not in PANELS:s.merge('E40:P41',style='bg')
  for row in range(22,42):s.row(row).set('ht','11')
  for row in range(42,50):s.row(row).set('ht','12')
  for row in [22,42,43]:s.row(row).set('ht','18')
  for row in [50,51]:s.row(row).set('ht','12')
  for row in range(58,90):s.row(row).attrib.pop('hidden',None);s.row(row).set('ht','5.5')
  for row in range(90,110):s.row(row).attrib.pop('hidden',None);s.row(row).set('ht','12')
  s.row(90).set('ht','18');s.row(57).set('ht','18')
  if sn=='DASH 05 V06':
   for row in [58,59,79,80,81]:s.row(row).set('ht','9')
  words=captions[idx]
  if sn!='DASH 03 V06':
   for area,pt,en in [('E57:K57',words[0],words[1]),('L57:R57',words[2],words[3]),('S57:Y57',words[4],words[5])]:s.text(area,pt,en,'section')
  s.text('W52:Y53','Resultados ativos','Active results','label');s.merge('W54:Y56',src.get('EXEC V06','H23')[1],q('EXEC V06','H23'),'card_text')
 # Active Scenario on Executive, including the actual bank/operating assumptions.
 s=sh('EXEC V06');s.text('E26:J26','Cenário ativo','Active scenario','section');s.row(26).set('ht','18');s.text('K23:P26','Base e Comparativo são independentes. Selecionar os resultados ativos não substitui os pressupostos guardados.','Base and Comparative are independent. Selecting active results does not replace saved assumptions.','note');s.text('Q22:Y22','Pressupostos do cenário ativo','Active scenario assumptions','section')
 for row,key,pt,en,kind in [(23,'bank_tan','TAN','TAN','pct'),(26,'bank_grace','Carência · meses','Grace · months','count'),(29,'adr','ADR inicial','Opening ADR','adr'),(32,'occ','Ocupação inicial','Opening occupancy','pct'),(35,'mice','Utilização MICE','MICE utilisation','pct'),(38,'ffe','Reserva FF&E','FF&E reserve','pct')]:s.text('Q'+str(row)+':U'+str(row+1),pt,en,'label');s.merge('V'+str(row)+':Y'+str(row+1),src.get(BASE7,'D'+str(ROWS[key]))[1],active(key),kind)
 def am(row):return 'IF(\'EXEC V06\'!$H$23="Comparative",'+q(MET7,'C'+str(row))+','+q(MET7,'B'+str(row))+')'
 warning='IF('+am(19)+'>0,'+tr('Défice de financiamento. ','Funding shortfall. ')+',"")&IF('+am(8)+'<0,'+tr('Reservas não cobertas. ','Reserve shortfall. ')+',"")&IF('+am(26)+'<'+active('bank_dscr_min')+','+tr('DSCR abaixo da meta interna (perímetro do motor). ','DSCR below internal target (engine perimeter). ')+',"")&IF('+am(10)+'<0,'+tr('VAL negativo à taxa de desconto ativa.','Negative NPV at the active discount rate.')+',"")';s.merge('E31:P35','DSCR abaixo da meta interna. VAL negativo.',warning,'note')
 # Chapter selection navigates directly to the original editable chapter row.
 s=sh('DASH 01 V06');codes=[src.get(CON,'A'+str(r))[1] for r in range(12,42)];s.text('E29:G30','Capítulo a consultar','Chapter to view','label');s.merge('H29:P30',codes[0],style='input_text');names=src.w.find(T('definedNames'));E.SubElement(names,T('definedName'),name='KK08_BudgetChapters').text=q(CON,'$A$12:$A$41');dvs=s.r.find(T('dataValidations'));node=E.SubElement(dvs,T('dataValidation'),type='list',sqref='H29',showErrorMessage='1',errorStyle='stop');E.SubElement(node,T('formula1')).text='KK08_BudgetChapters';s.merge('E32:P34','Abrir capítulo selecionado →','HYPERLINK("#\'DASH 01 V06\'!E"&(193+MATCH(H29,'+q(CON,'$A$12:$A$41')+',0)),'+tr('Abrir capítulo selecionado →','Open selected chapter →')+')','nav')
 # Inventory view charts use the existing query result (H mask, G status),
 # while project-wide receipts deliberately remain a whole-project indicator.
 state=sh(STATUS7);records=json.loads(Path(charts_proof or Path(source).parent/'Charts_Proof.json').read_text());types=list(dict.fromkeys(str(src.get('V06 Unidades','B'+str(r))[1]) for r in range(5,99)))
 for record in records:
  if record['sheet']!='DASH 04A V06' or record['chart_index'] not in [0,1]:continue
  start,n=record['support_rows']
  for i,typ in enumerate(types):
   lit='"'+typ+'"';mask=q('V06 Sales R5','H5:H98');tt=q('V06 Sales R5','B5:B98');status=q('V06 Sales R5','G5:G98');area='SUMIFS('+q('V06 Sales R5','D5:D98')+','+tt+','+lit+','+mask+',1)';price='SUMIFS('+q('V06 Sales R5','E5:E98')+','+tt+','+lit+','+mask+',1)'
   formula='SUMIFS('+mask+','+tt+','+lit+','+status+',"Available")' if record['chart_index']==0 else 'IF('+area+'>0,'+price+'/'+area+',"")';state.put('B'+str(start+i),None,formula)
 for sn,s in roots.items():changes[src.pp[sn]]=s.finish()
 changes['xl/styles.xml']=xml(src.styles);changes['xl/workbook.xml']=xml(src.w);rewrite(source,target,changes);src.z.close();ref.z.close();return proof
if __name__=='__main__':
 p=argparse.ArgumentParser();p.add_argument('source',type=Path);p.add_argument('target',type=Path);p.add_argument('--reference',type=Path,required=True);p.add_argument('--proof',type=Path,required=True);a=p.parse_args();a.proof.write_text(json.dumps(review(a.source,a.target,a.reference),ensure_ascii=False,indent=2))
