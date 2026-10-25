"""Align active-result tables and handle valid investor/view empty states."""
from hotel_helpers_v08 import *
def consistent_views(source,target,charts_proof=None):
 src=Source(source);changes={};styles=add_styles(src.styles)
 for sn,*_ in DASHES[:9]:
  if sn=='DASH 05 V06':continue
  s=Sheet(E.fromstring(src.z.read(src.pp[sn])),styles)
  for row,mr in [(92,6),(94,7),(96,8),(98,9),(100,10)]:
   addr='W'+str(row);f,v=src.get(sn,addr)
   if f and "'V06 Impact R5'!" in f:s.put(addr,v,'IF(\'EXEC V06\'!$H$23="Comparative",'+q(MET7,'C'+str(mr))+','+q(MET7,'B'+str(mr))+')')
  s.text('S90:Y91','Projeto — resultados ativos','Project — active results','section');changes[src.pp[sn]]=s.finish()
 sn='DASH 07 V06';s=Sheet(E.fromstring(changes[src.pp[sn]]),styles)
 def ai(row):return 'IF(\'EXEC V06\'!$H$23="Comparative",'+q(INV,'C'+str(row))+','+q(INV,'B'+str(row))+')'
 s.put('Q54',None,'IF(AND(ISNUMBER('+ai(26)+'),ISNUMBER('+ai(36)+')),MIN('+ai(26)+','+ai(26)+'*'+ai(36)+'),"N/A")',style='card_money');s.text('E125:M125','Cobertura do capital por dividendos','Capital coverage through dividends','label');s.merge('E127:Y127','', 'IF(SUM(\'V06 Investor R5\'!N5:N19)=0,'+tr('Nenhum bloco corresponde aos filtros.','No investment block matches the filters.')+',"")','note');changes[src.pp[sn]]=s.finish()
 sn='DASH 05 V06';s=Sheet(E.fromstring(src.z.read(src.pp[sn])),styles);s.put('M78');s.put('M82');s.text('L79:R81','ADR e ocupação iniciais · VAL em M€. Contorno: cenário mais próximo.','Opening ADR and occupancy · NPV in EUR m. Border: nearest scenario.','note');s.text('E133:I134','ADR mínimo testado · ocupação ativa','Minimum tested ADR · active occupancy','label');s.text('J133:N134','Ocupação mínima testada · ADR ativo','Minimum tested occupancy · active ADR','label');changes[src.pp[sn]]=s.finish()
 # The return dashboard combines a whole-project annual distribution chart,
 # a selected-investor dividend/exit chart, and the existing all-new-investor
 # cash-flow waterfall. Keep their populations explicit.
 records=json.loads(Path(charts_proof or Path(source).parent/'Charts_Proof.json').read_text());state=Sheet(E.fromstring(src.z.read(src.pp[STATUS7])),styles)
 for rec in records:
  if rec['sheet']!='DASH 07 V06':continue
  start,n=rec['support_rows']
  if rec['chart_index']==0:state.put('A'+str(start-2),'Dividendos anuais do projeto · M€',tr('Dividendos anuais do projeto · M€','Annual project dividends · EUR m'))
  if rec['chart_index']==1:
   state.put('B'+str(start-1),'Investidor selecionado',tr('Investidor selecionado','Selected investor'))
   for i,row in enumerate([28,34]):
    cc='C' if src.get('EXEC V06','H23')[1]=='Comparative' else 'B';formula='IF(\'EXEC V06\'!$H$23="Comparative",'+q(INV,'C'+str(row))+','+q(INV,'B'+str(row))+')/1000000';state.put('B'+str(start+i),src.get(INV,cc+str(row))[1]/1e6,formula)
 changes[src.pp[STATUS7]]=state.finish();s=Sheet(E.fromstring(changes[src.pp['DASH 07 V06']]),styles);s.text('E57:K57','Dividendos do projeto','Project dividends','section');s.text('L57:R57','Investidor — dividendos e saída','Investor — dividends and exit','section');s.text('S57:Y57','Waterfall — novos investidores','Waterfall — new investors','section');changes[src.pp['DASH 07 V06']]=s.finish();changes['xl/styles.xml']=xml(src.styles);rewrite(source,target,changes);src.z.close()
if __name__=='__main__':
 p=argparse.ArgumentParser();p.add_argument('source',type=Path);p.add_argument('target',type=Path);a=p.parse_args();consistent_views(a.source,a.target)
