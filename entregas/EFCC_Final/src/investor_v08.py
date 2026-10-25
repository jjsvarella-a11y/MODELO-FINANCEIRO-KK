"""Correct residual R4/active-context references in selected investor comparison.
No new financial engine: same existing block formulas rebound per scenario.
"""
from charts_v08 import *
def investor_fix(source,target):
 src=Source(source);styles=add_styles(src.styles);changes={};s=Sheet(E.fromstring(src.z.read(src.pp[INV])),styles);proof=[]
 for rr in [32,34,35,36]:
  original=src.get(INV,'C'+str(rr))[0]
  for cc,mode in [('B','Base'),('C','Comparative')]:
   old,cached=src.get(INV,cc+str(rr));new=rebound(original,mode);s.put(cc+str(rr),cached,new);proof.append({'cell':INV+'!'+cc+str(rr),'before':old,'after':new})
 changes[src.pp[INV]]=s.finish()
 for sn in ['DASH 06 V06','DASH 07 V06']:
  s=Sheet(E.fromstring(src.z.read(src.pp[sn])),styles);s.text('E114:Y115','Impacto no Investidor Selecionado — Base | Comparativo | Δ','Selected Investor Impact — Base | Comparative | Δ','section')
  s.text('E116:M116','Indicador','Metric','table_head');s.text('N116:Q116','Base','Base','table_head');s.text('R116:U116','Comparativo','Comparative','table_head');s.text('V116:Y116','Delta','Delta','table_head')
  for rr,ir,pt,en,kind in [(117,26,'Capital aplicado','Capital invested','money'),(118,28,'Dividendos','Dividends','money'),(119,29,'TIR','IRR','pct'),(120,30,'VAL','NPV','money'),(121,31,'MOIC','MOIC','ratio'),(122,33,'Payback · mês','Payback · month','count'),(123,34,'Proveitos na saída','Exit proceeds','money'),(124,35,'Retorno dependente da saída','Return dependent on exit','pct'),(125,36,'Capital recuperado por dividendos','Capital recovered through dividends','pct')]:
   s.text('E'+str(rr)+':M'+str(rr),pt,en,'label')
   for col0,col1,mc in [('N','Q','B'),('R','U','C'),('V','Y','D')]:s.merge(col0+str(rr)+':'+col1+str(rr),src.get(INV,mc+str(ir))[1],q(INV,mc+str(ir)),kind)
  s.link('E128:O130',sn,'E190','Registo e direitos do investidor →','Investor register and rights →');s.link('P128:Y130','EXEC V06','A1','Início','Home')
  # Retire obsolete duplicate R4-labelled Project Impact Summary, preserving its
  # cells for existing references. Primary current-Base summary is at row 42.
  for rr in range(150,181):s.row(rr).set('hidden','1')
  if sn=='DASH 07 V06':
   def active_inv(row):return 'IF(\'EXEC V06\'!$H$23="Comparative",'+q(INV,'C'+str(row))+','+q(INV,'B'+str(row))+')'
   for addr,row in [('E54',29),('H54',30),('K54',31),('N54',28),('T54',34)]:s.put(addr,src.get(sn,addr)[1],active_inv(row))
   s.text('N52:P53','Dividendos do investidor','Investor dividends','label');s.text('Q52:S53','Capital recuperado','Capital recovered','label');s.text('T52:V53','Proveitos na saída','Exit proceeds','label');s.put('Q54',None,active_inv(36)+'*'+active_inv(26),style='card_money')
  changes[src.pp[sn]]=s.finish()
 # Defined denominator checks report genuine absence of activity as text,
 # rather than introducing #DIV/0! in the new presentation.
 sn='DASH 05 V06';s=Sheet(E.fromstring(src.z.read(src.pp[sn])),styles)
 moved={sn+'!'+addr:STATUS7+'!'+ui_archive_address(sn,addr) for addr in ['H44','E44','H50']}
 for addr,original in [('E54','H44'),('H54','E44'),('T54','H50')]:
  den=q(*moved[sn+'!'+original].split('!'));f,cached=src.get(sn,addr);s.put(addr,cached,'IF(('+den+')>0,'+f+','+tr('Sem atividade','No activity')+')')
 changes[src.pp[sn]]=s.finish();changes['xl/styles.xml']=xml(src.styles);rewrite(source,target,changes);src.z.close();return proof
if __name__=='__main__':
 p=argparse.ArgumentParser();p.add_argument('source',type=Path);p.add_argument('target',type=Path);p.add_argument('--proof',type=Path,required=True);a=p.parse_args();a.proof.write_text(json.dumps(investor_fix(a.source,a.target),ensure_ascii=False,indent=2))
