"""Compare editable current Base and Comparative, not an obsolete raw reference.
Quarterly net/VAT/gross/cumulative amounts aggregate the existing 180-month motors.
Original Base adjustment input coordinates N194/T194..N223/T223 are preserved.
"""
from chart_caches_v08 import *
def budget_compare(source,target):
 src=Source(source);styles=add_styles(src.styles);s=Sheet(E.fromstring(src.z.read(src.pp['DASH 01 V06'])),styles)
 for addr,pt,en in [('K192','Orçamento fonte €','Source budget EUR'),('M192','Peso Base','Base share'),('N192','Δ % Base','Base adjustment %'),('O192','Δ Comp. €','Comparative delta EUR'),('Q192','Base atual €','Current Base EUR'),('T192','ETC Base €','Base ETC EUR'),('W192','Comparativo €','Comparative EUR')]:s.put(addr,pt,tr(pt,en))
 for j,row in enumerate(range(194,224),12):
  bs=MAPS['Base'][CON];cs=MAPS['Comparative'][CON];bf=q(bs,'G'+str(j))+'+'+q(bs,'H'+str(j));cf=q(cs,'G'+str(j))+'+'+q(cs,'H'+str(j));bv=src.get(bs,'G'+str(j))[1]+src.get(bs,'H'+str(j))[1];cv=src.get(cs,'G'+str(j))[1]+src.get(cs,'H'+str(j))[1]
  s.put('Q'+str(row),bv,bf);s.put('W'+str(row),cv,cf);s.put('O'+str(row),cv-bv,'('+cf+')-('+bf+')');den='SUM('+q(bs,'G12:G41')+')+SUM('+q(bs,'H12:H41')+')';total=sum(src.get(bs,c+str(r))[1] for r in range(12,42) for c in ['G','H']);s.put('M'+str(row),bv/total,'IF(('+den+')>0,('+bf+')/('+den+'),"N/A")')
 s.text('E230:S231','Desembolso trimestral — CAPEX líquido e acumulados','Quarterly disbursement — net CAPEX and cumulative amounts','section');s.link('T230:Y231','DASH 01 V06','A1','Voltar ao DASH 02','Back to DASH 02')
 for addr,pt,en in [('H232','Base atual · líquido','Current Base · net'),('L232','Comparativo · líquido','Comparative · net'),('P232','Base · acumulado','Base · cumulative'),('U232','Comparativo · acumulado','Comparative · cumulative')]:s.put(addr,pt,tr(pt,en))
 def amount(mode,key,lo,hi):
  sn=MAPS[mode]['V06 Motor'];c=col(CC[key]);return 'SUM('+q(sn,c+str(lo)+':'+c+str(hi))+')',sum(src.get(sn,c+str(r))[1] for r in range(lo,hi+1))
 for i,row in enumerate(range(234,294)):
  lo=5+3*i;hi=lo+2
  for dest,mode,cumulative in [('H','Base',False),('L','Comparative',False),('P','Base',True),('U','Comparative',True)]:f,v=amount(mode,'capex',5 if cumulative else lo,hi);s.put(dest+str(row),v,f)
 s.text('E296:Y298','CAPEX do motor inclui construção, fiscalização, contingência e ETC. O orçamento por capítulo apresenta custos diretos; o IVA e o desembolso bruto estão discriminados abaixo.','Engine CAPEX includes construction, supervision, contingency and ETC. Chapter budgets show direct costs; VAT and gross disbursement are detailed below.','note');s.text('E300:S301','IVA e desembolso bruto — Base vs Comparativo','VAT and gross disbursement — Base vs Comparative','section');s.link('T300:Y301','DASH 01 V06','A1','Voltar ao DASH 02','Back to DASH 02')
 for area,pt,en in [('E302:G303','Trimestre','Quarter'),('H302:K303','IVA Base','Base VAT'),('L302:O303','IVA Comparativo','Comparative VAT'),('P302:T303','Bruto Base','Base gross'),('U302:Y303','Bruto Comparativo','Comparative gross')]:s.text(area,pt,en,'table_head')
 for i,row in enumerate(range(304,364)):
  lo=5+3*i;hi=lo+2;s.merge('E'+str(row)+':G'+str(row),i+1,style='count')
  for start,end,mode,gross in [('H','K','Base',False),('L','O','Comparative',False),('P','T','Base',True),('U','Y','Comparative',True)]:
   f,v=amount(mode,'vat_incurred',lo,hi)
   if gross:
    nf,nv=amount(mode,'capex',lo,hi);f='('+nf+')+('+f+')';v+=nv
   s.merge(start+str(row)+':'+end+str(row),v,f,'money');s.row(row).set('ht','16')
 changes={src.pp['DASH 01 V06']:s.finish(),'xl/styles.xml':xml(src.styles)};rewrite(source,target,changes);src.z.close()
if __name__=='__main__':
 p=argparse.ArgumentParser();p.add_argument('source',type=Path);p.add_argument('target',type=Path);a=p.parse_args();budget_compare(a.source,a.target)
