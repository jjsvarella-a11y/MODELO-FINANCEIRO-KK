"""Scalar formula corrections, bilingual explanatory labels, no cache suppression."""
from zipfile import ZipFile,ZIP_DEFLATED
from lxml import etree as E
import copy,re
from assemble import paths,T,xml
from closure_r5 import MASTER,INV,CASH,IMP,ACT,OUT,Sheet,tr,q,motor

def apply():
 changes={}
 with ZipFile(MASTER) as z:
  _,pp=paths(z)
  for sn in [INV,CASH,IMP]:
   r=E.fromstring(z.read(pp[sn]));s=Sheet(r,{})
   for f in r.iter(T('f')):
    if re.fullmatch(r'0[A-Z]+220',f.text or ''):f.text='0+'+f.text[1:]
   if sn==INV:
    pairs=[('Capital investido','Capital invested'),('Participação','Share'),('Dividendos acumulados','Dividends'),('TIR','IRR'),('VAL','NPV'),('MOIC','MOIC'),('Cash-on-cash anual','Cash-on-cash annual'),('Payback · mês','Payback month'),('Proveitos na saída','Exit proceeds'),('Retorno dependente da saída','Return dependent on exit'),('Capital recuperado por dividendos','Capital returned via dividends')]
    for start in ['A','M']:
     for j,(pt,en) in enumerate(pairs,26):c=s.cells[start+str(j)];s.put(start+str(j),pt,tr(pt,en),int(c.get('s')))
   elif sn==CASH:
    pairs=[('Mês de menor caixa','Minimum cash'),('Menor folga sobre reservas','Minimum reserve headroom'),('Pico de dívida','Peak debt'),('Maior défice de financiamento','Maximum funding gap'),('DSCR mínimo com serviço de dívida','Minimum DSCR with service'),('Pico do crédito de IVA','Maximum VAT credit')]
    for j,(pt,en) in enumerate(pairs,190):c=s.cells['A'+str(j)];s.put('A'+str(j),pt,tr(pt,en),int(c.get('s','0')))
   else:
    for rr in range(5,14):
     c=s.cells['E'+str(rr)];s.put('E'+str(rr),0,'IF(ABS(B'+str(rr)+')>0.000001,D'+str(rr)+'/ABS(B'+str(rr)+'),"N/A")',int(c.get('s')))
    def m54(key):return q('V06 Motor',__import__('closure_r5').col(__import__('closure_r5').CC[key])+'58')
    legal,liq,cov,target=[m54(k) for k in ['legal_headroom','liquid_headroom','coverage_headroom','dividend_target']]
    condition='IF(MAX('+motor('funding_gap')+')>0.005,'+tr('Défice de financiamento','Funding shortfall')+',IF(MIN(\''+CASH+'\'!O5:O184)<-0.005,'+tr('Reserva não coberta','Reserve shortfall')+',IF(MIN('+legal+','+liq+','+cov+')<'+target+',IF('+legal+'<=MIN('+liq+','+cov+'),'+tr('M54: limite de lucros','M54: profit limit')+',IF('+liq+'<='+cov+','+tr('M54: liquidez / reservas','M54: liquidity / reserves')+','+tr('M54: cobertura de caixa 12M','M54: 12M cash coverage')+')),'+tr('Sem limite de distribuição ativo em M54','No active distribution cap at M54')+')))'
    c=s.cells['B16'];s.put('B16','A verificar',condition,int(c.get('s')))
    comment='IF(D5>0.005,'+tr('CAPEX aumentou. ','CAPEX increased. ')+',IF(D5<-0.005,'+tr('CAPEX diminuiu. ','CAPEX decreased. ')+',""))&IF(D6>0.005,'+tr('Pico da dívida agravou. ','Debt peak worsened. ')+'," ")&IF(D9>0.005,'+tr('VAL melhorou','NPV improved')+',IF(D9<-0.005,'+tr('VAL deteriorou','NPV worsened')+','+tr('VAL sem alteração','NPV unchanged')+'))&IF(C9<0,'+tr(' e permanece negativo. ',' and remains negative. ')+',". ")&IF(MAX('+motor('funding_gap')+')>0.005,'+tr('Liquidez insuficiente.','Insufficient liquidity.')+','+tr('Sem défice de financiamento modelado.','No modelled funding shortfall.')+')'
    c=s.cells['B18'];s.put('B18','A verificar',comment,int(c.get('s')))
   changes[pp[sn]]=s.finish()
  temp=MASTER.with_suffix('.adjust.xlsx')
  with ZipFile(temp,'w',ZIP_DEFLATED,compresslevel=6) as out:
   for item in z.infolist():out.writestr(copy.copy(item),changes.get(item.filename,z.read(item.filename)))
 temp.replace(MASTER);print('Adjusted scalar cumulative sums, labels and comparisons',flush=True)
if __name__=='__main__':apply()
