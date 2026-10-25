"""Retire duplicate bank edit points and keep comparison narratives context-pure."""
from scenario_v07 import *
def apply(file=FINAL7):
 source=Source(file);changes={};sheet=Sheet(E.fromstring(source.z.read(source.pp[BN])),{})
 for key,pt,en,rv,addr,kind in PARAMS:
  if int(addr[1:])<450:continue
  row=ROWS['bank_'+key];target='IF(\'EXEC V06\'!$H$23="Comparative","#\''+CMP7+'\'!B'+str(row)+'","#\''+BASE7+'\'!B'+str(row)+'")';sheet.put(addr,source.get(store('Base'),'D'+str(row))[1],'HYPERLINK('+target+','+active('bank_'+key)+')');dvs=sheet.r.find(T('dataValidations'))
  if dvs is not None:
   for dv in list(dvs):
    if addr in dv.get('sqref','').split():dvs.remove(dv)
 changes[source.pp[BN]]=sheet.finish();sheet=Sheet(E.fromstring(source.z.read(source.pp[IMP])),{})
 gap=q(MET7,'$C$19');head=q(MET7,'$C$8');npv=q(MET7,'$C$10');f='IF('+gap+'>0.005,'+tr('Comparativo: défice de financiamento','Comparative: funding shortfall')+',IF('+head+'<-0.005,'+tr('Comparativo: reserva não coberta','Comparative: reserve shortfall')+',IF('+npv+'<0,'+tr('Comparativo: VAL negativo','Comparative: negative NPV')+','+tr('Comparativo: sem défice modelado','Comparative: no modelled shortfall')+')))';sheet.put('B16','Comparativo: VAL negativo',f)
 f='IF(D9>0,'+tr('VAL do Comparativo melhorou','Comparative NPV improved')+','+tr('VAL do Comparativo não melhorou','Comparative NPV did not improve')+')&IF(C9<0,'+tr(' mas permanece negativo.',' but remains negative.')+',".")&" "&IF('+gap+'>0.005,'+tr('Comparativo com défice de financiamento.','Comparative has a funding shortfall.')+','+'IF('+head+'<-0.005,'+tr('Comparativo sem cobertura integral das reservas.','Comparative does not fully cover protected reserves.')+','+tr('Comparativo sem défice de financiamento modelado.','Comparative has no modelled funding shortfall.')+'))';sheet.put('B18','VAL do Comparativo melhorou mas permanece negativo. Comparativo sem défice de financiamento modelado.',f);changes[source.pp[IMP]]=sheet.finish()
 for sn in [REP,BRF]:
  sheet=Sheet(E.fromstring(source.z.read(source.pp[sn])),{})
  for rr in range(8,520,52):
   if 'A'+str(rr) in sheet.cells:sheet.put('A'+str(rr),'Base | To identify | Hypothetical',q('EXEC V06','H23')+'&" | "&'+d('institution')+'&" | "&'+d('status'))
  sheet.put('A2','Legacy print view — active results; controlled four-format reporting: V07 Relatorios')
  changes[source.pp[sn]]=sheet.finish()
 tmp=WORK7/'Controls.xlsx';rewrite(file,tmp,changes);source.z.close();shutil.move(tmp,file)
if __name__=='__main__':apply()
