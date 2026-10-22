# Additive V04 financial sources, all formula linked. Executed in inherited builder scope.
from waterfall import constant_policy,priority,solve,COUNTS,TICKETS
from completion import completion
from v04_formula_helpers import projected_bank_formula,terminal_assets_formula
from xlsxwriter.utility import xl_cell_to_rowcol
import numpy as np
# Derived calendar; user controls delay and operational readiness, not duplicate dates.
fw(S['Pressupostos'],PR['end'],1,'30+'+pv('construction_delay'),BASE['p']['end'],'int');fw(S['Pressupostos'],PR['open'],1,pv('end')+'+1+'+pv('opening_delay'),BASE['p']['open'],'int')

s=sh('V04 Calendarios',['Código','Acelerado','Base','Prudente'],'Meses de pagamento de cenário por bloco. Editáveis; não datas contratadas. Compromisso previsto um mês antes, mínimoM1.')
for i,x in enumerate(pipeline(BASE['p'])):
 r=i+5;text(s,r,0,x['code'])
 for c,mode in enumerate(['Acelerado','Base','Prudente'],1):text(s,r,c,pipeline(config(capital_schedule=mode))[i]['pay_month'],'input_int');s.data_validation(r-1,c,r-1,c,{'validate':'integer','criteria':'between','minimum':1,'maximum':180})
s=sh('V04 Pipeline',['Código','Categoria','Vendedor — proposta','Investidor potencial','Estado documentado','Ticket referência','Valor proposto — plano','Selecionado','Mês compromisso — plano','Mês pagamento — plano','Valor contratado','Data compromisso','Data pagamento','Recebido comprovado','Retido comprovado','Transferido KK comprovado','Instrumento jurídico','Validação documental','Elegível banco confirmado','Recebido — cenário','Retido — cenário','KK — cenário','Pago equivalente / ticket'],'12blocos; potenciais/compromissos/recebidos/transferidos separados. Campos documentados vazios significam indisponíveis. Calendário é hipótese, não execução.')
s.set_column('A:B',16);s.set_column('C:C',42);s.set_column('D:G',24);s.set_column('K:S',25)
for i,plan in enumerate(pipeline(BASE['p'])):
 r=i+5;cat=plan['category'];ticket=plan['ticket'];text(s,r,0,plan['code']);text(s,r,1,cat);text(s,r,2,'JSCV11.000 + JALS9.000 ações' if cat=='Platinum' else 'JALS — proposta de origem')
 for c in [3,4,10,11,12,13,14,15,16,17,18]:s.write_blank(r-1,c,None,F['input_text' if c in [3,4,11,12,16,17] else 'input_eur'])
 s.data_validation(r-1,4,r-1,4,{'validate':'list','source':['Prospect','Qualified','Term Sheet','Signed','Paid','Capitalized'],'ignore_blank':True})
 text(s,r,5,ticket,'eur');text(s,r,6,ticket,'input_eur')
 qty=pv('n_platinum') if i==0 else pv('n_gold') if i==1 else f'IF({pv("n_silver")}>={i-1},1,0)'
 fw(s,r,7,f'IF({pv("equity_mode")}="Combinação tickets",MIN(1,MAX(0,{qty})),1)',1,'int')
 months={'Acelerado':[1,1]+[2]*5+[3]*5,'Base':[1,2]+[3]*3+[4]*3+[5]*2+[6]*2,'Prudente':[1,3]+[6,7,8,9,10,12,14,15,17,18]}
 base=f'IF({pv("capital_schedule")} ="Acelerado",\'V04 Calendarios\'!B{r},IF({pv("capital_schedule")} ="Prudente",\'V04 Calendarios\'!D{r},\'V04 Calendarios\'!C{r}))'
 fw(s,r,8,f'MAX(1,({base})-1)',plan['commit_month'],'int');fw(s,r,9,f'({base})+{pv("capital_delay")}',plan['pay_month'],'int')
 scale=f'IF({pv("equity_mode")}="Combinação tickets",1,MAX(0,{pv("equity")}-{pv("conversion")})/25000000)'
 fw(s,r,19,f'IF(I{r}<={pv("capital_stop")},F{r}*H{r}*({scale})*{pv("received_ratio")}*{pv("capital_fraction")},0)',plan['planned_received'])
 fw(s,r,20,f'T{r}*(1-{rawpv("treasury_ratio")})',0);fw(s,r,21,f'T{r}*{rawpv("treasury_ratio")}',plan['transfer']);fw(s,r,22,f'IF(F{r}*({scale})>0,MIN(1,T{r}/(F{r}*({scale}))),0)',1,'pct')
s.autofilter('A4:W16')
for r,label,formula in [(20,'Potencial oferta','SUM(F5:F16)'),(21,'Proposto no plano','SUM(G5:G16)'),(22,'Contratado documentado','IF(COUNT(K5:K16)>0,SUM(K5:K16),"Indisponível")'),(23,'Recebido documentado','IF(COUNT(N5:N16)>0,SUM(N5:N16),"Indisponível")'),(24,'Transferido KK documentado','IF(COUNT(P5:P16)>0,SUM(P5:P16),"Indisponível")'),(25,'Elegível banco documentado','IF(COUNT(S5:S16)>0,SUM(S5:S16),"Indisponível")'),(26,'Recebimento previsto','SUM(T5:T16)'),(27,'Entrada KK prevista','SUM(V5:V16)'),(28,'Retido vendedores previsto','SUM(U5:U16)'),(29,'Reconciliação previsto','B26-B27-B28')]:text(s,r,0,label);fw(s,r,1,formula,None)
text(s,32,0,'Antes do banco: objetivo15M€');fw(s,32,1,'SUMIFS(T5:T16,J5:J16,"<=3")',None);text(s,33,0,'Colocação atéM18 — plano');fw(s,33,1,'SUMIFS(T5:T16,J5:J16,"<=18")/25000000',None,'pct')

s=sh('V04 Capital180',['Mês','Peso oferta integral','Peso seleção ativa','Recebido cenário','Retido cenário','KK cenário','Quota económica paga equivalente'],'Não confundir previsão com capital realizado. Direitos jurídicos após pagamento dependem do instrumento. Capital histórico4,5M€não é entrada futura.')
for m in range(1,181):
 r=m+4;text(s,r,0,m,'int');terms=[]
 for i in range(12):
  pr=i+5;terms.append(f'IF(AND(\'V04 Pipeline\'!J{pr}=A{r},\'V04 Pipeline\'!I{pr}<={pv("capital_stop")} ),\'V04 Pipeline\'!F{pr},0)')
 fw(s,r,1,f'SUM({",".join(terms)})/25000000*{pv("received_ratio")}*{pv("capital_fraction")}',investor_outlay(BASE['p'],m)/25e6,'pct')
 fw(s,r,3,f'IF({pv("capital_schedule")} ="V03",({pv("equity")}-{pv("conversion")})*{pv("received_ratio")}*SUM(IF(A{r}={pv("fund_month1")},{pv("fund_part1")},0),IF(A{r}={pv("fund_month2")},{pv("fund_part2")},0),IF(A{r}={pv("fund_month3")},{pv("fund_part3")},0)),SUMIF(\'V04 Pipeline\'!$J$5:$J$16,A{r},\'V04 Pipeline\'!$T$5:$T$16))',investor_outlay(BASE['p'],m))
 fw(s,r,2,f'IF({pv("equity")}>{pv("conversion")},D{r}/({pv("equity")}-{pv("conversion")} ),0)',investor_outlay(BASE['p'],m)/25e6,'pct');fw(s,r,4,f'D{r}*(1-{rawpv("treasury_ratio")})',0);fw(s,r,5,f'D{r}*{rawpv("treasury_ratio")}',BASE['flow'][m-1]['equity']);fw(s,r,6,f'IF({pv("equity")}>{pv("conversion")},SUM(D$5:D{r})/({pv("equity")}-{pv("conversion")})*\'Capital V03\'!B7,0)',economic_share(BASE['p'],m),'pct')
s.hide()

# Original chapter profile remains unmodified; execution timing and EAC adjust an additive forecast.
s=sh('V04 CAPEX Base',['Mês','Construção','Fiscalização','Contingência','Total','ETC fator'],'Perfil de30meses por capítulo; pausa e custo de atraso são cenários. ETC documentado não é inventado.')
for m in range(1,181):
 r=m+4;text(s,r,0,m,'int');src=f'IF(A{r}<{pv("pause_month")},A{r},IF(A{r}>={pv("pause_month")}+{pv("construction_delay")},A{r}-{pv("construction_delay")},0))';base=f'IF(AND(({src})>=1,({src})<=30),SUMPRODUCT(\'Obra\'!$D$5:$D$34,INDEX(\'Obra\'!$E$5:$AH$34,0,MAX(1,MIN(30,{src})))),0)';hold=f'IF(AND({pv("construction_delay")}>0,A{r}>={pv("pause_month")},A{r}<{pv("pause_month")}+{pv("construction_delay")} ),{pv("delay_cost")},0)'
 fw(s,r,1,f'{base}+{hold}',None);fw(s,r,2,f'({base})*{pv("supervision")}',None);fw(s,r,3,f'({base})*{pv("contingency")}*{pv("cont_used")}',None);fw(s,r,4,f'SUM(B{r}:D{r})',None)
 etc=f'SUMIFS($E$5:$E$184,$A$5:$A$184,">"&{pv("status_month")})';fw(s,r,5,f'IF(AND({pv("etc_override")}=1,ISNUMBER({pv("etc_amount")}),A{r}>{pv("status_month")}),IF({etc}>0,{pv("etc_amount")}/{etc},0),1)',1,'mult')
 for c in range(1,5):fw(S['CAPEX'],r,c,f"'V04 CAPEX Base'!{col(c)}{r}*'V04 CAPEX Base'!F{r}",BASE['cap'][m-1][['','base','supervision','cont','total'][c]])
s.hide()

# New reference holdings20blocks; historical versions remain unchanged.
s=S['Matriz Societaria V03'];headers(s,['Acionista','Platinum','Gold','Silver','Blocos','Quota alvo','Ações alvo','Quota prevista','Ações previstas'])
for r,counts in [(5,[1,2,0]),(6,[0,0,5]),(7,[1,1,10])]:
 for c,n in enumerate(counts,1):text(s,r,c,n,'int')
 fw(s,r,4,f'SUM(B{r}:D{r})',sum(counts),'int');fw(s,r,5,f'B{r}*20%+C{r}*10%+D{r}*2%',None,'pct');fw(s,r,6,f'ROUND(F{r}*100000,0)',None,'int')
text(s,14,0,'Blocos totais =20');fw(s,14,1,'IF(SUM(B9:D9)=20,"Conforme","Alerta")','Conforme','text');text(s,20,0,'V04: JSCV1Platinum+2Gold');text(s,21,0,'Política preferencial não aprovada')
S['Societario'].write('D5','V04: JSCV1Platinum+2Gold=40%; não confere controlo jurídico automático.',F['subtitle'])

s=sh('V04 Bancabilidade',['Fonte / natureza','Referência histórica','Contabilístico confirmado','Realizado confirmado','Elegível banco confirmado','Documento / validação'],'Campos confirmados vazios=indisponíveis. Não somar vendas secundárias e transferência do mesmo preço como duas entradas. Histórico4,5M€sujeito a prova.')
for r,label in enumerate(['Investimento histórico promotores','Capital social realizado','Contribuições patrimoniais','Suprimentos/subordinados','Venda secundária de blocos','Transferências àKK','Compromissos firmes não realizados','Recursos previstos não contratados'],5):
 text(s,r,0,label);text(s,r,1,4.5e6 if r==5 else None,'eur')
 for c in range(2,6):s.write_blank(r-1,c,None,F['input_text' if c==5 else 'input_eur'])
for r,c in [(15,2),(16,3),(17,4)]:text(s,r,0,['Equity contabilístico conhecido','Capital realizado conhecido','Elegível banco conhecido'][r-15]);fw(s,r,1,f'IF(COUNT({col(c)}5:{col(c)}12)>0,SUM({col(c)}5:{col(c)}8,{col(c)}10),"Indisponível")',None)
text(s,20,0,'Limite estratégico contratado — cenário');fw(s,20,1,pv('limit'),30e6);text(s,21,0,'Utilização acumulada');fw(s,21,1,kref('used'),None);text(s,22,0,'Disponibilidade contratual / mobilizável');fw(s,22,1,f'{pv("limit")}*{pv("bank_available")}',None);text(s,23,0,'Condições precedentes — hipótese');fw(s,23,1,pv('bank_cp'),1,'int');text(s,24,0,'Documentos deCP / garantias');s.write_blank('B24',None,F['input_text']);text(s,26,0,'30M€é proteção de execução');text(s,27,0,'Não recomendar redução só pela referência')

s=sh('V04 Completion',['IndicadorM24','Projetado','Documentado / assegurado','Critério'],'FINANCIAL COMPLETION READINESS: seis meses antes deM30. Recursos projetados não são secured. Cobertura documental exige custos e calendário validados.')
c=completion(BASE);metricrows=[(5,'CaixaM24','projected_cash'),(6,'Recebimentos líquidos futuros','projected_receipts'),(7,'Capital futuroKK','projected_equity'),(8,'Banco remanescente mobilizável','projected_bank'),(9,'Liquidez total','projected_liquidity'),(11,'Cost to complete','cost_to_complete'),(12,'Encargos até abertura','charges_to_open'),(13,'Défices operacionais até abertura','opening_operating_need'),(14,'Reserva abertura','opening_reserve'),(15,'Necessidade total','total_uses'),(17,'Completion Funding Coverage','projected_coverage'),(18,'Funding gap','projected_gap')]
uses={5:"'Fluxo 180M'!AE28",6:f'SUM(\'Fluxo 180M\'!D29:INDEX(\'Fluxo 180M\'!D5:D184,{pv("open")}))-SUM(\'Fluxo 180M\'!E29:INDEX(\'Fluxo 180M\'!E5:E184,{pv("open")}))-SUM(\'Fluxo 180M\'!F29:INDEX(\'Fluxo 180M\'!F5:F184,{pv("open")}))-SUM(\'Fluxo 180M\'!G29:INDEX(\'Fluxo 180M\'!G5:G184,{pv("open")}))',7:f'SUM(\'Fluxo 180M\'!C29:INDEX(\'Fluxo 180M\'!C5:C184,{pv("open")}))',8:f'IF(AND({pv("bank_cp")}=1,24<{pv("draw_end")}),MAX(0,{pv("limit")}*{pv("bank_available")}-\'Fluxo 180M\'!AC28),0)',9:'SUM(B5:B8)',11:f'SUM(\'CAPEX\'!E29:INDEX(\'CAPEX\'!E5:E184,{pv("open")}))',12:'0',13:f'MAX(0,-SUM(\'Fluxo 180M\'!L29:INDEX(\'Fluxo 180M\'!L5:L184,{pv("open")}))-SUM(\'Fluxo 180M\'!M29:INDEX(\'Fluxo 180M\'!M5:M184,{pv("open")})))',14:f'INDEX(\'Fluxo 180M\'!Q5:Q184,{pv("open")})',15:'SUM(B11:B14)',17:'IF(B15>0,B9/B15,"")',18:'MAX(0,B15-B9)'}
chargecols=['interest','principal','openingfee','broker','mobil','loan_interest','loan_principal','commitment','hotpay'];uses[12]='+'.join(f'SUM(\'Fluxo 180M\'!{col(FI[k])}29:INDEX(\'Fluxo 180M\'!{col(FI[k])}5:{col(FI[k])}184,{pv("open")}))' for k in chargecols)
uses[8]=projected_bank_formula('Fluxo 180M','CAPEX',pv)
for r,label,key in metricrows:text(s,r,0,label);fw(s,r,1,uses[r],c[key],'mult' if r==17 else 'eur')
fw(s,8,1,projected_bank_formula('Fluxo 180M','CAPEX',pv),None)
for r,key in [(5,'completion_cash'),(6,'completion_receipts'),(7,'completion_equity'),(8,'completion_bank')]:fw(s,r,2,f'IF(ISNUMBER({rawpv(key)}),{rawpv(key)},"Indisponível")',None)
fw(s,9,2,'IF(COUNT(C5:C8)=4,SUM(C5:C8),"Indisponível")',None);fw(s,17,2,'IF(ISNUMBER(C9),C9/B15,"Indisponível")',None,'mult');fw(s,18,2,'IF(ISNUMBER(C9),MAX(0,B15-C9),"Indisponível")',None)
fw(s,21,1,f'IF(COUNT(C5:C8)<4,"Pendente — recursos não documentados",IF({pv("completion_cost_validated")}=0,"Pendente — custos não validados",IF({pv("completion_timing_validated")}=0,"Pendente — calendário não validado",IF(C17>=1,"Cobertura agregada documentada no cenário","Alerta — défice conclusão"))))',c['status'],'text');text(s,23,0,'Marco crítico');text(s,23,1,'M24; complementar teste mensal de liquidez eCP')
# Individual pricing: original prices are preserved unless Motor is explicitly selected.
s=sh('V04 Pricing',['ID','Tipologia','Piso','Área','Origem€','Base€/m²','Piso fator','Vista fator','Posição fator','Tipologia fator','Prestígio fator','Desconto','Absorção fator','Fator combinado','Preço/m² modelado','Preço modelado'],'94IDs preservados. Fatores neutros são hipóteses editáveis; não inventar vista/prestígio de cada unidade. Origem mantém preços daV45.')
for i,u in enumerate(UNITS):
 r=i+5
 for c,uc in [(0,0),(1,1),(2,2),(3,3),(4,6)]:fw(s,r,c,vr('Unidades',r,uc),u[['id','tipo','piso','area','price_source'][c]],'text' if c<3 else 'num' if c==3 else 'eur')
 fw(s,r,5,pv('price_m2'),DEFAULT['price_m2'])
 for c in [6,7,8,9,10,12]:text(s,r,c,1,'input_mult');s.data_validation(r-1,c,r-1,c,{'validate':'decimal','criteria':'between','minimum':0,'maximum':3})
 text(s,r,11,0,'input_pct');s.data_validation(r-1,11,r-1,11,{'validate':'decimal','criteria':'between','minimum':0,'maximum':1})
 fw(s,r,13,f'PRODUCT(G{r}:K{r})*(1-L{r})*M{r}',1,'mult');fw(s,r,14,f'F{r}*N{r}',DEFAULT['price_m2']);fw(s,r,15,f'ROUND(D{r}*O{r},2)',u['price'])
s.autofilter('A4:P98')

# A six-weight analytical policy:3PV target equations +2conservation equations.
def write_policy(name,rr,flow_sn,hotel_sn):
 p=rr['p'];pol=constant_policy(rr);ss=summary(rr);s=sh(name,['Categoria','Blocos','Base por bloco','Base classe','TIR-alvo','PV dividendos','PV saída','RácioPV','Base/PVsaída','Peso dividendos','Peso saída'],'Proposta não aprovada; bases dos promotores são convencionadas, não novas entradas. Pesos recalculam com fluxos, saídas e metas.');s.hide() if name!='V04 Politica' else None
 for r,label,n,tick,key in [(5,'Platinum',2,1e7,'target_platinum'),(6,'Gold',3,5e6,'target_gold'),(7,'Silver',15,1e6,'target_silver')]:
  text(s,r,0,label);text(s,r,1,n,'int');text(s,r,2,tick,'eur');fw(s,r,3,f'B{r}*C{r}',n*tick);fw(s,r,4,pv(key),p[key],'pct')
  fw(s,r,5,f'SUMPRODUCT(\'{flow_sn}\'!$AD$5:$AD$184,(1+E{r})^(-(\'{flow_sn}\'!$A$5:$A$184-1)/12))',None)
  exit=f'MAX(0,SUM(\'{hotel_sn}\'!T173:T184)*{pv("exit_mult")}*(1-{pv("exit_cost")})-\'{flow_sn}\'!AB184-\'{flow_sn}\'!AX184+\'{flow_sn}\'!AE184-\'{flow_sn}\'!AP184-\'{flow_sn}\'!AQ184)'
  fw(s,r,6,f'({exit})/(1+E{r})^(179/12)',None);fw(s,r,7,f'IF(G{r}>0,F{r}/G{r},0)',None,'mult');fw(s,r,8,f'IF(G{r}>0,D{r}/G{r},0)',None,'mult')
 text(s,11,0,'Valor de saída equity M180');fw(s,11,1,exit,ss['exit_equity']);text(s,12,0,'Denominador');fw(s,12,1,'H6-H7',None)
 text(s,13,0,'Alpha');fw(s,13,1,'IF(ABS(B12)>0.000000000001,-(H5-H7)/B12,0)',None)
 text(s,14,0,'Beta');fw(s,14,1,'IF(ABS(B12)>0.000000000001,(SUM(I5:I7)-1-H7)/B12,0)',None)
 for r,label,aa,bb in [(20,'P dividendos','1','0'),(21,'G dividendos','$B$13','$B$14'),(22,'S dividendos','-1-$B$13','1-$B$14'),(23,'P saída','-H5','I5'),(24,'G saída','-$B$13*H6','I6-$B$14*H6'),(25,'S saída','(1+$B$13)*H7','I7-(1-$B$14)*H7')]:
  text(s,r,0,label);fw(s,r,1,aa,None);fw(s,r,2,bb,None);fw(s,r,3,f'IF(B{r}>0.000000000001,-C{r}/B{r},IF(B{r}<-0.000000000001,(1-C{r})/B{r},0))',None);fw(s,r,4,f'IF(B{r}>0.000000000001,(1-C{r})/B{r},IF(B{r}<-0.000000000001,-C{r}/B{r},1))',None);fw(s,r,5,f'IF(AND(ABS(B{r})<0.000000000001,OR(C{r}<-0.00000001,C{r}>1.00000001)),1,0)',None,'int')
 for r,label,f in [(28,'Limite inferior','MAX(0,D20:D25)'),(29,'Limite superior','MIN(1,E20:E25)'),(30,'Metas exatas / pesos constantes','AND(MIN(G5:G7)>0,ABS(B12)>0.000000000001,B28<=B29+0.0000000001,SUM(F20:F25)=0)'),(31,'PesoPdiv selecionado',f'IF(B30,MAX(0,MIN(1,B28+(B29-B28)*{pv("allocation_preference")})),40%)')]:text(s,r,0,label);fw(s,r,1,f,pol.get('feasible') if r==30 else None,'text' if r==30 else 'num')
 for r,i in [(5,20),(6,21),(7,22)]:
  fw(s,r,9,f'IF($B$30,MAX(0,MIN(1,B{i}*$B$31+C{i})),D{r}/SUM(D5:D7))',None,'pct');fw(s,r,10,f'IF($B$30,MAX(0,MIN(1,B{i+3}*$B$31+C{i+3})),D{r}/SUM(D5:D7))',None,'pct')
 text(s,16,0,'Ativos líquidos para equity e crédito');text(s,17,0,'Crédito vendedores recuperável');text(s,18,0,'Crédito vendedores sem cobertura')
 fw(s,16,1,terminal_assets_formula(flow_sn,hotel_sn,pv),None);fw(s,17,1,f'MIN(\'{flow_sn}\'!AX184,B16)',None);fw(s,18,1,'MAX(0,\''+flow_sn+'\'!AX184-B17)',None)
 text(s,34,0,'Se não houver pesos constantes');text(s,34,1,'ModoOtimizado usa preferência combinada e alerta; análiseLP testa alocações variáveis por período. Não declarar impossibilidade geral só por falha deste método.')
 return name

WF_KEYS=['m','div_pool','exit_pool','balance_p','balance_g','balance_s','div_p','div_g','div_s','exit_p','exit_g','exit_s','close_p','close_g','close_s','check','new_p','new_g','new_s','new_received','new_outlay','new_cf','jscv_received','jals_received','promoter_cf','new_cum','promoter_cum']
def write_waterfall(name,policy_name,rr,flow_sn='Fluxo 180M',extra_funding_delay=0,hotel_sn='Hotel',capital_sn='V04 Capital180',pipeline_sn='V04 Pipeline'):
 s=sh(name,['Mês','Dividendos caixa','Saída estimada','SaldoP','SaldoG','SaldoS','DivP','DivG','DivS','SaídaP','SaídaG','SaídaS','SaldoPfinal','SaldoGfinal','SaldoSfinal','Check100%','P novos equivalentes pagos','G novos equivalentes pagos','S novos equivalentes pagos','Novos recebido','Novos capital pago','Fluxo novos','JSCV recebido','JALS recebido','Fluxo promotores','Novos acumulado','Promotores acumulado'],'Distribuições por bloco iguais dentro de cada categoria, incluindo promotores. Não são direitos aprovados. Bases10/5/1M€convencionadas; investimento histórico separado.');s.hide() if name!='V04 Waterfall180' else None
 pol=policy_name;P=lambda cell:f"'{pol}'!{cell}";mode=pv('waterfall_mode');mode=f'IF(AND({pv("capital_scenario")}<>"KK-E25",{pv("equity_mode")}<>"Combinação tickets"),"Proporcional",{mode})';targets=[pv('target_platinum'),pv('target_gold'),pv('target_silver')]
 for m in range(1,181):
  r=m+4;text(s,r,0,m,'int');fw(s,r,1,vr(flow_sn,r,FI['dividend']),rr['flow'][m-1]['dividend']);fw(s,r,2,f'IF(A{r}=180,{P("B11")},0)',None)
  for k in range(3):
   bal=f'{[20e6,15e6,15e6][k]}' if m==1 else f'{col(12+k)}{r-1}*(1+{targets[k]})^(1/12)';fw(s,r,3+k,bal,None)
  # Raw priority payments, then surplus shared in ordinary economic proportions.
  for c,label in [(28,'PprefD'),(29,'GprefD'),(30,'SprefD'),(31,'RestoD'),(32,'PprefE'),(33,'GprefE'),(34,'SprefE'),(35,'RestoE')]:
   if m==1:text(s,4,c,label,'head')
  for c,f in [(28,f'MIN(B{r},MAX(0,D{r}))'),(29,f'MIN(MAX(0,B{r}-AC{r}),MAX(0,E{r}))'),(30,f'MIN(MAX(0,B{r}-AC{r}-AD{r}),MAX(0,F{r}))'),(31,f'MAX(0,B{r}-SUM(AC{r}:AE{r}))')]:fw(s,r,c,f,None)
  for k,share in enumerate([.4,.3,.3]):
   prefer=f'{col(28+k)}{r}+AF{r}*{share}';f=f'IF(AND({mode}="Otimizado",{P("$B$30")}),B{r}*{P("$J$"+str(k+5))},IF(OR({mode}="Dividendos",{mode}="Combinado",{mode}="Otimizado"),{prefer},B{r}*{share}))';fw(s,r,6+k,f,None)
  for c,f in [(32,f'MIN(C{r},MAX(0,D{r}-G{r}))'),(33,f'MIN(MAX(0,C{r}-AG{r}),MAX(0,E{r}-H{r}))'),(34,f'MIN(MAX(0,C{r}-AG{r}-AH{r}),MAX(0,F{r}-I{r}))'),(35,f'MAX(0,C{r}-SUM(AG{r}:AI{r}))')]:fw(s,r,c,f,None)
  for k,share in enumerate([.4,.3,.3]):
   prefer=f'{col(32+k)}{r}+AJ{r}*{share}';fw(s,r,9+k,f'IF(AND({mode}="Otimizado",{P("$B$30")}),C{r}*{P("$K$"+str(k+5))},IF(OR({mode}="Saída",{mode}="Combinado",{mode}="Otimizado"),{prefer},C{r}*{share}))',None);fw(s,r,12+k,f'MAX(0,{col(3+k)}{r}-{col(6+k)}{r}-{col(9+k)}{r})',None)
  fw(s,r,15,f'SUM(G{r}:L{r})-B{r}-C{r}',0)
  legacy=f'AND({pv("capital_scenario")}<>"KK-E25",{pv("equity_mode")}<>"Combinação tickets")'
  for k,label in enumerate(['Platinum','Gold','Silver']):
   terms=[]
   for j in [0] if k==0 else [1] if k==1 else list(range(2,12)):
    pr=j+5;terms.append(f'IF(\'{pipeline_sn}\'!J{pr}<={m}-{extra_funding_delay},\'{pipeline_sn}\'!W{pr},0)')
   fw(s,r,16+k,f'SUM({",".join(terms)})',None)
  new=f'(G{r}+J{r})/2*Q{r}+(H{r}+K{r})/3*R{r}+(I{r}+L{r})/15*S{r}'
  ratio=f'IF({pv("equity")}>{pv("conversion")},SUMIFS(\'{capital_sn}\'!$D$5:$D$184,\'{capital_sn}\'!$A$5:$A$184,"<="&(A{r}-{extra_funding_delay}))/({pv("equity")}-{pv("conversion")})*\'Capital V03\'!B7,0)'
  fw(s,r,19,f'IF({legacy},(B{r}+C{r})*({ratio}),{new})',None)
  fw(s,r,20,f'IF(AND(A{r}-{extra_funding_delay}>=1,A{r}-{extra_funding_delay}<=180),INDEX(\'{capital_sn}\'!$D$5:$D$184,A{r}-{extra_funding_delay}),0)',investor_outlay(rr['p'],m));fw(s,r,21,f'T{r}-U{r}',None)
  allP=f'(G{r}+J{r})/2';allG=f'(H{r}+K{r})/3';allS=f'(I{r}+L{r})/15'
  jscv=f'({allP})*(1+0.55*(1-Q{r}))+({allG})*2';jals=f'({allP})*0.45*(1-Q{r})+({allG})*(1-R{r})+({allS})*(15-S{r})'
  fw(s,r,22,f'IF({legacy},(B{r}+C{r})*\'Capital V03\'!B8,{jscv})',None);fw(s,r,23,f'B{r}+C{r}-T{r}-W{r}',None)
  retained=f'U{r}*(1-{rawpv("treasury_ratio")})';credit=f"'{flow_sn}'!AV{r}+'{flow_sn}'!AW{r}+IF(A{r}=180,'{policy_name}'!B17,0)";fw(s,r,24,f'W{r}+X{r}+{retained}+{credit}',None)
  fw(s,r,25,f'SUM(V$5:V{r})',None);fw(s,r,26,f'SUM(Y$5:Y{r})-{pv("historic")}',None)
 s.set_column('AC:AJ',None,None,{'hidden':True});return name

write_policy('V04 Politica',BASE,'Fluxo 180M','Hotel');write_waterfall('V04 Waterfall180','V04 Politica',BASE)
s=sh('V04 Retornos',['Categoria','Base convencional','Meta15A','TIR convencional15A','Desvio p.p.','Dividendos por bloco','Saída por bloco','MOIC convencional','VPL convencional12%','TIR mínimo novos','TIR máximo novos'],'Bases dos promotores não são dinheiro investido. TIR efetiva de cada novo bloco usa o calendário de pagamento; data diferente pode gerarTIRdiferente com igual direito económico por ação.')
v=sh('V04 Vet Retorno',['Mês','P convencional','G convencional','S convencional','Novos efetivo','JSCV recebidos','JALS recebidos','Promotores proxy'],'Vetores mensais0..180; histórico4,5M€grupo noM0éproxy, não há base individual documentada.');v.hide()
for m in range(181):
 r=m+5;fr=m+4;text(v,r,0,m,'int')
 for k,(ticket,count) in enumerate([(1e7,2),(5e6,3),(1e6,15)]):fw(v,r,k+1,'0' if m==0 else f'(\'V04 Waterfall180\'!{col(6+k)}{fr}+\'V04 Waterfall180\'!{col(9+k)}{fr})/{count}'+(f'-{ticket}' if m==1 else ''),None)
 for c,f in [(4,'0' if m==0 else f"'V04 Waterfall180'!V{fr}"),(5,'0' if m==0 else f"'V04 Waterfall180'!W{fr}"),(6,'0' if m==0 else f"'V04 Waterfall180'!X{fr}"),(7,f'-{pv("historic")}' if m==0 else f"'V04 Waterfall180'!Y{fr}")]:fw(v,r,c,f,None)
 for i,plan in enumerate(pipeline(BASE['p'])):
  c=i+8;catk=0 if i==0 else 1 if i==1 else 2;pr=i+5;count=[2,3,15][catk]
  outlay=f'IF(A{r}=\'V04 Pipeline\'!J{pr},\'V04 Pipeline\'!T{pr},0)';inc=f'IF(A{r}>=\'V04 Pipeline\'!J{pr},(\'V04 Waterfall180\'!{col(6+catk)}{fr}+\'V04 Waterfall180\'!{col(9+catk)}{fr})/{count}*\'V04 Pipeline\'!W{pr},0)'
  fw(v,r,c,'0' if m==0 else inc+'-'+outlay,None)
for k,label in enumerate(['Platinum','Gold','Silver']):
 r=k+5;vc=col(k+1);count=[2,3,15][k];ticket=[1e7,5e6,1e6][k];text(s,r,0,label);text(s,r,1,ticket,'eur');fw(s,r,2,pv(['target_platinum','target_gold','target_silver'][k]),None,'pct');fw(s,r,3,f'IFERROR((1+IRR(\'V04 Vet Retorno\'!{vc}5:{vc}185,0.005))^12-1,"TIR não definida")',None,'pct');fw(s,r,4,f'IF(ISNUMBER(D{r}),D{r}-C{r},"Indisponível")',None,'pct');fw(s,r,5,f'SUM(\'V04 Waterfall180\'!{col(6+k)}5:{col(6+k)}184)/{count}',None);fw(s,r,6,f'SUM(\'V04 Waterfall180\'!{col(9+k)}5:{col(9+k)}184)/{count}',None);fw(s,r,7,f'(F{r}+G{r})/B{r}',None,'mult');fw(s,r,8,f'NPV((1+{pv("equity_discount")})^(1/12)-1,\'V04 Vet Retorno\'!{vc}6:{vc}185)',None)
 for j in [0] if k==0 else [1] if k==1 else list(range(2,12)):
  rr=12+j;cc=col(j+8);text(s,rr,0,pipeline(BASE['p'])[j]['code']);text(s,rr,1,label);fw(s,rr,2,f"'V04 Pipeline'!T{j+5}",None);fw(s,rr,3,f'IF(C{rr}>0,IFERROR((1+IRR(\'V04 Vet Retorno\'!{cc}5:{cc}185,0.005))^12-1,"TIR não definida"),"")',None,'pct')
 rang=[12] if k==0 else [13] if k==1 else list(range(14,24));fw(s,r,9,f'IF(COUNT(D{rang[0]}:D{rang[-1]})>0,MIN(D{rang[0]}:D{rang[-1]}),"")',None,'pct');fw(s,r,10,f'IF(COUNT(D{rang[0]}:D{rang[-1]})>0,MAX(D{rang[0]}:D{rang[-1]}),"")',None,'pct')
for r,label,expr,fmt in [(27,'TIR novos efetiva15A','IFERROR((1+IRR(\'V04 Vet Retorno\'!E5:E185,0.005))^12-1,"TIR não definida")','pct'),(28,'MOIC novos efetivo','IF(SUM(\'V04 Pipeline\'!T5:T16)>0,SUM(\'V04 Waterfall180\'!T5:T184)/SUM(\'V04 Pipeline\'!T5:T16),"")','mult'),(29,'VPL novos12%','NPV((1+'+pv('equity_discount')+')^(1/12)-1,\'V04 Vet Retorno\'!E6:E185)','eur'),(30,'JSCV recebido15A',"SUM('V04 Vet Retorno'!F5:F185)",'eur'),(31,'JALS recebido15A',"SUM('V04 Vet Retorno'!G5:G185)",'eur'),(32,'TIR Grupo promotor — proxy','IFERROR((1+IRR(\'V04 Vet Retorno\'!H5:H185,0.005))^12-1,"TIR não definida")','pct'),(33,'Reconciliação distribuições',"SUM('V04 Waterfall180'!T5:T184,'V04 Waterfall180'!W5:X184)-SUM('V04 Waterfall180'!B5:C184)",'eur')]:text(s,r,0,label);fw(s,r,1,expr,None,fmt)
# Current executive KPIs use the selected waterfall, not the historical pro-rata benchmark.
for key,addr in [('equityirr','B27'),('moic','B28')]:
 if key in KROW:fw(S['KPI'],KROW[key],1,"'V04 Retornos'!"+addr,None,'pct' if key=='equityirr' else 'mult')

for item in SCMAP:
 if 'delay' in item['scenario']:
  savedpv=pv;extra_delay_legacy=item['scenario']['delay']
  def pv(k):return f'({rawpv(k)}+{extra_delay_legacy})' if k=='commercial_delay' else savedpv(k)
  try:write_case_re('LEG'+str(item['index']),SCRESULT[item['index']],f"CalcR{item['index']+1:02d}")
  finally:pv=savedpv

from returns10 import returns10_updates
for addr,value in returns10_updates(pv).items():
 kind,f,fmt=value
 if kind=='text':S['V04 Retornos'].write(addr,f,F['text'])
 else:S['V04 Retornos'].write_formula(addr,'='+f,F[fmt],'')
S['V04 Retornos'].set_column('M:Z',None,None,{'hidden':True})
