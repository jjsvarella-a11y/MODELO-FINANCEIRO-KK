# Executed by the existing builder before workbook.close(). Additive V02 UI and reporting.
from PIL import Image,ImageDraw,ImageFont

# Capital purchase price vs project treasury. These values must never be conflated.
s=sh('Ronda Investidores',['Categoria','Ticket referência','Participação ref.20M','Quantidade selecionada','Capital adquirido','Participação na ronda','Ações equivalentes'],'20 M€ por 40.000 ações da JALS. Combinações são hipóteses; benefícios não acrescentam ações, votos ou obrigações.')
s.set_column('A:A',37);s.set_column('B:G',22)
classes=[('Platinum',10000000,'n_platinum'),('Gold',5000000,'n_gold'),('Silver Plus',2000000,'n_silver_plus'),('Silver',1000000,'n_silver')]
for r,(label,ticket,key) in enumerate(classes,5):
 text(s,r,0,label);text(s,r,1,ticket,'eur');text(s,r,2,ticket/20000000*.4,'pct');fw(s,r,3,rawpv(key),DEFAULT[key],'int');fw(s,r,4,f'B{r}*D{r}',ticket*DEFAULT[key]);fw(s,r,5,f'E{r}/{pv("equity_price")}*{pv("equity_stake")}',ticket*DEFAULT[key]/DEFAULT['equity_price']*DEFAULT['equity_stake'],'pct');fw(s,r,6,f'F{r}*{pv("shares_total")}',ticket*DEFAULT[key]/DEFAULT['equity_price']*DEFAULT['equity_stake']*DEFAULT['shares_total'],'num')
for c in [4,5,6]:fw(s,10,c,f'SUM({col(c)}5:{col(c)}8)',20000000 if c==4 else .4 if c==5 else 40000,'total_eur' if c==4 else 'total_pct' if c==5 else 'total_num')
text(s,12,0,'Alterar quantidades');text(s,12,1,'Pressupostos → n.º por categoria');text(s,13,0,'Modo de captação');fw(s,13,1,rawpv('equity_mode'),DEFAULT['equity_mode'],'text')
text(s,20,0,'Preço efetivo de aquisição');fw(s,20,1,f'IF({rawpv("equity_mode")}="Combinação tickets",SUM(E5:E8),{rawpv("equity")})',DEFAULT['equity'])
text(s,21,0,'Parcela transferida à KK');fw(s,21,1,pv('treasury_ratio'),1,'pct');text(s,22,0,'Caixa que entra na KK');fw(s,22,1,f'(B20-{pv("conversion")})*B21',DEFAULT['equity'])
text(s,23,0,'Retido na JALS / vendedor');fw(s,23,1,f'(B20-{pv("conversion")} )*(1-B21)',0)
text(s,24,0,'Novo capital / participação');fw(s,24,1,f'(B20-{pv("conversion")})/{pv("equity_price")}*{pv("equity_stake")}',.4,'pct')
text(s,25,0,'Estrutura jurídica');fw(s,25,1,f'IF({pv("legal_confirmed")}=1,"Formalização assinalada — rever documento","Pendente de Validação Externa")','Pendente de Validação Externa','text');alert(s,'B25')
text(s,27,0,'Limites da combinação');fw(s,27,1,'IF(AND(E10<=20000000,F10<=40%),"Conforme hipótese","Alerta — excede 20 M€ / 40%")','Conforme hipótese','text')
s.merge_range('A29:G31','Compra de ações existentes não é aumento de capital. O cenário assume que o preço chega integralmente à KK por instrumento irrevogável sem obrigação de reembolso, ainda por formalizar. Se for suprimento reembolsável, deve ser modelado como dívida adicional antes de validar a viabilidade.',F['warn'])
text(s,34,0,'Blocos de benefícios JSCV');text(s,34,1,'2 Platinum + 1 Silver');text(s,35,0,'Blocos de benefícios JALS');text(s,35,1,'1 Gold + 2 Silver Plus + 2 Silver');s.merge_range('A37:G39','Benefícios sujeitos ao Acordo de Acionistas e Estatutos. Estes blocos não são quantidades de ações, preços de aquisição, votos adicionais, dividendos preferenciais ou garantias.',F['warn'])

# Operator conditions: every column has an explicit contractual status.
s=sh('Condicoes Operador',['Componente','Proposta / simulação','Objetivo KK','Limite a definir','Personalizado','Selecionado'],'Proposta não recebida por defeito. Coluna B é hipótese 4% / 3% / 10%, não condição contratual. Limites não fornecidos ficam vazios.')
for r,key,custom,label in [(5,'base','base_fee','Fee base'),(6,'systems','systems','Distribuição / sistemas'),(7,'incentive','incentive','Fee incentivo')]:
 text(s,r,0,label);fw(s,r,1,rawpv('offer_'+key),DEFAULT['offer_'+key],'pct');fw(s,r,2,rawpv('target_'+key),DEFAULT['target_'+key],'pct');fw(s,r,3,f'IF(ISNUMBER({rawpv("limit_"+key)}),{rawpv("limit_"+key)},"")',None,'pct');fw(s,r,4,rawpv(custom),DEFAULT[custom],'pct')
 fw(s,r,5,f'IF({pv("fee_mode")}="Objetivo KK",C{r},IF({pv("fee_mode")}="Proposta simulada",B{r},IF({pv("fee_mode")}="Limites propostos",IF(ISNUMBER(D{r}),D{r},C{r}),E{r})))',BASE['p'][custom],'pct');s.write_comment(r-1,3,'Limite negocial ainda não fornecido. Preencher apenas como decisão explícita; não é HMA aprovado.')
text(s,9,0,'Modo ativo');fw(s,9,1,pv('fee_mode'),DEFAULT['fee_mode'],'text');text(s,10,0,'Proposta recebida?');fw(s,10,1,f'IF({pv("offer_received")}=1,"Documento assinalado — validar","Não recebida: simulação")','Não recebida: simulação','text')
text(s,11,0,'HMA aprovado?');text(s,11,1,'Não — pendente de validação','input_text');text(s,12,0,'Alerta total / CFADS');fw(s,12,1,pv('total_alert'),.5,'pct')
NEG=[('Performance test','Definir períodos, testes cumulativos e remédios; benchmark não fornecido.'),('GOP mínimo','Preencher pressuposto gop_floor; 0 significa limiar não definido.'),('EBITDA mínimo','Base pré-incentivo anualizada para evitar circularidade; limiar não definido.'),('Incentivo condicionado','Seletor incentive_gate; hipótese desativada por defeito.'),('FF&E','Responsabilidades, dotação, reposição e titularidade do fundo por acordar.'),('CAPEX manutenção','Definir aprovação, orçamento e despesas extraordinárias sem duplicar FF&E.'),('Aprovação orçamental','Calendário e matérias reservadas ao proprietário por negociar.'),('Reporting e auditoria','Periodicidade, plano de contas, direito de inspeção e auditor independente.'),('Distribuição e sistemas','Inventariar bases, encargos pass-through e sobreposições; objetivo 3%.'),('Rescisão e proteção','Performance, prazo, cura, indemnizações e continuidade operacional por acordar.')]
s.write_row(14,0,['Matéria','Objetivo / campo negocial','Estado documental'],F['head'])
for r,(label,note) in enumerate(NEG,16):text(s,r,0,label);text(s,r,1,note,'input_text');text(s,r,2,'Pendente de Validação Externa','input_text')
s.set_column('A:A',37);s.set_column('B:B',73);s.set_column('C:F',25);alert(s,'C16:C25')

# Correct the societary bridge, leaving prior worksheets and functionality in place.
s=S['Societario'];fw(s,5,1,pv('legacy_jscv'),.51,'pct');fw(s,5,2,pv('legacy_jscv'),.51,'pct');fw(s,6,1,f'1-{pv("legacy_jscv")}',.49,'pct');fw(s,6,2,f'1-{pv("legacy_jscv")}-\'Retornos\'!C13',.09,'pct');text(s,5,3,'JSCV mantém 51%; não há diluição proporcional presumida.');text(s,6,3,'JALS cede 40.000 ações; fica com 9% na referência 20 M€ / 40%.');text(s,7,3,'Novos investidores 40%; preço de aquisição não é entrada automática na KK.')
text(s,18,0,'Ações existentes');fw(s,18,1,pv('shares_total'),100000,'int')
for r,cshare in [(19,'C5'),(20,'C6'),(21,'C7')]:text(s,r,0,['JSCV ações pós','JALS ações pós','Investidores ações'][r-19]);fw(s,r,1,f'{cshare}*{pv("shares_total")}',[51000,9000,40000][r-19],'num')
s=S['Retornos']
for c in [1,2]:
 C=col(c);fw(s,21,c,f'({C}11+{C}12)*{C}20+({pv("equity")}-{pv("conversion")})*(1-{pv("treasury_ratio")})',RET[120 if c==1 else 180]['promoter_total'])
text(s,27,3,'JSCV mantém 51%. Estrutura proposta depende de formalização societária e da entrada efetiva de fundos.');text(s,28,3,'JALS = 49% menos participação cedida; referência 9%. Não se criam direitos adicionais dos blocos de benefícios.')
s=S['Fluxos Retorno']
for c in [3,4]:
 rc='B' if c==3 else 'C';fw(s,6,c,f'\'Anual\'!V5*\'Retornos\'!{rc}20+({pv("equity")}-{pv("conversion")})*(1-{pv("treasury_ratio")})',0)

# Complete scenario metrics from original CalcH/CalcF blocks. Return vectors remain auditable.
RESHEAD=['Cenário','Preço aquisição','Entrada na KK','Preço por 40%','Participação novos','Usado acumulado','Pico dívida','Juros','Comissões','Caixa mínima obra','Caixa mínima 180M','Falta reserva máx.','Défice não financiado','Caixa M30','Reserva M30','Dívida M30','Dívida abertura','DSCR ano parcial','DSCR anos completos','Dividendos 15A','TIR novos 15A','MOIC 15A','VPL 15A','Total promotores','TIR promotores proxy','Saída equity 15A','Estado','MOIC 10A','TIR novos 10A','VPL 10A','Dividendos 10A','Necessidade 4T','Cobertura equity 4T']
resheet=sh('Resultados Cenarios',RESHEAD,'Todas as métricas provêm dos blocos do motor original; valor de saída é estimado. Retorno condicionado a liquidez e transferência jurídica dos fundos.')
resheet.set_column('A:A',38);vec=sh('Vet Cenarios',['Mês'],'Fluxos investidor/promotor por cenário e horizonte.');vec.hide()
for m in range(181):text(vec,m+5,0,m,'int')
SCMAP=[]
for si,(sc,rr) in enumerate(zip(SCEN,SCRESULT)):
 r=si+5;fsn=f'CalcF{si+1:02d}';hsn=f'CalcH{si+1:02d}';Fv=lambda k,j:vr(fsn,j,FI[k]);R=lambda c:col(c)+str(r)
 text(resheet,r,0,sc['label']+(' / '+str(int(sc['equity']/1e6))+'M' if sc['equity'] is not None else ' / ativo'))
 eqexpr=vr('Comparacao',r,0)+'*1000000';price=str(sc['equity_price']) if 'equity_price' in sc else pv('equity_price')
 vals=summary(rr);C=lambda formula:fw(resheet,r,C.i,formula);fw(resheet,r,1,f'{eqexpr}-{pv("conversion")}',vals['equity']);fw(resheet,r,2,Fv('equity',5),rr['flow'][0]['equity']);fw(resheet,r,3,price,rr['p']['equity_price']);fw(resheet,r,4,f'B{r}/D{r}*{sc.get("equity_stake", .4) if sc.get("equity") is not None else pv("equity_stake")}',vals['share'],'pct')
 for c,key,formula in [(5,None,f'MAX(\'{fsn}\'!AC5:AC184)'),(6,'peak',vr('Comparacao',r,3)),(7,'interest',vr('Comparacao',r,9)),(8,'fees',f'SUM(\'{fsn}\'!T5:V184)'),(9,None,f'MIN(\'{fsn}\'!AE5:INDEX(\'{fsn}\'!AE5:AE184,{pv("end")}))'),(10,'cash_min',vr('Comparacao',r,6)),(11,'reserve_gap',vr('Comparacao',r,7)),(12,'funding_gap',vr('Comparacao',r,8)),(13,'cash30',Fv('cash',34)),(14,None,Fv('reserve',34)),(15,'debt30',Fv('debt',34)),(16,None,f'INDEX(\'{fsn}\'!O5:O184,{pv("open")})')]:fw(resheet,r,c,formula,vals.get(key) if key else None)
 partialyear=f'ROUNDUP({pv("open")}/12,0)';service=f'SUMIF(\'{fsn}\'!$B$5:$B$184,{partialyear},\'{fsn}\'!$AH$5:$AH$184)';cf=f'SUMIF(\'{fsn}\'!$B$5:$B$184,{partialyear},\'{fsn}\'!$L$5:$L$184)';fw(resheet,r,17,f'IF({service}>0,{cf}/{service},"")',vals['min_dscr'],'mult')
 terms=[];count=[]
 for y in range(1,16):
  st=5+(y-1)*12;en=st+11;ss=f'SUM(\'{fsn}\'!AH{st}:AH{en})';cond=f'AND({st-4}>={pv("open")},{ss}>0)';terms.append(f'IF({cond},SUM(\'{fsn}\'!L{st}:L{en})/{ss},1E99)');count.append(f'IF({cond},1,0)')
 fw(resheet,r,18,f'IF(SUM({",".join(count)})>0,MIN({",".join(terms)}),"")',None,'mult')
 exitexpr={}
 for h in [120,180]:
  en=h+4;st=en-11;ev=f'SUM(\'{hsn}\'!T{st}:T{en})*{pv("exit_mult")}';exitexpr[h]=f'MAX(0,({ev})*(1-{pv("exit_cost")})-{Fv("debt",en)}+{Fv("cash",en)}-{Fv("hothold",en)}-{Fv("conthold",en)})'
 for c,h in [(19,180),(30,120)]:fw(resheet,r,c,f'SUM(\'{fsn}\'!AD5:AD{h+4})',summary(rr,h)['dividends'])
 fw(resheet,r,25,exitexpr[180],vals['exit_equity']);fw(resheet,r,23,f'(T{r}+Z{r})*(1-E{r})+B{r}*(1-{pv("treasury_ratio")})',None)
 # Four columns each scenario: investor 10 / 15 and promoter 10 / 15.
 vc=si*4+1
 for m in range(181):
  mr=m+5;fr=m+4
  for j,(h,promoter) in enumerate([(120,False),(180,False),(120,True),(180,True)]):
   share=f'(1-\'Resultados Cenarios\'!$E${r})' if promoter else f'\'Resultados Cenarios\'!$E${r}'
   formula=f'-{pv("historic")}' if m==0 and promoter else '0' if m==0 or m>h else (f'{Fv("dividend",fr)}*{share}' if promoter else f'{Fv("dividend",fr)}*{share}-IF(A{mr}=1,\'Resultados Cenarios\'!$B${r},0)')
   if m==1 and promoter:formula+=f'+\'Resultados Cenarios\'!$B${r}*(1-{pv("treasury_ratio")})'
   if m==h:formula+=f'+({exitexpr[h]})*{share}'
   fw(vec,mr,vc+j,formula,None)
 for h,ic,mc,nc,j in [(180,20,21,22,1),(120,28,27,29,0)]:
  vv=f"'Vet Cenarios'!{col(vc+j)}5:{col(vc+j)}{h+5}";nv=f"'Vet Cenarios'!{col(vc+j)}6:{col(vc+j)}{h+5}"
  fw(resheet,r,ic,f'IF(B{r}>0,IFERROR((1+IRR({vv},0.005))^12-1,"TIR não definida"),"")',summary(rr,h)['equity_irr'],'pct')
  fw(resheet,r,mc,f'IF(B{r}>0,(SUM(\'{fsn}\'!AD5:AD{h+4})+({exitexpr[h]}))*E{r}/B{r},"")',summary(rr,h)['moic'],'mult')
  fw(resheet,r,nc,f'NPV((1+{pv("equity_discount")})^(1/12)-1,{nv})',None)
 fw(resheet,r,24,f'IFERROR((1+IRR(\'Vet Cenarios\'!{col(vc+3)}5:{col(vc+3)}185,0.005))^12-1,"TIR não definida")',None,'pct')
 fw(resheet,r,26,f'IF(OR(L{r}>0.01,M{r}>0.01,E{r}>{pv("equity_stake")}+0.000001),"Alerta","Hipótese condicionada à formalização")',None,'text')
 # Four-quarter uses include all charges actually modeled; no invented actuals.
 need=f'SUM(\'{fsn}\'!E5:K16)+SUM(\'{fsn}\'!T5:V16)+SUM(\'{fsn}\'!Y5:Y16)';fw(resheet,r,31,need,None);fw(resheet,r,32,f'IF(AF{r}>0,C{r}/AF{r},"")',None,'mult')
 SCMAP.append(dict(index=si,row=r,flow=fsn,hotel=hsn,scenario=sc))
resheet.autofilter(3,0,len(SCEN)+3,len(RESHEAD)-1);alert(resheet,f'AA5:AA{len(SCEN)+4}')

# Strategic 15 M€ for40 vs 20 M€ for40, calculated under identical common assumptions.
PAIR={int(sc['equity']/1e6):i for i,sc in enumerate(SCEN) if sc.get('label') in ['15M por 40%','20M por 40%']}
s=sh('Comparativo 15 20',['Indicador','15 M€ / 40%','20 M€ / 40%','Diferença 20−15','Leitura'],'Mesma operação e condições bancárias. Redução de risco não equivale a maior TIR pelo mesmo stake. Ambos os montantes de caixa estão condicionados à transferência jurídica.')
s.set_column('A:A',43);s.set_column('B:D',23);s.set_column('E:E',45)
CP=[('Preço de aquisição',1,'eur'),('Participação novos',4,'pct'),('Entrada efetiva na KK',2,'eur'),('Utilização acumulada linha',5,'eur'),('Pico da dívida',6,'eur'),('Juros 180M',7,'eur'),('Comissões',8,'eur'),('Caixa mínima construção',9,'eur'),('Caixa mínima total',10,'eur'),('Falta máxima reserva',11,'eur'),('Défice não financiado',12,'eur'),('Caixa mês 30',13,'eur'),('Reserva protegida mês 30',14,'eur'),('Dívida mês 30 / abertura',15,'eur'),('DSCR ano parcial',17,'mult'),('DSCR anos completos',18,'mult'),('Dividendos totais 15A',19,'eur'),('TIR investidores 15A',20,'pct'),('MOIC investidores 15A',21,'mult'),('VPL investidores 15A',22,'eur'),('Total promotores 15A',23,'eur'),('TIR promotores proxy',24,'pct'),('MOIC investidores 10A',27,'mult'),('TIR investidores 10A',28,'pct'),('VPL investidores 10A',29,'eur'),('Necessidades 4 primeiros T',31,'eur'),('Equity / necessidades 4T',32,'mult')]
for r,(label,c,fmt) in enumerate(CP,5):
 text(s,r,0,label)
 for cc,eq in [(1,15),(2,20)]:fw(s,r,cc,vr('Resultados Cenarios',PAIR[eq]+5,c),None,fmt)
 fw(s,r,3,f'IF(AND(ISNUMBER(B{r}),ISNUMBER(C{r})),C{r}-B{r},"")',None,fmt)
 if c in [20,21,22]:text(s,r,4,'Maior preço pelos mesmos 40% pode reduzir retorno.');s.set_row(r-1,25)
text(s,34,0,'Hipótese jurídica');text(s,34,1,'Pendente de formalização');text(s,35,0,'Sem serviço da dívida');text(s,35,1,'DSCR fica vazio; não é zero nem 999.')
s.write_row(38,0,['Trimestre','Usos 15M','Equity+vendas 15M','Cobertura 15M','Usos 20M','Equity+vendas 20M','Cobertura 20M'],F['head'])
for q in range(1,5):
 r=39+q;text(s,r,0,q,'int');st=5+(q-1)*3;en=st+2
 for ec,rc,eq in [(1,2,15),(4,5,20)]:
  fn=SCMAP[PAIR[eq]]['flow'];uses=f'SUM(\'{fn}\'!E{st}:K{en})+SUM(\'{fn}\'!T{st}:V{en})+SUM(\'{fn}\'!Y{st}:Y{en})';resources=f'SUM(\'{fn}\'!C{st}:D{en})'
  fw(s,r,ec,uses,None);fw(s,r,rc,resources,None);fw(s,r,rc+1,f'IF({col(ec)}{r}>0,{col(rc)}{r}/{col(ec)}{r},"")',None,'mult')
for m in range(1,181):
 r=m+60;text(s,r,0,m,'int')
 for c,eq in [(1,15),(2,20)]:fw(s,r,c,vr(SCMAP[PAIR[eq]]['flow'],m+4,FI['debt']),None)
chart(s,'G5','line','Dívida mensal: 15 M€ vs 20 M€',['Comparativo 15 20',60,0,239,0],[('15 M€ / 40%',['Comparativo 15 20',60,1,239,1]),('20 M€ / 40%',['Comparativo 15 20',60,2,239,2])],620,300)
s.print_area('A1:O44')

# Execution data deliberately blank, not invented from forecasts.
s=sh('Execucao',['Mês','CAPEX previsto','Comprometido efetivo €','Realizado efetivo €','Desvio vs previsto'],'Campos efetivos vazios. Preencher só com informação documental; vazio não significa execução zero. Mês de corte em Pressupostos.')
for m in range(1,181):
 r=m+4;text(s,r,0,m,'int');fw(s,r,1,vr('CAPEX',r,4),BASE['cap'][m-1]['total']);s.write_blank(r-1,2,None,F['input_eur']);s.write_blank(r-1,3,None,F['input_eur']);fw(s,r,4,f'IF(ISNUMBER(D{r}),D{r}-B{r},"")',None)
s.autofilter('A4:E184')

# Four new specialized panels retain all original dashboards and filters.
s=dash('DASH Investimento','A | Investimento e CAPEX — previsão identificada; comprometido e realizado apenas quando informados.')
for r,label,f,fmt in [(5,'Orçamento futuro base',f'{pv("area")}*{pv("cost")}*(1+{pv("supervision")}+{pv("contingency")})','eur'),(6,'Estimativa à conclusão',"SUM('CAPEX'!E5:E184)",'eur'),(7,'CAPEX previsto €/m²',f'B6/{pv("area")}', 'eur'),(8,'Histórico já investido',pv('historic'),'eur'),(9,'Investimento total económico','B6+B8','eur'),(10,'Sobrecusto adicional',pv('cost_overrun'),'pct'),(11,'Contingência orçamentada',f'{pv("area")}*{pv("cost")}*(1+{pv("cost_overrun")} )*{pv("contingency")}', 'eur'),(12,'Contingência realizada prevista',"SUM('CAPEX'!D5:D184)",'eur'),(13,'Mês de referência',pv('actual_month'),'int'),(14,'Comprometido informado',f'IF(COUNT(\'Execucao\'!C5:C184)>0,SUMIF(\'Execucao\'!A5:A184,"<="&{pv("actual_month")},\'Execucao\'!C5:C184),"Indisponível")','eur'),(15,'Realizado informado',f'IF(COUNT(\'Execucao\'!D5:D184)>0,SUMIF(\'Execucao\'!A5:A184,"<="&{pv("actual_month")},\'Execucao\'!D5:D184),"Indisponível")','eur'),(16,'Por executar estimado','IF(ISNUMBER(B15),MAX(0,B6-B15),"Indisponível")','eur'),(17,'Desvio EAC / orçamento','B6-B5','eur')]:metric(s,r,label,f,fmt=fmt)
chart(s,'D5','column','Previsão CAPEX mensal',['CAPEX',4,0,33,0],[('Construção',['CAPEX',4,1,33,1]),('Fiscalização',['CAPEX',4,2,33,2]),('Contingência paga',['CAPEX',4,3,33,3])])

s=dash('DASH Residencial','F | Pool das 94 frações — renda-alvo não garantida, FM e insuficiências visíveis; fundos fora do caixa distribuível KK.')
for r,label,formula,fmt in [(5,'ADR início',pv('res_adr'),'eur'),(6,'Ocupação início',pv('res_occ'),'pct'),(7,'Receitas pool 180M',"SUM('Residencial'!D5:D184)",'eur'),(8,'OPEX pool 180M',"SUM('Residencial'!E5:E184)",'eur'),(9,'Renda-alvo 180M',"SUM('Residencial'!L5:L184)",'eur'),(10,'Renda suportável 180M',"SUM('Residencial'!M5:M184)",'eur'),(11,'Renda não coberta',"SUM('Residencial'!N5:N184)",'eur'),(12,'Yield efetivo médio anual',"B10/SUM('Unidades'!H5:H98)/((180-'Pressupostos'!B"+str(PR['open'])+"+1)/12)",'pct'),(13,'Fundos FF&E + adicional',"'Residencial'!V184",'eur'),(14,'Residual KK 180M',"SUM('Residencial'!U5:U184)",'eur'),(15,'Estado arranque','IF(B11>0.01,"Alerta — yield não financiado","Conforme cenário")','text')]:metric(s,r,label,formula,fmt=fmt)
chart(s,'D5','line','Renda alvo e suportável por mês',['Residencial',4,0,183,0],[('Alvo',['Residencial',4,11,183,11]),('Suportável',['Residencial',4,12,183,12])]);alert(s,'B15')

s=dash('DASH Operador','E | Negociação com o Operador Internacional — proposta e limites não contratados; impactos calculados por cenário completo.')
metric(s,5,'Condições selecionadas',pv('fee_mode'),fmt='text')
for r,key,label in [(6,'base_fee','Fee base'),(7,'systems','Distribuição / sistemas'),(8,'incentive','Fee incentivo')]:metric(s,r,label,pv(key),fmt='pct')
metric(s,9,'HMA / estado',"'Condicoes Operador'!B11",fmt='text');metric(s,10,'Gestão / CFADS ref.',"'DASH Hotel MICE'!B14",fmt='pct');metric(s,11,'Sistemas / CFADS ref.',"'DASH Hotel MICE'!B15",fmt='pct');metric(s,12,'Total / CFADS ref.',"'DASH Hotel MICE'!B16",fmt='pct');metric(s,13,'Estado encargos',"'DASH Hotel MICE'!B17",fmt='text');metric(s,14,'Condicionar incentivo',pv('incentive_gate'),fmt='int');metric(s,15,'GOP mínimo anualizado',pv('gop_floor'));metric(s,16,'Resultado pré-incentivo mín.',pv('ebitda_floor'))
s.write_row(21,0,['Condição simulada','Fee base','Sistemas','Incentivo','EBITDA 15A','CFADS 15A','DSCR parcial','Dividendos 15A','TIR novos 15A','Estado'],F['head'])
for r,kind in enumerate(['target','offer','limit','custom'],23):
 si=next(i for i,sc in enumerate(SCEN) if sc.get('fee_kind')==kind);sc=SCMAP[si];hsn=sc['hotel'];frn=sc['flow'];sr=sc['row'];text(s,r,0,{'target':'Objetivo KK','offer':'Proposta simulada*','limit':'Limites propostos*','custom':'Personalizado'}[kind]);c={'target':2,'offer':1,'limit':3,'custom':4}[kind]
 for cc,rr in [(1,5),(2,6),(3,7)]:fw(s,r,cc,f'IF(ISNUMBER({vr("Condicoes Operador",rr,c)}),{vr("Condicoes Operador",rr,c)},{vr("Condicoes Operador",rr,2)})' if kind=='limit' else vr('Condicoes Operador',rr,c),None,'pct')
 fw(s,r,4,f'SUM(\'{hsn}\'!T5:T184)',None);fw(s,r,5,f'SUM(\'{hsn}\'!AB5:AB184)',None);fw(s,r,6,vr('Resultados Cenarios',sr,17),None,'mult');fw(s,r,7,vr('Resultados Cenarios',sr,19),None);fw(s,r,8,vr('Resultados Cenarios',sr,20),None,'pct');fw(s,r,9,vr('Resultados Cenarios',sr,26),None,'text')
s.write_url('A28',"internal:'Condicoes Operador'!A1",F['link'],'✎ Abrir agenda HMA e limites');s.write_url('A30',"internal:'Operador'!A1",F['link'],'↗ Matriz 4 × 3 de fees');text(s,32,0,'* Proposta não recebida');text(s,32,1,'Valores ilustrativos até documento.');alert(s,'B13')
chart(s,'L5','column','Impacto da negociação no caixa',['DASH Operador',22,0,25,0],[('CFADS 15A',['DASH Operador',22,5,25,5])],550,280)

# Hotel RevPAR and margins, added to the existing specialized panel.
s=S['DASH Hotel MICE'];metric(s,23,'RevPAR inicial',"'Cenarios'!B19*'Cenarios'!C19");metric(s,24,'Margem GOP ref.','IF(B8>0,B9/B8,"")',fmt='pct');metric(s,25,'Margem EBITDA ref.','IF(B8>0,B10/B8,"")',fmt='pct')
# New audit controls are additive and keep original control numbers intact.
s=S['AUDITORIA'];start=5+len(AUD)
extra=[('JSCV mantém 51%',"'Societario'!C5-51%",'ABS(B{r})<0.000001','Sem diluição proporcional automática.',False),('Ponte societária soma 100%',"SUM('Societario'!C5:C7)-1",'ABS(B{r})<0.000001','JSCV 51%, JALS 9%, investidores40% na referência.',False),('Participação dentro 40%',"'Retornos'!C13",'AND(B{r}>=0,B{r}<=40%)','Preço de ronda 15/20 M€ editável; não exceder ações disponíveis.',False),('JALS sem participação negativa',"'Societario'!C6",'B{r}>=0','Cessão não pode superar49%.',False),('Mecanismo entrada na KK',pv('legal_confirmed'),'B{r}=1','Compra de ações não é financiamento automático; documentar transferência irrevogável.',True),('Proposta real operador',pv('offer_received'),'B{r}=1','Simulação não é proposta recebida.',True),('Combinação cabe na ronda',"'Ronda Investidores'!E10",'B{r}<=20000000','Blocos de benefícios não são subscrições.',False)]
for r,(label,formula,condition,note,external) in enumerate(extra,start):
 text(s,r,0,label);fw(s,r,1,formula,None,'num');fw(s,r,2,f'IF({condition.format(r=r)},"Conforme","'+('Pendente de Validação Externa' if external else 'Alerta')+'")',None,'text');text(s,r,3,note)
alert(s,f'C5:C{start+len(extra)-1}');s.autofilter(3,0,start+len(extra)-2,3)

s=dash('DASH Riscos','H | Riscos e auditoria — distinguir desvios financeiros, ausência de dados e condições externas por formalizar.')
for r,label,f,fmt in [(5,'Alertas técnicos/financeiros',f'COUNTIF(\'AUDITORIA\'!C5:C{start+len(extra)-1},"Alerta")','int'),(6,'Pendências externas',f'COUNTIF(\'AUDITORIA\'!C5:C{start+len(extra)-1},"Pendente de Validação Externa")','int'),(7,'Défice não financiado',kref('fundgap'),'eur'),(8,'Falta de reservas',kref('reservegap'),'eur'),(9,'Transferência jurídica fundos',"'Ronda Investidores'!B25",'text'),(10,'Yield não coberto',"SUM('Residencial'!N5:N184)",'eur'),(11,'Sobrecusto simulado',pv('cost_overrun'),'pct'),(12,'Atraso comercial simulado',pv('commercial_delay'),'int'),(13,'HMA',"'Condicoes Operador'!B11",'text')]:metric(s,r,label,f,fmt=fmt)
s.write_row(17,0,['Controlo','Valor','Estado','Ação / limite'],F['head'])
for i,r in enumerate(range(5,start+len(extra)),19):
 for c in range(4):fw(s,i,c,vr('AUDITORIA',r,c),None,'text' if c in [0,2,3] else 'num')
alert(s,f'C19:C{18+start+len(extra)-5}')
s.set_column('D:D',95)

# Shareholder income split follows the proposed rights; individual historical costs are unavailable.
s=sh('Retorno Socios',['Acionista','Participação pós','Dividendos 15A','Saída estimada 15A','Preço venda retido','Total recebido','Base histórica individual','TIR individual'],'JSCV51/JALS9/Novos40 na referência. Contribuições históricas individuais não fornecidas: não inferir TIR por percentagens societárias.')
for r,label,share in [(5,'JSCV',"'Societario'!C5"),(6,'JALS',"'Societario'!C6"),(7,'Novos investidores',"'Retornos'!C13")]:
 text(s,r,0,label);fw(s,r,1,share,None,'pct');fw(s,r,2,f'\'Retornos\'!C12*B{r}',None);fw(s,r,3,f'\'Retornos\'!C11*B{r}',None);fw(s,r,4,"'Ronda Investidores'!B23" if r==6 else '0',None);fw(s,r,5,f'SUM(C{r}:E{r})',None);text(s,r,6,'Indisponível' if r<7 else 'Preço de aquisição');fw(s,r,7,"'Retornos'!C19" if r==7 else '"Indisponível"',None,'pct')

# Board landing page: 12 live KPI cards, eight actual linked icons, no macros.
execs=wb.add_worksheet('EXEC KK01');S['EXEC KK01']=execs;execs.hide_gridlines(2);execs.set_zoom(85);execs.set_tab_color(TEAL);execs.set_column('A:P',11);execs.set_default_row(21);execs.set_landscape();execs.fit_to_pages(1,1);execs.print_area('A1:P41');execs.set_margins(.25,.25,.25,.25)
execs.merge_range('A1:P2','KK-01  |  EXECUTIVE CONTROL CENTER',F['title']);execs.merge_range('A3:P4','V02 • 20 M€ / 40% • 180 meses • Hipótese de fundos disponíveis na KK, sujeita a formalização jurídica • Sem dados reais de execução fornecidos',F['subtitle'])
small=wb.add_format({'font_color':'#526575','font_size':10,'bold':True,'bg_color':GRAY,'text_wrap':True,'valign':'vcenter'})
cardfmts={k:wb.add_format({'font_size':21,'bold':True,'font_color':NAVY,'bg_color':GRAY,'num_format':F[k].num_format,'valign':'vcenter'}) for k in ['eur','pct','mult','num']}
CARDS=[('Capital disponível KK',"'Ronda Investidores'!B22",'eur'),('Participação investidores',"'Retornos'!C13",'pct'),('Pico da dívida',kref('peak'),'eur'),('Utilização acumulada',kref('used'),'eur'),('Caixa mínima',kref('cashmin'),'eur'),('Dívida na abertura',kref('debt30'),'eur'),('Juros 180M',kref('interest'),'eur'),('DSCR ano parcial',kref('dscrstub'),'mult'),('EBITDA ano completo',"'DASH Hotel MICE'!B10",'eur'),('TIR novos 15A',"'Retornos'!C19",'pct'),('MOIC novos 15A',"'Retornos'!C18",'mult'),('VPL novos 15A',"'Retornos'!C26",'eur')]
for i,(label,formula,fmt) in enumerate(CARDS):
 row=6+(i//4)*5;column=(i%4)*4;execs.merge_range(row-1,column,row-1,column+3,label,small);execs.merge_range(row,column,row+2,column+3,'',cardfmts[fmt]);execs.write_formula(row,column,'='+formula,cardfmts[fmt],'');execs.write_comment(row,column,'KPI ligado à fonte: '+formula+'. Resultado de cenário, não certificação. Dados de execução não inventados.')
execs.merge_range('A22:P23','NAVEGAÇÃO POR NATUREZA  •  Clique no ícone ou no título',F['head'])
MENU=[('A','INVESTIMENTO / CAPEX','DASH Investimento'),('B','REAL ESTATE / COMERCIAL','DASH Real Estate'),('C','FINANCIAMENTO / LIQUIDEZ','DASH Consolidado'),('D','HOTEL + MICE','DASH Hotel MICE'),('E','OPERADOR INTERNACIONAL','DASH Operador'),('F','EXPLORAÇÃO RESIDENCIAL','DASH Residencial'),('G','INVESTIDORES / ACIONISTAS','DASH Acionistas'),('H','RISCOS / AUDITORIA','DASH Riscos')]
ICONDIR=ROOT/'src/icons';ICONDIR.mkdir(exist_ok=True)
fontpath='/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf'
try:font=ImageFont.truetype(fontpath,36)
except OSError:font=ImageFont.load_default()
for i,(letter,label,dest) in enumerate(MENU):
 row=24+(i//4)*6;column=(i%4)*4
 im=Image.new('RGBA',(72,72),(0,0,0,0));dr=ImageDraw.Draw(im);dr.rounded_rectangle((1,1,70,70),radius=16,fill=TEAL);# Semantic vector pictograms; each image carries a real internal Excel hyperlink.
 if letter in ['A','C']:
  dr.line([(15,29),(36,15),(57,29)],fill='white',width=4);dr.line((14,57,58,57),fill='white',width=4)
  for x in [20,34,48]:dr.line((x,33,x,53),fill='white',width=4)
 elif letter=='B':
  dr.line([(13,34),(36,15),(59,34)],fill='white',width=4);dr.rectangle((21,33,52,57),outline='white',width=3);dr.rectangle((33,43,41,57),fill='white')
 elif letter=='D':
  dr.line((15,22,15,57),fill='white',width=4);dr.line((15,45,58,45),fill='white',width=4);dr.line((58,32,58,57),fill='white',width=4);dr.rectangle((22,32,54,43),outline='white',width=3);dr.ellipse((20,24,30,34),fill='white')
 elif letter=='E':
  dr.line([(15,36),(27,27),(37,33),(47,27),(59,36)],fill='white',width=4);dr.line([(17,39),(31,51),(40,42)],fill='white',width=4);dr.line([(55,39),(43,51),(29,37)],fill='white',width=4)
 elif letter=='F':
  dr.rectangle((21,14,52,59),outline='white',width=3)
  for x in [28,42]:
   for y in [23,34,45]:dr.rectangle((x,y,x+5,y+5),fill='white')
 elif letter=='G':
  for x,y in [(19,25),(43,25),(31,17)]:dr.ellipse((x,y,x+12,y+12),outline='white',width=3)
  dr.arc((14,39,35,64),180,360,fill='white',width=4);dr.arc((37,39,60,64),180,360,fill='white',width=4)
 else:
  dr.line([(36,14),(56,22),(53,45),(36,59),(19,45),(16,22),(36,14)],fill='white',width=4);dr.line([(26,35),(33,43),(47,28)],fill='white',width=4)
 path=ICONDIR/(letter+'.png');im.save(path)
 execs.insert_image(row,column,str(path),{'x_scale':.75,'y_scale':.75,'url':f"internal:'{dest}'!A1",'description':'Abrir '+label,'object_position':1})
 execs.merge_range(row,column+1,row+2,column+3,'',small);execs.write_url(row,column+1,f"internal:'{dest}'!A1",F['link'],label)
execs.merge_range('A37:P38','DECISÃO: maior equity reduz dívida e juros; o preço mais alto pelos mesmos 40% pode reduzir a TIR. Comparar 15 M€ / 20 M€ antes de decidir.',F['warn']);execs.write_url('A40',"internal:'Comparativo 15 20'!A1",F['link'],'↔ Comparativo 15 M€ / 20 M€');execs.write_url('G40',"internal:'Pressupostos'!A1",F['link'],'✎ Selecionar cenários e pressupostos');execs.write_url('L40',"internal:'Ronda Investidores'!A1",F['link'],'↗ Estrutura jurídica')
for ws in S.values():
 if ws.name!='EXEC KK01' and not ws.name.startswith('Calc') and ws.name not in ['Listas','Vet Cenarios']:ws.write_url('L1',"internal:'EXEC KK01'!A1",F['link'],'⌂ Painel Executivo')
S['DASH Real Estate'].write_url('A20',"internal:'DASH Comercial'!A1",F['link'],'↗ Abrir filtros comerciais por ID / tipo / piso / mês');S['DASH Acionistas'].write_url('A22',"internal:'Ronda Investidores'!A1",F['link'],'↗ Composição por categorias / benefícios');S['DASH Consolidado'].write_url('A26',"internal:'Comparativo 15 20'!A1",F['link'],'↔ Comparativo 15 M€ / 20 M€')
# Append structural and scenario metadata for reproducible regression.
(ROOT/'src/v02_metadata.json').write_text(json.dumps({'pair':PAIR,'scenario_map':SCMAP,'menu':MENU,'cards':CARDS,'extra_audit_start':start},ensure_ascii=False,indent=2))
# Keep the preserved 4x3 fee matrix consistent with the optional performance gate.
for r in range(5,17):
 fw(S['Operador'],r,4,f'C{r}*A{r}+MAX(0,D{r}-C{r}*A{r}-C{r}*{pv("systems")})*B{r}*IF(OR({pv("incentive_gate")}=0,AND(D{r}>={pv("gop_floor")},D{r}-C{r}*A{r}-C{r}*{pv("systems")}>={pv("ebitda_floor")})),1,0)',None)

# Selected scenario caption follows the active model rather than a fixed 20M label.
caption_formula="=\"V02 • \"&TEXT('Ronda Investidores'!B20/1000000,\"0.0\")&\" M€ / \"&TEXT('Retornos'!C13,\"0%\")&\" • Hotel \"&"+pv('hotel')+"&\" • Banco \"&"+pv('bank')+"&\" • Fundos sujeitos a formalização jurídica\""
execs.write_formula('A3',caption_formula,F['subtitle'],'V02 • 20,0 M€ / 40% • Hotel Favorável • Banco Limite aceitável • Fundos sujeitos a formalização jurídica')

metric(S['DASH Acionistas'],20,'Payback por dividendos (meses)',kref('investor_payback'),fmt='int')
S['DASH Acionistas'].write_url('A24',"internal:'Societario'!A1",F['link'],'↗ Ponte acionista JSCV / JALS / novos')
S['DASH Acionistas'].write_url('A26',"internal:'Retorno Socios'!A1",F['link'],'↗ Retorno por acionista e limites históricos')
