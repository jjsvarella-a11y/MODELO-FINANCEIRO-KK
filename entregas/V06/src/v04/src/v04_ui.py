from v04_formula_helpers import note_method
from returns10 import payback_formula
# Executive navigation and one source of contextual KPI definitions.
DEFS={
'vgv':('Valor Global de Vendas','Soma dos preços das94frações, antes de custos comerciais.','Potencial comercial; não é caixa recebida nem contratos assinados.','€','Comparar com recebimentos e custos comerciais.'),
'area':('Área privativa comercializada','Soma das áreas individuais do inventário.','Permite reconciliar quantidades e preço médio; não é área construída total.','m²','Conferir projeto e áreas comuns.'),
'pm2':('Preço médio por metro quadrado','VGV dividido pela área privativa total.','Média ponderada; não média simples dos preços unitários.','€/m²','Comparar por piso,tipologia e unidade.'),
'commercial':('Custos comerciais','Transmissão,marketing e comissãoKyenu separadamente calendarizados.','Diminuem o rendimento líquido das vendas; incidência fiscal por validar.','€','Hipóteses1,5%/1%/1,5%.'),
'netre':('Rendimento líquido das vendas','VGV menos custos comerciais previstos.','Antes de construção,financiamento e eventual imposto da promoção.','€','Não confundir com margem líquida final.'),
'capex':('Investimento futuro na obra','Construção,FF&E inicial,fiscalização e contingência paga.','Histórico4,5M€não é nova saída futura; atraso/ETCrevisto alteram forecast.','€','Orçamento aprovado por confirmar.'),
'rescapex':('CAPEX imputado residencial','Quota do investimento atribuída às frações por área.','É proxy de segmentação; carece de orçamento por componente.','€','Não é custo específico contratualmente aprovado.'),
'remargin':('Resultado da promoção — proxy','Vendas líquidas menosCAPEXimputado ao residencial.','Margem antes de custos financeiros e fiscalidade completa.','€','Confirmar imputação e impostos.'),
'remarginpct':('Margem imobiliária — proxy','Resultado da promoção dividido peloVGV.','Permite comparar preços e custo imputado; não representa lucro líquido auditado.','%','Sem limiar universal.'),
'peak':('Pico intramensal da dívida bancária','Maior saldo inicial acrescido do desembolso do mês.','Mede a exposição máxima; distinto da utilização acumulada da linha não renovável.','€','Comparar com limite e capacidade mobilizável.'),
'debt30':('Dívida bancária noM30','Saldo bancário após pagamentos do mês30.','Não inclui automaticamente empréstimos dos vendedores, mostrados separadamente.','€','M30é referência; abertura pode atrasar.'),
'cash30':('Caixa noM30','Saldo de tesouraria após entradas,saídas e distribuições doM30.','Pode conter reservas protegidas; não é todo livre para dividendos.','€','Comparar com reserva protegida.'),
'debt42':('Dívida bancária noM42','Saldo do banco no final do mês42.','Permite testar transição da carência de capital.','€','Carência42Mé hipótese de negociação.'),
'cash42':('Caixa noM42','Tesouraria final do mês42.','Verificar reserva e serviço obrigatório na transição da carência.','€','Não confundir com caixa livre.'),
'interest':('Juros bancários acumulados','Soma dos juros sobre saldos bancários efetivamente usados no cenário.','Menor utilização/taxa pode reduzir custo; linha disponível é proteção,não recomendação de redução.','€','Outros encargos e crédito de vendedores separados.'),
'fees':('Comissões e success fee','Abertura,mobilização e broker condicionados à primeira utilização.','Taxas de referência não são contrato; verificar sobreposição.','€','Custo de compromisso adicional editável.'),
'cashmin':('Caixa mínima projetada','Menor saldo mensal do fluxo integrado.','Saldo positivo não prova cumprimento da reserva nem funding assegurado.','€','Ver insuficiência de reservas e completionsecured.'),
'reservegap':('Insuficiência máxima de reservas','Maior diferença positiva entre reserva exigida e caixa.','Mesmo sem caixa negativa,insuficiência viola a política de segurança financeira.','€','0é referência interna de cumprimento.'),
'fundgap':('Défice não financiado','Maior falta de caixa após financiamento contratualmente simulável.','TIR positiva não torna um cenário com défice viável.','€','0necessário,mas não suficiente.'),
'debtend':('Dívida bancária final','Saldo do crédito bancário noM180.','Empréstimos dos vendedores constam em coluna própria e são deduzidos na saída.','€','Prazo contratual pode exceder horizonte.'),
'cashend':('Caixa final','Tesouraria noM180antes da venda hipotética dos ativos.','Somada uma única vez na saída,com reservas restritas deduzidas.','€','Não duplicar no valor terminal.'),
'npv':('VPL do projeto sem saída','Valor presente dos fluxos futuros do projeto sem alienação terminal.','Positivo indica excedente acima da taxa de desconto adotada.','€','Taxa de projeto editável; não valor de mercado.'),
'npvexit':('VPL do projeto com saída','VPL dos fluxos futuros acrescido do valor terminal económico.','Depende de avaliação e liquidez não contratadas.','€','Taxa de projeto editável; saída estimada.'),
'projectirr':('TIR do projeto sem saída','Taxa anual efetiva dos fluxos incrementais do empreendimento.','Fluxos não convencionais podem admitir mais de uma raiz; avaliar tambémVPL.','%','Não substitui análise de liquidez.'),
'projectirrexit':('TIR do projeto com saída','Retorno anualizado incluindo alienação estimada doHotel.','Depende do múltiplo,EBITDA e custos da saída.','%','Não retorno garantido.'),
'equityirr':('TIR efetiva dos novos investidores','Retorno dos pagamentos previstos dos investidores e recebimentos pela waterfall selecionada.','Diferente daTIRconvencional dos blocos promotores; datas de pagamento importam.','%','Comparar com retorno exigido e riscos.'),
'equitymoic':('MOIC dos novos investidores','Total recebido dividido pelo capital efetivamente pago no cenário.','Não mede o tempo até aos recebimentos; usar comTIR.','x','Não inclui capital meramente prometido.'),
'equitynpv':('VPL dos novos investidores','Valor presente de recebimentos menos pagamentos pela política selecionada.','Negativo significa retorno inferior à taxa de desconto escolhida.','€','12%é hipótese editável,não limite universal.'),
'dscrstub':('DSCR do ano inicial parcial','CFADSHotel dividido pelo serviço obrigatório no ano de abertura parcial.','Pode incluir juros da construção; não comparar indiscriminadamente com ano estabilizado.','x','Alerta1,3xé interno,não covenant aprovado.'),
'dscrfull':('DSCR dos anos operacionais completos','Menor cobertura anual com serviço obrigatório e operação completa.','Exclui amortização voluntária; sem serviço de dívida o rácio fica indefinido.','x','Não usar999para ausência de dívida.'),
'headroom':('Margem até ao limite de crédito','Diferença entre limite estratégico e exposição bancária máxima.','Não significa capacidade disponível para novo desembolso numa linha não renovável.','€','Conferir utilizado acumulado eCP.'),
'safety':('Margem de segurança de caixa','Folga de tesouraria face às reservas do cenário.','Folga projetada não é liquidez documentada nem garantia de conclusão.','€','Política de reservas editável.'),
'payback':('Payback do projeto','Mês de recuperação acumulada do investimento futuro sem saída.','Não inclui avaliações hipotéticas; pode não ocorrer no horizonte.','meses','Usar comVPL eTIR.'),
'used':('Utilização acumulada da linha','Soma de todos os desembolsos bancários.','Amortizações não repõem capacidade de uma linha não renovável.','€','Limite estratégico30M€na referência.'),
'unused':('Linha não utilizada','Limite estratégico menos desembolsos acumulados.','Não equivale a crédito efetivamente mobilizável: depende deCP,calendário e contrato.','€','Distinguir contratado/utilizado/mobilizável.'),
'investor_payback':('Payback por dividendos dos investidores','Mês em que dividendos acumulados cobrem o capital pago.','Exclui a saída estimada; retorno pode depender dessa saída.','meses','Não promessa de recuperação.'),
'cfc_projected':('Completion Funding Coverage projetada','Fontes projetadas noM24divididas pelas necessidades para concluir e abrir.','Maior que1não prova financiamento firme; exige análise mensal e documentação.','x','M24é marco estratégico antes da conclusãoM30.'),
'cfc_secured':('Completion Funding Coverage assegurada','Recursos firmes documentados divididos por custo de conclusão validado.','Indisponível enquanto faltarem extratos,CP,compromissos e validação de custos/calendário.','x','Não classificar como coberto sem evidência.'),
'eac':('Estimate at Completion','Actual Cost documentado mais previsão do custo remanescente.','Não apresentar plano como executado;semActual Cost documentado fica indisponível.','€','Comparar comBAC aprovado.'),
'etc':('Cost to Complete','Custo remanescente previsto após o mês de controlo.','RevisãoETCalteraCAPEXfuturo e financiamento; não reescreve dados históricos desconhecidos.','€','Validar medições e compromissos.'),
'bac':('Budget at Completion','Orçamento total de referência para a obra.','Distinguir orçamento inicial,aprovado e forecast.','€','Aprovação documental separada.'),
'vac':('Variance at Completion','BAC menosEAC.','Negativo indica sobrecusto previsto; indisponível seActual Cost não comprovado.','€','SemActual Costnão afirmar desvio executado.'),
'conventional':('TIR convencional por bloco','Retorno sobre a base10/5/1M€referida aoM1,com direitos iguais dentro da categoria.','Para promotores é referência negocial,não investimento novo;não confundir comTIRhistórica.','%','Metas10%/6%/5,5%não garantidas nem aprovadas.'),
'opening':('Prontidão documental da abertura','Proporção dos13requisitos com progresso,aprovação e documento.','Não equivale a progresso físico;sem dados a prontidão é indisponível.','%','Abertura31Mna referência;licenças e aceitação obrigatórias.'),
'adr':('ADR — Average Daily Rate','Receita média de alojamento por quarto ocupado.','DistinguirHotel200quartos do pool94apartamentos.','€/noite','Hipóteses por cenário e crescimento anual.'),
'occupancy':('Taxa de ocupação','Noites ocupadas divididas pela capacidade disponível.','Quartos disponíveis dependem de abertura e,na residência,de escritura no cenário.','%','Não é forecast contratualmente garantido.'),
'gop':('GOP — resultado operacional bruto','Receitas menos custos departamentais e gastos gerais.','Base para avaliar gestão antes dos fees e custos subsequentes.','€','Comparar cenário sem operador e remuneração.'),
'ebitda':('EBITDAHotel+MICE','GOPmenos fees e custos de benefícios atribuídos.','Não é caixa distribuível;faltamFF&E,CAPEX,impostos,fundo de maneio e dívida.','€','Hotel200quartos,sem duplicar94frações.'),
'cfads':('CFADS — caixa disponível para dívida','Caixa operacional apósFF&E,reposição,imposto proxy e variação do fundo de maneio.','Base de cobertura do serviço obrigatório; não confundir comEBITDA.','€','Modelo fiscal simplificado por validar.'),
'operator':('Encargos do Operador Internacional','Fees base,incentivo e distribuição/sistemas em relação aoCFADS.','Avaliar custo e valor produzido;ratios altos sinalizam pressão na caixaKK.','% / €','Alertas internos35%gestão,25%sistemas,50%total; não limites contratuais.'),
}
# Selected waterfall totals replace only dashboard references; historical pro-rata tables remain explicit benchmarks.
ret=S['V04 Retornos']
newexit="'V04 Waterfall180'!J184/2*'V04 Waterfall180'!Q184+'V04 Waterfall180'!K184/3*'V04 Waterfall180'!R184+'V04 Waterfall180'!L184/15*'V04 Waterfall180'!S184"
legacy=f'AND({pv("capital_scenario")}<>"KK-E25",{pv("equity_mode")}<>"Combinação tickets")'
for r,label,f,fmt in [(34,'Promotores total recebido',"SUM('V04 Waterfall180'!Y5:Y184)",'eur'),(35,'PromotoresMOICproxy',"IF("+pv('historic')+">0,B34/"+pv('historic')+",\"\")",'mult'),(36,'Novos valor saída',f'IF({legacy},\'V04 Politica\'!B11*\'V04 Capital180\'!G184,{newexit})','eur'),(37,'Novos dividendos',"SUM('V04 Waterfall180'!T5:T184)-B36",'eur'),(38,'Novos total recebido',"SUM('V04 Waterfall180'!T5:T184)",'eur'),(39,'Novos lucro',"B38-SUM('V04 Waterfall180'!U5:U184)",'eur'),(40,'Retido pelos vendedores',"SUM('V04 Capital180'!E5:E184)",'eur'),(41,'Crédito vendedores recebido',"SUM('Fluxo 180M'!AV5:AW184)+'V04 Politica'!B17",'eur')]:text(ret,r,0,label);fw(ret,r,1,f,None,fmt)
remap={'C19':"'V04 Retornos'!B27",'C18':"'V04 Retornos'!B28",'C26':"'V04 Retornos'!B29",'C14':"'V04 Retornos'!B36",'C15':"'V04 Retornos'!B37",'C16':"'V04 Retornos'!B38",'C17':"'V04 Retornos'!B39",'C21':"'V04 Retornos'!B34",'C22':"'V04 Retornos'!B35",'C23':"'V04 Retornos'!B32",'B19':"'V04 Retornos'!B55",'B18':"'V04 Retornos'!B56"}
for sn,ws in S.items():
 if sn.startswith('DASH') or sn=='EXEC KK01':
  for rr,cells in list(ws.table.items()):
   for cc,cell in list(cells.items()):
    if hasattr(cell,'formula'):
     f=cell.formula
     for addr,target in remap.items():f=f.replace("'Retornos'!"+addr,target)
     if f!=cell.formula:ws.write_formula(rr,cc,'='+f.lstrip('='),cell.format,'')
S['Ronda Investidores'].write('B34','V04: JSCV1Platinum+2Gold',F['text'])
# Current performance fields use the selected policy and actual planned payment dates.
for key,addr,fmt in [('equityirr','B27','pct'),('equitymoic','B28','mult'),('equitynpv','B29','eur')]:fw(S['KPI'],KROW[key],1,"'V04 Retornos'!"+addr,None,fmt)
fw(S['KPI'],KROW['investor_payback'],1,payback_formula(),None,'text')
fw(S['Ronda Investidores'],22,1,"SUM('V04 Capital180'!F5:F184)",None);fw(S['Ronda Investidores'],23,1,"SUM('V04 Capital180'!E5:E184)",None)
S['Retornos'].write('A2','Comparação proporcional histórica. Para políticaV04selecionada e pagamentos efetivos de cenário usarV04Retornos/Waterfall180.',F['subtitle'])
S['Retorno Categorias V03'].write('A2','Referência proporcionalV03preservada. V04usaV04Retornos:20blocos e política diferenciada, sem direitos aprovados.',F['subtitle'])
for r,label,f in [(5,'Platinum','2'),(6,'Gold','3'),(7,'Silver','15')]:
 text(S['Beneficios V03'],r,0,label);fw(S['Beneficios V03'],r,2,f,None,'int')
S['Opening Readiness'] if False else None
fw(S['V04 Opening Readiness'],20,1,'IF(COUNT(E5:E17)>0,COUNTIF(I5:I17,"Pronto documentado")/13,"Indisponível")','Indisponível','pct')
# Replace obsolete five tiny metrics and their comments; add usable native navigation tiles.
e=S['EXEC KK01'];e.charts=[c for c in e.charts if c[0]<40]
for rr in range(42,79):
 for cc in range(16):
  if rr in e.table and cc in e.table[rr]:del e.table[rr][cc]
  if rr in e.comments and cc in e.comments[rr]:del e.comments[rr][cc]
e.write('A1','KK-01 | EXECUTIVE CONTROL CENTER V04',F['title']);fw(e,3,0,'"V04 • CapitalKKprevisto "&TEXT(\'Ronda Investidores\'!B22/1000000,"0.0")&" M€ • JSCV40% / JALS10% / novosaté50% • Waterfall "&'+pv('waterfall_mode'),None,'text')
for addr,label,formula,fmt in [('A6','CapitalKKprevisto',"'Ronda Investidores'!B22",'eur'),('E6','Quota paga equivalente — cenário',"'V04 Capital180'!G184",'pct'),('E16','TIR novos efetiva15A',"'V04 Retornos'!B27",'pct'),('I16','MOIC novos efetivo',"'V04 Retornos'!B28",'mult'),('M16','VPL novos15A',"'V04 Retornos'!B29",'eur')]:
 row,c=xl_cell_to_rowcol(addr);e.write(row,c,label,small);fw(e,row+2,c,formula,None,fmt)
e.write('A37','PROPOSTA DE WATERFALL: direitos preferenciais não aprovados. Liquidez projetada não é financiamento assegurado. Consultar Completion e decisões doBoard.',F['warn'])
shortcuts=[('Master Workflow','V04 Master Workflow'),('Equity & Pipeline','V04 Pipeline'),('Funding CompletionM24','V04 Completion'),('Construção &EAC','V04 EAC'),('Decisões doBoard','V04 Decisoes Board'),('Auditoria','AUDITORIA'),('DicionárioKPI','V04 Dicionario KPI'),('Investidores &Waterfall','V04 Retornos')]
for i,(label,dest) in enumerate(shortcuts):
 row=42+(i//4)*3;cc=(i%4)*4;e.merge_range(row-1,cc,row,cc+3,'',F['head']);e.write_url(row-1,cc,"internal:'"+dest+"'!A1",F['link'],label)
for i,(label,f,fmt) in enumerate([('PlatinumTIRconvencional',"'V04 Retornos'!D5",'pct'),('GoldTIRconvencional',"'V04 Retornos'!D6",'pct'),('SilverTIRconvencional',"'V04 Retornos'!D7",'pct'),('Completionsecured',"'V04 Completion'!C17",'mult')]):
 cc=i*4;e.merge_range(49,cc,49,cc+3,label,small);e.merge_range(50,cc,52,cc+3,'',cardfmts[fmt]);e.write_formula(50,cc,'='+f,cardfmts[fmt],'')
for cc,label,f,fmt in [(0,'EAC comActual Costdocumentado',"'V04 EAC'!B193",'eur'),(8,'Prontidão documental abertura',"'V04 Opening Readiness'!B20",'pct')]:e.merge_range(54,cc,54,cc+7,label,small);e.merge_range(55,cc,57,cc+7,'',cardfmts[fmt]);e.write_formula(55,cc,'='+f,cardfmts[fmt],'')
chart(e,'A61','column','TIR convencional por categoria: meta e cenário',['V04 Retornos',4,0,6,0],[('Meta',['V04 Retornos',4,2,6,2]),('Selecionado',['V04 Retornos',4,3,6,3])],width=1000,height=300);e.print_area('A1:P79');e.fit_to_pages(1,1)

# Dictionary rows are the only definition source for generated classic notes.
DICT_ENTRIES=[];strlookup={i:s for s,i in wb.str_table.string_table.items()};locations={}
for key,r in KROW.items():locations[('KPI',r,1)]=key
for item in [('V04 Completion',17,1,'cfc_projected'),('V04 Completion',17,2,'cfc_secured'),('V04 EAC',188,1,'bac'),('V04 EAC',192,1,'etc'),('V04 EAC',193,1,'eac'),('V04 EAC',194,1,'vac'),('V04 Opening Readiness',20,1,'opening')]:locations[item[:3]]=item[3]
for r in [5,6,7]:locations[('V04 Retornos',r,3)]='conventional'
for sn,ws in S.items():
 if sn.startswith('DASH'):
  for row,cells in list(ws.table.items()):
   for cc,cell in list(cells.items()):
    if cc!=1 or not hasattr(cell,'formula'):continue
    import re
    k=None;match=re.search(r"'KPI'!\$?B\$?(\d+)",cell.formula)
    if match:k=next((key for key,val in KROW.items() if val==int(match.group(1))),None)
    if not k:
     title=cells.get(0);label=strlookup.get(getattr(title,'string',-1),'').lower()
     for token,key in [('adr','adr'),('ocupação','occupancy'),('gop','gop'),('ebitda','ebitda'),('cfads','cfads'),('operador','operator'),('fee','operator'),('gestão','operator'),('sistemas','operator'),('caixa','cashmin'),('dívida','peak'),('reserva','reservegap'),('vpl','equitynpv'),('tir','equityirr'),('moic','equitymoic')]:
      if token in label:k=key;break
    if k:locations[(sn,row+1,cc)]=k
for r,c,k in [(7,0,'capex'),(7,4,'equityirr'),(7,8,'peak'),(7,12,'used'),(12,0,'cashmin'),(12,4,'debt30'),(12,8,'interest'),(12,12,'dscrstub'),(17,0,'ebitda'),(17,4,'equityirr'),(17,8,'equitymoic'),(17,12,'equitynpv'),(51,0,'conventional'),(51,4,'conventional'),(51,8,'conventional'),(51,12,'cfc_secured'),(56,0,'eac'),(56,8,'opening')]:locations[('EXEC KK01',r,c)]=k
# Correct capital and paid-equivalent definitions for first two executive cards.
DEFS['capital_kk']=('Capital previsto paraKK','Soma dos pagamentos de cenário efetivamente transferidos à sociedade.','Não confundir preço recebido pelos vendedores com entradaKK ou capital documentado.','€','Transferência e instrumento jurídico por validar.');locations[('EXEC KK01',7,0)]='capital_kk'
DEFS['quota_paga']=('Participação paga equivalente','Quota económica contratada multiplicada pela fração do preço paga no cenário.','Diagnóstico económico;não reconhece titularidade jurídica ou direitos antes da formalização.','%','Quadros distinguem estrutura alvo e pagamento.');locations[('EXEC KK01',7,4)]='quota_paga'
s=sh('V04 Dicionario KPI',['Código','Nome completo','Definição','Interpretação','Unidade','Referência / limite','Motor / fonte','Célula / fórmula','Método'],'Fonte única das definições. Notas clássicas geradas destes registos; atualizar pelo script ao alterar textos. ExcelDesktopnão foi testado; reaberturaOOXML/LibreOffice e dimensões serão verificadas.')
s.set_column('B:D',58);s.set_column('F:I',70)
for i,(key,record) in enumerate(DEFS.items(),5):
 title,definition,interpretation,unit,reference=record;linked=[(sn,r,c) for (sn,r,c),k in locations.items() if k==key];source='; '.join(sn+'!'+col(c)+str(r) for sn,r,c in linked[:8]);formulas=[]
 for sn,r,c in linked[:2]:
  cell=S[sn].table.get(r-1,{}).get(c);formulas.append(getattr(cell,'formula',''))
 method='Fórmulas mensais180M; transformação e agregado conforme fonte; sem substituição por constantes.'
 for c,val in enumerate([key,title,definition,interpretation,unit,reference,source,' | '.join(formulas),method]):text(s,i,c,val)
 DICT_ENTRIES.append({'key':key,'title':title,'definition':definition,'interpretation':interpretation,'unit':unit,'reference':reference,'source':source,'formula':' | '.join(formulas),'method':method})
s.autofilter(3,0,len(DEFS)+3,8)
for (sn,r,c),key in locations.items():
 title,definition,interpretation,unit,reference=DEFS[key];cell=S[sn].table.get(r-1,{}).get(c);formula=getattr(cell,'formula','');note=f'{title}\n\nDefinição: {definition}\n\nInterpretação: {interpretation}\n\nUnidade: {unit}. {reference}\n\nAUDITORIA\nMotor/fonte: {sn}!{col(c)}{r}\nMétodo: {note_method(key)}\nFórmula integral no dicionário,colunaH.\nDicionário: V04 Dicionario KPI / {key}'
 S[sn].write_comment(r-1,c,note,{'author':'KK-01 | KPI','width':500,'height':360,'font_size':10,'visible':False})
# Native navigation and active reference labels.
for name,ws in S.items():
 if name.startswith('V04 ') and not ws.hidden:ws.write_url('L1',"internal:'EXEC KK01'!A1",F['link'],'← Executivo');ws.set_column('A:A',40)
S['LEIA-ME'].write('B8','V04:20blocos;JSCV1P+2G40%,JALS5S10%,novos1P1G10S50%.25M€planeados;linha30M€back-up. Waterfall é proposta não aprovada.',F['subtitle'])
S['LEIA-ME'].write_url('D5',"internal:'V04 Master Workflow'!A1",F['link'],'Iniciar pelo Workflow e suas dependências')
S['DASH Acionistas'].write_url('A28',"internal:'V04 Retornos'!A1",F['link'],'V04: remuneração diferenciada e pagamentos por bloco')
S['DASH Consolidado'].write_url('A28',"internal:'V04 Completion'!A1",F['link'],'V04: liquidez projetada vssecured /M24')
# Explicit conditional financial controls; external documentation absence never becomes an approved state.
a=S['AUDITORIA'];start=max(a.table)+3
for r,label,formula,external in [(start,'V04:20blocos',"SUM('Matriz Societaria V03'!B9:D9)=20",False),(start+1,'V04:ações100%',"'Matriz Societaria V03'!G9=100000",False),(start+2,'V04:conservação waterfall',"MAX('V04 Waterfall180'!P5:P184)<0.01",False),(start+3,'V04:capital sem duplicação',"ABS('V04 Pipeline'!B29)<0.01",False),(start+4,'V04:recursos firmesM24',"COUNT('V04 Completion'!C5:C8)=4",True),(start+5,'V04:preferências aprovadas','FALSE',True),(start+6,'V04:custos para conclusão validados',pv('completion_cost_validated')+'=1',True)]:text(a,r,0,label);fw(a,r,1,f'IF({formula},0,1)',None);fw(a,r,2,f'IF({formula},"Conforme","'+('Pendente de Validação Externa' if external else 'Alerta')+'")',None,'text')
(ROOT/'src/v04_kpi_dictionary.json').write_text(json.dumps(DICT_ENTRIES,ensure_ascii=False,indent=2));(ROOT/'src/v04_risk_metadata.json').write_text(json.dumps(RISK_META,ensure_ascii=False,indent=2));(ROOT/'src/v04_risk_results.json').write_text(json.dumps(RISK_RESULTS,ensure_ascii=False))

S['DASH Acionistas'].write_formula('B18',"='V04 Riscos'!N5",F['text'],'')
