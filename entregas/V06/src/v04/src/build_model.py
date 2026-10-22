from engine import *
import xlsxwriter, json,datetime
from xlsxwriter.utility import xl_col_to_name as col
OUT=ROOT/'entrega/KK01_Modelo_Integrado_V04.xlsx'
wb=xlsxwriter.Workbook(OUT,{'use_future_functions':False});wb.set_properties({'title':'KK-01 | Modelo financeiro integrado','author':'KK — reconstrução técnica','comments':'Cenários auditáveis; hipóteses não constituem contratos.'});wb.set_calc_mode('auto')
NAVY='#102C46';TEAL='#007F86';GREEN='#13835C';BLUE='#E6F2FF';GOLD='#FFF2CC';RED='#FCE4D6';GRAY='#EDF2F5'
F={}
F['title']=wb.add_format({'bold':True,'font_size':20,'font_color':'white','bg_color':NAVY,'valign':'vcenter'})
F['subtitle']=wb.add_format({'font_color':'#526575','text_wrap':True,'valign':'top'})
F['head']=wb.add_format({'bold':True,'font_color':'white','bg_color':TEAL,'text_wrap':True,'valign':'vcenter','border':1,'border_color':'white'})
for key,num in [('eur','#,##0.00 "€";[Red](#,##0.00) "€";–'),('num','#,##0.00;[Red](#,##0.00);–'),('int','0'),('pct','0.00%'),('mult','0.00"x"'),('text','@')]:
 F[key]=wb.add_format({'font_name':'Calibri','font_size':10,'num_format':num,'valign':'vcenter'})
 F['input_'+key]=wb.add_format({'font_name':'Calibri','font_size':10,'num_format':num,'bg_color':BLUE,'font_color':'#1454A3','bottom':1,'bottom_color':'#98B9DD'})
 F['total_'+key]=wb.add_format({'bold':True,'num_format':num,'bg_color':GRAY,'top':1,'top_color':TEAL})
F['good']=wb.add_format({'font_color':GREEN,'bg_color':'#E2F0D9','bold':True});F['bad']=wb.add_format({'bg_color':RED,'font_color':'#9C0006'});F['warn']=wb.add_format({'bg_color':GOLD,'font_color':'#7F6000'});F['good_pct']=wb.add_format({'font_color':GREEN,'bg_color':'#E2F0D9','bold':True,'num_format':'0.00%'});F['good_int']=wb.add_format({'font_color':GREEN,'bg_color':'#E2F0D9','bold':True,'num_format':'0'});F['link']=wb.add_format({'font_color':TEAL,'underline':True});F['kpi']=wb.add_format({'bold':True,'font_size':15,'num_format':'#,##0.00','font_color':NAVY,'bg_color':GRAY})
S={};BASE=run();PR={k:i+5 for i,(k,*_) in enumerate(PARAMS)}
def sh(name,headers=None,note=''):
 s=wb.add_worksheet(name);S[name]=s;s.hide_gridlines(2);s.set_tab_color(TEAL);s.set_zoom(85);s.set_default_row(18);s.set_column(0,0,14);s.set_column(1,30,18);s.set_row(0,34);s.merge_range(0,0,0,9,'KK-01 | '+name,F['title']);s.merge_range(1,0,2,9,note,F['subtitle']);s.set_row(1,22);s.freeze_panes(4,1);s.set_landscape();s.fit_to_pages(1,0);s.repeat_rows(0,3);s.set_footer('&LKK-01 — hipóteses por validar&C&P / &N&RModelo integrado 180 meses')
 if headers:
  s.write_row(3,0,headers,F['head']);s.set_row(3,42)
 return s

def rawpv(k):return f"'Pressupostos'!$B${PR[k]}"
def pv(k):
 if k in ['equity_price','equity_stake','legacy_jscv']:return "'Capital V03'!$B$"+str({'equity_price':6,'equity_stake':7,'legacy_jscv':8}[k])
 if k=='treasury_ratio':return f"({rawpv('treasury_ratio')}*{rawpv('received_ratio')})"
 if k=='equity':return "'Ronda Investidores'!$B$20"
 if k in ['base_fee','systems','incentive']:return "'Condicoes Operador'!$F$"+str({'base_fee':5,'systems':6,'incentive':7}[k])
 return rawpv(k)

def vr(sn,r,c,abs=False):return f"'{sn}'!{'$' if abs else ''}{col(c)}{'$' if abs else ''}{r}"
def fw(s,r,c,formula,value=None,fmt='eur'):
 if value is None:value=''
 s.write_formula(r-1,c,'='+formula,F[fmt],value)
def text(s,r,c,value,fmt='text'):s.write(r-1,c,value,F[fmt])
def headers(s,names):s.write_row(3,0,names,F['head'])
def table(s,lastrow,lastcol,name):
 s.autofilter(3,0,lastrow-1,lastcol);s.print_area(0,0,lastrow-1,min(lastcol,14))
def alert(s,cell):
 s.conditional_format(cell,{'type':'text','criteria':'containing','value':'Alerta','format':F['bad']});s.conditional_format(cell,{'type':'text','criteria':'containing','value':'Conforme','format':F['good']});s.conditional_format(cell,{'type':'text','criteria':'containing','value':'Pendente','format':F['warn']})

read=sh('LEIA-ME',['Conteúdo','Instrução'],'Modelo nominal em euros | meses relativos 1–180 | sem macros | cenários não contratados')
read.set_column('A:A',30);read.set_column('B:B',105)
notes=[('Estado','Reconstrução técnica da V45. Resultados condicionais às hipóteses e confirmações externas; consultar AUDITORIA.'),('Como começar','Alterar apenas células azuis em Pressupostos, Unidades, Obra e Cenarios. Excel com cálculo automático.'),('Sequência','Pressupostos → Unidades / Obra → Imobiliario / CAPEX → Hotel / Residencial → Fluxo 180M → Anual / KPI → dashboards.'),('Referência','Equity 20 M€ por 40%, condicionado à entrada na KK; banco Limite aceitável (6%, 42 meses, 10 anos); Hotel Favorável. Testar também Base e Conservador.'),('Calendário','Obra meses 1–30. Abertura mês 31: convenção explícita após concluir a obra, diferente da abertura mês 30 da V45.'),('Azul / verde','Azul = editável. Verde = condição negocial favorável; seleção ativa não significa condição contratada.'),('Reservas','Caixa operacional mínima + buffer obra + HOTUSA + contingência não realizada. Inclui colchão pré-abertura calculado para fundo de maneio e juros.'),('Cash sweep','Preserva reservas; obra: limitado às vendas líquidas e proteção de compromissos futuros. Banco não renovável.'),('DSCR','CFADS Hotel / (juros + capital contratual). Não inclui amortização voluntária; DSCR total separado.'),('Residencial','Pool separado, rendas não garantidas. Residual KK 45% antes de imposto, hipótese herdada e por validar; testar res_include=0.'),('TIR','Fluxos não convencionais podem não ter TIR única. Consultar sinalizações; VPL sem saída e com saída separados.'),('Valor terminal','EBITDA Hotel últimos 12 meses × múltiplo; custos de saída e dívida deduzidos; caixa somada uma vez; HOTUSA e reserva contingente deduzidas.'),('Sem dupla contagem','Apartamentos vendidos não entram no valor terminal Hotel; histórico 4,5 M€ não é saída futura; fundos restritos fora dos dividendos.'),('504 / 810','Mês 29 mantido como hipótese V45, não contratualmente validado. Modificar em Unidades após confirmação.'),('Conversão','Alternativa ilustrativa: fração escolhida recebe abatimento de cobranças; capital novo diminui pelo mesmo montante. Não aumenta retorno nem garante recompra.'),('Relatórios','Auditoria.md, Viabilidade_Executiva.md, Documentacao_Tecnica.md e testes acompanham este Excel.'),('Limitações','Modelo de decisão, não certificação fiscal, jurídica, bancária ou avaliação independente; ver pendências específicas.'),('Fonte','V45 original preservada em original/V45_original.xlsx; inventário das 70 folhas e dependências exportados.'),('Garantia de uso','Não apagar folhas de cálculo nem colar valores sobre fórmulas. Reexecutar testes após alterações estruturais.')]
for r,(a,b) in enumerate(notes,5):text(read,r,0,a);text(read,r,1,b);read.set_row(r-1,32)

s=sh('Pressupostos',['Parâmetro','Valor editável','Unidade','Origem / critério / validação'],'Única fonte de pressupostos gerais. Células azuis editáveis; cenários e direitos ainda não contratualizados.')
s.set_column('A:A',47);s.set_column('B:B',23);s.set_column('C:C',17);s.set_column('D:D',106)
for k,label,value,unit,note in PARAMS:
 r=PR[k];fmt='text' if isinstance(value,str) else ('pct' if '%' in unit or unit=='p.p.' else 'int' if unit in ['mês','meses','anos','unidades','pessoas','0/1'] else 'mult' if unit=='x' else 'eur' if '€' in unit else 'num')
 text(s,r,0,label);s.write(r-1,1,value,F['input_'+fmt]);text(s,r,2,unit);text(s,r,3,note);s.write_comment(r-1,1,note,{'author':'KK-01 — pressuposto'})
 def dv(spec):s.data_validation(r-1,1,r-1,1,dict(spec,error_type='stop',error_title='Valor não permitido',error_message='Utilize os valores indicados e consulte a nota.'))
 if k=='equity':dv({'validate':'list','source':['0','5000000','10000000','15000000','20000000','25000000']})
 elif k=='limit':dv({'validate':'list','source':['20000000','25000000','30000000']})
 elif k=='capital_schedule':dv({'validate':'list','source':['Acelerado','Base','Prudente','V03']})
 elif k=='waterfall_mode':dv({'validate':'list','source':['Proporcional','Dividendos','Saída','Combinado','Otimizado']})
 elif k=='pricing_mode':dv({'validate':'list','source':['Origem','Motor']})
 elif k=='payout_basis':dv({'validate':'list','source':['Proxy prudente','Documentado']})
 elif k=='dividend_frequency':dv({'validate':'list','source':['1','12']})
 elif k=='capital_scenario':dv({'validate':'list','source':['KK-E15','KK-E20-A','KK-E20-B','KK-E25']})
 elif k=='funding_kind':dv({'validate':'list','source':['Não reembolsável','Empréstimo']})
 elif k=='bank':dv({'validate':'list','source':['Abertura','Objetivo','Limite aceitável','Personalizado']})
 elif k=='fee_mode':dv({'validate':'list','source':['Objetivo KK','Proposta simulada','Limites propostos','Personalizado']})
 elif k=='equity_mode':dv({'validate':'list','source':['Capital selecionado','Combinação tickets']})
 elif k=='equity_price':dv({'validate':'list','source':['15000000','20000000','25000000']})
 elif k=='hotel':dv({'validate':'list','source':list(HOTELS)})
 elif k=='hotusa':dv({'validate':'list','source':['0','500000','1000000']})
 elif k=='rate_custom':dv({'validate':'list','source':['0.045','0.05','0.055','0.06','0.065','0.07','0.075','0.08']})
 elif k=='term_custom':dv({'validate':'list','source':['5','10','15','20']})
 elif k=='base_fee':dv({'validate':'list','source':['0.04','0.035','0.03','0.025']})
 elif k=='incentive':dv({'validate':'list','source':['0.10','0.08','0.06']})
 elif unit=='0/1':dv({'validate':'integer','criteria':'between','minimum':0,'maximum':1})
 elif unit=='p.p.' or '%' in unit:dv({'validate':'decimal','criteria':'between','minimum':0,'maximum':1})
 elif k in ['commercial_delay','actual_month','n_platinum','n_gold','n_silver_plus','n_silver']:dv({'validate':'integer','criteria':'between','minimum':0,'maximum':180 if k in ['commercial_delay','actual_month'] else 100000})
 elif k in ['open','end','hotmonth','draw_end']:dv({'validate':'integer','criteria':'between','minimum':1,'maximum':180})
 elif isinstance(value,(int,float)):dv({'validate':'decimal','criteria':'>=','value':0})

s=sh('Cenarios',['Condição','Abertura','Objetivo','Limite aceitável','Personalizado','Ativo'],'Verde indica condições favoráveis. Limite aceitável é o teste bancário de referência; sem contrato aprovado.')
for i,(label,vals,key) in enumerate([('Taxa juro',[.05,.055,.06],'rate_custom'),('Carência meses',[48,48,42],'grace_custom'),('Amortização anos',[15,15,10],'term_custom'),('Comissão abertura',[.005,.0075,.01],'opening_custom'),('Penalização antecipação',[0,0,0],None)],5):
 text(s,i,0,label)
 for j,v in enumerate(vals,1):s.write(i-1,j,v,F['good_pct' if i in [5,8,9] else 'good_int'] if j<3 else F['pct'] if i in [5,8,9] else F['int'])
 if key:fw(s,i,4,pv(key),BASE['p'][key],'pct' if i in [5,8,9] else 'int')
 else:text(s,i,4,0,'pct')
 val=BASE['p'][['rate','grace','term','opening'][i-5]] if i<9 else 0
 fw(s,i,5,f'INDEX(B{i}:E{i},1,MATCH({pv("bank")},$B$4:$E$4,0))',val,'pct' if i in [5,8,9] else 'int')
s.write_row(11,0,['Cenário Hotel','ADR inicial','Ocupação','MICE / quartos','Outras / quartos'],F['head'])
for r,(name,vals) in enumerate(HOTELS.items(),13):
 text(s,r,0,name)
 for c,v in enumerate(vals,1):s.write(r-1,c,v,F['input_num' if c==1 else 'input_pct']);s.write_comment(r-1,c,'Hipótese editorial herdada da V45. Não é forecast nem garantia do Operador Internacional.')
s.write_row(17,0,['Hotel ativo','ADR','Ocupação','MICE / quartos','Outras / quartos'],F['head'])
fw(s,19,0,pv('hotel'),BASE['p']['hotel'],'text')
for c,key in enumerate(['adr','occ','mice','other'],1):fw(s,19,c,f'INDEX({col(c)}13:{col(c)}15,MATCH({pv("hotel")},A13:A15,0))',BASE['p'][key],'num' if c==1 else 'pct')

s=sh('Unidades',['ID','Tipologia','Piso','Área total m²','Interior m²','Varanda m²','Preço origem €','Preço a cêntimos €','Mês venda','Trimestre venda','Mês tranche 20%','Mês escritura','N.º semestres','Sinal €','Tranche €','55% €','Escritura €','Reconciliação €','Imposto €','Marketing €','Kyenu €','Preço cobrável €','Estado calendário','Venda cenário','Tranche cenário','Escritura cenário'],'94 frações preservadas. Editar venda/escritura e pricing nas células azuis. 504 / 810: escritura provisória, validar externamente.')
s.set_column('A:C',12);s.set_column('G:H',23);s.set_column('W:W',35)
for i,u in enumerate(UNITS):
 r=i+5
 for c,k in enumerate(['id','tipo','piso','area','interior','varanda','price_source']):s.write(r-1,c,u[k],F['input_text' if c<3 else 'input_num' if c<6 else 'input_eur'])
 fw(s,r,7,f'IF({pv("pricing_mode")} ="Motor",\'V04 Pricing\'!P{r},ROUND(G{r},2))',u['price']);text(s,r,8,u['sale'],'input_int');fw(s,r,9,f'ROUNDUP((I{r}+{pv("commercial_delay")})/3,0)',math.ceil(u['sale']/3),'int');fw(s,r,10,f'I{r}+{pv("lag")}',u['sale']+BASE['p']['lag'],'int');text(s,r,11,u['deed'],'input_int')
 n=(u['deed']-u['sale']-BASE['p']['lag']-1)//6
 fw(s,r,12,f'MAX(0,INT((Z{r}-Y{r}-1)/6))',n,'int')
 net=u['price']-(BASE['p']['conversion'] if u['id']==BASE['p']['conversion_id'] else 0)
 fw(s,r,21,f'H{r}-IF(A{r}={pv("conversion_id")},{pv("conversion")},0)',net)
 for c,k in [(13,'deposit'),(14,'tranche'),(15,'progress')]:fw(s,r,c,f'ROUND(V{r}*{pv(k)},2)',round(net*BASE['p'][k],2))
 fw(s,r,16,f'V{r}-SUM(N{r}:P{r})',round(net-round(net*.05,2)-round(net*.20,2)-round(net*.55,2),2))
 fw(s,r,17,f"SUM('Recebimentos'!B{r}:FY{r})-V{r}",0)
 for c,k in [(18,'tax_transfer'),(19,'marketing'),(20,'kyenu')]:fw(s,r,c,f'H{r}*{pv(k)}',u['price']*BASE['p'][k])
 fw(s,r,22,f'IF(OR(K{r}>L{r},I{r}<1,L{r}>180,AND(M{r}=0,P{r}>0)),"Alerta calendário",IF(AND(OR(A{r}="504",A{r}="810"),{pv("validate_deeds")}=0),"Pendente de Validação Externa","Conforme"))','Pendente de Validação Externa' if u['id'] in ['504','810'] else 'Conforme','text')
 for c,letter in [(23,'I'),(24,'K'),(25,'L')]:fw(s,r,c,f'{letter}{r}+{pv("commercial_delay")}' + ('+'+pv('deed_delay') if c==25 else ''),u['sale'] if c<25 else u['deed'],'int')
 for c in [6,8,11]:s.write_comment(r-1,c,'Origem: V45 Unidades Pricing linha '+str(u['source_row'])+'. Mês 29 de 504 e 810 não contratualmente validado.' if u['id'] in ['504','810'] else 'Origem V45; alterar apenas com suporte comercial/contratual.')
s.data_validation('I5:I98',{'validate':'integer','criteria':'between','minimum':1,'maximum':180});s.data_validation('L5:L98',{'validate':'integer','criteria':'between','minimum':1,'maximum':180});s.autofilter('A4:W98');alert(s,'W5:W98')
text(s,100,0,'TOTAL')
for c in [3,4,5,6,7,13,14,15,16,17,18,19,20,21]:fw(s,100,c,f'SUM({col(c)}5:{col(c)}98)',sum(u['price'] for u in UNITS) if c in [7,21] else 0,'total_eur' if c>=6 else 'total_num')

s=sh('Recebimentos',['ID']+list(range(1,181)),'Mapa individual mensal de 180 meses. 5% + 20% + 55% semestral + saldo final; arredondamento absorvido na última prestação.')
s.set_column(1,180,15)
for i,u in enumerate(UNITS):
 r=i+5;U=lambda c:vr('Unidades',r,c)
 fw(s,r,0,U(0),u['id'],'text')
 for m in range(1,181):
  prog=f'IF(AND({U(12)}>0,{m}>{U(24)},{m}<{U(25)},MOD({m}-{U(24)},6)=0,INT(({m}-{U(24)})/6)<={U(12)}),IF({m}={U(24)}+{U(12)}*6,{U(15)}-ROUND({U(15)}/{U(12)},2)*({U(12)}-1),ROUND({U(15)}/{U(12)},2)),0)'
  f=f'IF({m}={U(23)},{U(13)},0)+IF({m}={U(24)},{U(14)},0)+{prog}+IF({m}={U(25)},{U(16)},0)'
  fw(s,r,m,f'IF({U(23)}<={pv("sales_stop")},{f},0)',BASE['re']['units'][i][m-1])
text(s,100,0,'TOTAL')
for m in range(1,181):fw(s,100,m,f'SUM({col(m)}5:{col(m)}98)',BASE['re']['total'][m-1],'total_eur')
s.autofilter(3,0,97,180)

s=sh('Obra',['ID','Capítulo','Incidência','Montante base €']+[f'M{m} peso' for m in range(1,31)]+['Soma pesos'],'30 capítulos e pesos mensais originais preservados. Valores base = área × €/m² × incidência. Pesos não são normalizados silenciosamente.')
s.set_column('B:B',43)
for i,ch in enumerate(CHAPTERS):
 r=i+5;text(s,r,0,ch['id'],'int');text(s,r,1,ch['name']);text(s,r,2,ch['share'],'input_pct');fw(s,r,3,f'C{r}*{pv("area")}*{pv("cost")}*(1+{pv("cost_overrun")})',ch['share']*DEFAULT['area']*DEFAULT['cost'])
 for m,v in enumerate(ch['weights'],4):text(s,r,m,v,'input_pct')
 fw(s,r,34,f'SUM(E{r}:AH{r})',sum(ch['weights']),'pct')
 s.write_comment(r-1,2,'V45 Orcamento Capitulos linha '+str(i+4)+'; preservar orçamento e faseamento aprovado.')
text(s,36,1,'TOTAL')
fw(s,36,2,'SUM(C5:C34)',sum(c['share'] for c in CHAPTERS),'pct');fw(s,36,3,'SUM(D5:D34)',DEFAULT['area']*DEFAULT['cost'],'total_eur')
s.autofilter('A4:AI34')
s=sh('CAPEX',['Mês','Construção + FF&E','Fiscalização','Contingência paga','Total caixa obra'],'Investimento histórico excluído. Contingência não realizada e HOTUSA são reservas separadas, sem pagamento automático.')
for m,x in enumerate(BASE['cap'],1):
 r=m+4;text(s,r,0,m,'int')
 f=f"SUMPRODUCT('Obra'!$D$5:$D$34,'Obra'!{col(m+3)}$5:{col(m+3)}$34)" if m<=30 else '0'
 fw(s,r,1,f,x['base']);fw(s,r,2,f'B{r}*{pv("supervision")}',x['supervision']);fw(s,r,3,f'B{r}*{pv("contingency")}*{pv("cont_used")}',x['cont']);fw(s,r,4,f'SUM(B{r}:D{r})',x['total'])
table(s,184,4,'CAPEX')

REKEY=['m','deposit','tranche','progress','final','total','tax','marketing','kyenu','net','cum','sales']
s=sh('Imobiliario',['Mês','Sinal 5%','Tranche 20%','Semestrais 55%','Escrituras 20%','Recebimentos','Imposto transmissão','Marketing','Kyenu','Caixa líquido','Acumulado líquido','Vendas contratadas'],'Custos sobre preço total (inclui conversões); imposto na escritura; marketing e Kyenu no contrato. Incidência fiscal a confirmar.')
for m in range(1,181):
 r=m+4;text(s,r,0,m,'int');end=98
 for c,moncol,valcol,key in [(1,'X','N','deposit'),(2,'Y','O','tranche'),(4,'Z','Q','final'),(6,'Z','S','tax'),(7,'X','T','marketing'),(8,'X','U','kyenu'),(11,'X','H',None)]:
  value=BASE['re'][key][m-1] if key else sum(u['price'] for u in UNITS if u['sale']==m)
  fw(s,r,c,f'SUMIFS(\'Unidades\'!${valcol}$5:${valcol}$98,\'Unidades\'!${moncol}$5:${moncol}$98,A{r},\'Unidades\'!$X$5:$X$98,"<="&{pv("sales_stop")})',value)
 fw(s,r,5,f"'Recebimentos'!{col(m)}100",BASE['re']['total'][m-1]);fw(s,r,3,f'F{r}-B{r}-C{r}-E{r}',BASE['re']['progress'][m-1]);net=BASE['re']['total'][m-1]-sum(BASE['re'][k][m-1] for k in ['tax','marketing','kyenu']);fw(s,r,9,f'F{r}-SUM(G{r}:I{r})',net);fw(s,r,10,f'SUM(J$5:J{r})',sum(BASE['re']['total'][:m])-sum(sum(BASE['re'][k][:m]) for k in ['tax','marketing','kyenu']))
table(s,184,11,'Imobiliario')

HK=list(BASE['hotel'][0]);HIDX={k:i for i,k in enumerate(HK)}
HHEAD=['Mês','Ano operação','ADR','Ocupação','Quartos','MICE espaços','MICE banquetes','MICE serviços','F&B Hotel','Serviços Hotel','Receita total','Custo quartos','Custo MICE','Custo outras','Gerais não distribuídos','GOP','Fee base','Sistemas','Fee incentivo','EBITDA','Dotação FF&E','FF&E pago','Fundo FF&E restrito','CAPEX reposição','Imposto proxy','Fundo maneio','Variação FM','CFADS Hotel']
def hotel_formulas(r,scenario_expr=None,fees_override=None):
 a=lambda k:col(HIDX[k])+str(r);m=a('m');age=a('age');R=a('rooms');rev=a('revenue');gop=a('gop');base=a('base');systems=a('systems');inc=a('incentive');eb=a('ebitda');ffe=a('ffe');spent=a('ffe_spent');maint=a('maintenance');wc=a('wc')
 fee=lambda k:fees_override[{'base_fee':0,'systems':1,'incentive':2}[k]] if fees_override else pv(k)
 init=lambda c:f"'Cenarios'!${col(c)}$19" if scenario_expr is None else f'INDEX(\'Cenarios\'!${col(c)}$13:${col(c)}$15,MATCH({scenario_expr},\'Cenarios\'!$A$13:$A$15,0))'
 return {'age':f'IF({m}<{pv("open")},0,INT(({m}-{pv("open")})/12)+1)','adr':f'IF({age}=0,0,{init(1)}*(1+{pv("adr_growth")})^({age}-1))','occ':f'IF({age}=0,0,MIN({pv("occ_max")},{init(2)}+{pv("occ_step")}*({age}-1)))','rooms':f'{pv("rooms")}*365/12*{a("adr")}*{a("occ")}','mice_events':f'{R}*{init(3)}*{pv("mice_events")}','mice_banquet':f'{R}*{init(3)}*{pv("mice_banquet")}','mice_services':f'{R}*{init(3)}*(1-{pv("mice_events")}-{pv("mice_banquet")})','fb':f'{R}*{init(4)}*{pv("fb_share")}','services':f'{R}*{init(4)}*(1-{pv("fb_share")})','revenue':f'SUM({R}:{a("services")})','room_cost':f'{R}*{pv("room_cost")}','mice_cost':f'SUM({a("mice_events")}:{a("mice_services")})*{pv("mice_cost")}','other_cost':f'SUM({a("fb")}:{a("services")})*{pv("other_cost")}','overhead':f'{rev}*{pv("overhead")}','gop':f'{rev}-SUM({a("room_cost")}:{a("overhead")})','base':f'{rev}*{fee("base_fee")}','systems':f'{rev}*{fee("systems")}','incentive':f'MAX(0,{gop}-{base}-{systems})*{fee("incentive")}*IF(OR({pv("incentive_gate")}=0,AND({gop}*12>={pv("gop_floor")},({gop}-{base}-{systems})*12>={pv("ebitda_floor")})),1,0)','ebitda':f'{gop}-SUM({base}:{inc})','ffe':f'{rev}*{pv("ffe")}','ffe_spent':f'{ffe}*{pv("ffe_spend")}','ffe_fund':f'{a("ffe_fund")[:-len(str(r))]}{r-1}+{ffe}-{spent}' if r%1000!=5 and r!=5 else f'{ffe}-{spent}','maintenance':f'{rev}*{pv("maintenance")}','tax':f'MAX(0,{eb}-{spent}-{maint})*{pv("tax")}','wc':f'{rev}*12/365*{pv("wc_days")}','dwc':f'{wc}-{col(HIDX["wc"])}{r-1}' if r%1000!=5 and r!=5 else wc,'cf':f'{eb}-{ffe}-{maint}-{a("tax")}-{a("dwc")}'}
s=sh('Hotel',HHEAD,'DRE autónoma de 200 quartos + MICE. Fundo de maneio e perdas entram no caixa sem piso artificial zero. Aniversário operacional rege ADR e ocupação.')
for i,x in enumerate(BASE['hotel']):
 r=i+5;text(s,r,0,i+1,'int');fs=hotel_formulas(r)
 for k in HK[1:]:fw(s,r,HIDX[k],fs[k],x[k],'pct' if k=='occ' else 'int' if k=='age' else 'eur')
table(s,184,len(HK)-1,'Hotel')

RK=list(BASE['res'][0]);RIDX={k:i for i,k in enumerate(RK)}
def res_formulas(r):
 a=lambda k:col(RIDX[k])+str(r);m=a('m');age=f'MAX(0,INT(({m}-{pv("open")})/12))';rev=a('revenue');cash=a('cash');ex=a('excess');wc=a('wc');gross=a('kkgross')
 return {'adr':f'IF({m}<{pv("open")},0,{pv("res_adr")}*(1+{pv("res_growth")})^{age})','occ':f'IF({m}<{pv("open")},0,MIN({pv("res_occ_max")},{pv("res_occ")}+{pv("res_occ_step")}*{age}))','revenue':f'COUNTA(\'Unidades\'!$A$5:$A$98)*365/12*{a("adr")}*{a("occ")}*(1+{pv("res_other")})','opex':f'{rev}*{pv("res_cost")}','base':f'{rev}*{pv("res_base")}','systems':f'{rev}*{pv("res_system")}','ffe':f'{rev}*{pv("res_ffe")}','wc':f'{rev}*12/365*{pv("res_wc")}','dwc':f'{wc}-I{r-1}' if r!=5 else wc,'cash':f'{rev}-SUM({a("opex")}:{a("ffe")})-{a("dwc")}','target':f'IF({m}<{pv("open")},0,SUM(\'Unidades\'!$H$5:$H$98)*MIN({pv("res_yield_max")},{pv("res_yield")}+{age}*{pv("res_yield_step")})/12)','owner':f'MIN(MAX(0,{cash}),{a("target")})','short':f'MAX(0,{a("target")}-{a("owner")})','excess':f'MAX(0,{cash}-{a("owner")})','reserve':f'{ex}*{pv("res_fund")}','kyenu':f'{ex}*{pv("res_kyenu")}','operator':f'{ex}*{pv("res_operator")}','kkgross':f'{ex}-SUM({a("reserve")}:{a("operator")})+MIN(0,{cash})','tax':f'MAX(0,{gross})*{pv("tax")}','kk':f'({gross}-{a("tax")})*{pv("res_include")}','fund':f'{a("reserve")}+{a("ffe")}' if r==5 else f'V{r-1}+{a("reserve")}+{a("ffe")}'}
s=sh('Residencial',['Mês','ADR','Ocupação','Receita pool','OPEX','Fee base','Sistemas','Dotação FF&E','Fundo maneio','Variação FM','Caixa antes yield','Yield alvo €','Renda suportável','Yield não coberto','Excedente','Fundo adicional','Kyenu','Operador','Residual KK bruto','Imposto KK','Caixa KK consolidado','Fundos pool restritos'],'DRE e waterfall das 94 frações; não confundir com Hotel. Yield limitado ao disponível; residual KK e direitos dependem de validação contratual.')
for i,x in enumerate(BASE['res']):
 r=i+5;text(s,r,0,i+1,'int');fs=res_formulas(r)
 for k in RK[1:]:fw(s,r,RIDX[k],fs[k],x[k],'pct' if k=='occ' else 'eur')
table(s,184,len(RK)-1,'Residencial')

FK=list(BASE['flow'][0]);FI={k:i for i,k in enumerate(FK)}
FH=['Mês','Ano','Equity realizado','Recebimentos RE','Transmissão','Marketing','Kyenu','Construção','Fiscalização','Contingência paga','HOTUSA pago','CFADS Hotel','Caixa residual KK','Caixa inicial','Dívida inicial','Usado acumulado inicial','Reserva protegida','Caixa antes banco','1ª utilização','Abertura banco','Success fee broker','Mobilização','Capital obrigatório','Desembolso','Juros','Caixa após obrigatório','Amortização voluntária','Dívida final','Utilização acumulada','Dividendos','Caixa final','Défice não financiado','Falta para reserva','Serviço obrigatório','DSCR Hotel','FCF projeto','Fluxo equity total','Check caixa','Check dívida','EBITDA Hotel','Receita Hotel','Reserva HOTUSA','Contingência em reserva','Proteção futura obra','FCF projeto acumulado','Equity acumulado']

FH+=['Entrada empréstimo vendedores','Juro empréstimo vendedores','Amortização empréstimo vendedores','Saldo empréstimo vendedores','Custo compromisso','Lucros acumulados — proxy','Teto dividendos disponível','Dividendos acumulados','Crédito mobilizável no cenário']

def flow_formulas(r,first=5,last=184,equity_expr=None,hotel_sn='Hotel',hotel_offset=0,bank_override=None,capex_sn='CAPEX',re_sn='Imobiliario',cost_factor='1'):
 a=lambda k:col(FI[k])+str(r);pre=lambda k:col(FI[k])+str(r-1);m=a('m');idx=r-first;sr=idx+5;hr=sr+hotel_offset
 H=lambda k:vr(hotel_sn,hr,HIDX[k]);R=lambda k:vr('Residencial',sr,RIDX[k]);RE=lambda c:vr(re_sn,sr,c);CAP=lambda c:vr(capex_sn,sr,c)
 rate=bank_override[0] if bank_override else "'Cenarios'!$F$5";grace=bank_override[1] if bank_override else "'Cenarios'!$F$6";term=bank_override[2] if bank_override else "'Cenarios'!$F$7";opfee=bank_override[3] if bank_override else "'Cenarios'!$F$8"
 eq=equity_expr or pv('equity');debt=a('debt_open');avail=a('available');reserve=a('reserve');cum=a('draw_open');princ=a('principal');draw=a('draw');post=a('post');sweep=a('sweep');closing=a('debt');hot=a('hothold');cont=a('conthold');principal=f'IF({m}>{grace},MIN({debt},{debt}/MAX(1,{grace}+{term}*12-{m}+1)),0)'
 futurecap=f"SUM('{capex_sn}'!E{sr+1}:E184)" if sr<184 else '0';futureincome=f"SUM('{re_sn}'!J{sr+1}:J184)" if sr<184 else '0'
 nextcap=f'SUMIFS(\'{capex_sn}\'!$E$5:$E$184,\'{capex_sn}\'!$A$5:$A$184,">"&{m},\'{capex_sn}\'!$A$5:$A$184,"<="&({m}+{pv("buffer")}))'
 firstop=f'INDEX(\'{hotel_sn}\'!$AB${5+hotel_offset}:$AB${184+hotel_offset},{pv("open")} )';firstres=f'INDEX(\'Residencial\'!$U$5:$U$184,{pv("open")})'
 remain=f'IF({m}<={pv("draw_end")},MAX(0,{pv("limit")}-{cum}),0)';fees=f'({a("openingfee")}+{a("broker")}+{a("mobil")})'
 d=dict(year=f'ROUNDUP({m}/12,0)',equity=f'IF({m}=1,({eq}-{pv("conversion")})*{pv("treasury_ratio")},0)',receipts=RE(5),transfer=RE(6),marketing=RE(7),kyenu=RE(8),construction=CAP(1),supervision=CAP(2),contingency=CAP(3),hotpay=f'IF(AND({pv("hotpay")}=1,{m}={pv("hotmonth")}),{pv("hotusa")},0)',hotel_cf=H('cf'),res_cf=R('kk'),cash_open='0' if r==first else pre('cash'),debt_open='0' if r==first else pre('debt'),draw_open='0' if r==first else pre('cumulative'),reserve=f'MAX({pv("reserve")},IF({m}<={pv("end")},{nextcap},0))+{hot}+{cont}+IF({m}<={pv("end")},MAX(0,-{firstop}-{firstres}+{pv("limit")}*{rate}/12),0)',available=f'{a("cash_open")}+{a("equity")}+{a("receipts")}+{a("hotel_cf")}+{a("res_cf")}-SUM({a("transfer")}:{a("hotpay")})',trigger=f'IF(AND({cum}<0.005,{avail}-{princ}-{debt}*{rate}/12<{reserve}-0.005,{remain}>0),1,0)',openingfee=f'{a("trigger")}*{pv("limit")}*{opfee}',broker=f'{a("trigger")}*{pv("limit")}*{pv("broker")}',mobil=f'{a("trigger")}*{pv("limit")}*{pv("mobil")}',principal=principal,draw=f'MIN({remain},MAX(0,({reserve}-{avail}+{princ}+{debt}*{rate}/12+{fees})/(1-{rate}/12)))',interest=f'({debt}+{draw})*{rate}/12',post=f'{avail}+{draw}-{fees}-{a("interest")}-{princ}',sweep=f'MIN(MAX(0,{debt}+{draw}-{princ}),MAX(0,{post}-{reserve}-{a("forward")})*{pv("sweep")},IF({m}<={pv("end")},MAX(0,{RE(9)}),MAX(0,{post}-{reserve}-{a("forward")})))',debt=f'MAX(0,{debt}+{draw}-{princ}-{sweep})',cumulative=f'{cum}+{draw}',dividend=f'IF(AND({m}>={pv("open")},MOD({m},12)=0,{closing}<0.005,SUM({col(FI["gap"])}${first}:{a("gap")})<0.005),MAX(0,{post}-{sweep}-{reserve})*{pv("payout")},0)',cash=f'{post}-{sweep}-{a("dividend")}',gap=f'MAX(0,-{post})',reservegap=f'MAX(0,{reserve}-{a("cash")})',service=f'{princ}+{a("interest")}',dscr=f'IF({a("service")}>0,{a("hotel_cf")}/{a("service")},"")',project_cf=f'{a("receipts")}+{a("hotel_cf")}+{a("res_cf")}-SUM({a("transfer")}:{a("hotpay")})',equity_cf=f'{a("dividend")}-{a("equity")}',cash_check=f'{a("cash")}-{a("cash_open")}-{a("equity")}-{a("receipts")}-{a("hotel_cf")}-{a("res_cf")}-{draw}+SUM({a("transfer")}:{a("hotpay")})+{fees}+{a("interest")}+{princ}+{sweep}+{a("dividend")}',debt_check=f'{closing}-{debt}-{draw}+{princ}+{sweep}',hotel_ebitda=H('ebitda'),hotel_rev=H('revenue'),hothold=f'IF(OR({pv("hotpay")}=0,{m}<{pv("hotmonth")}),{pv("hotusa")},0)',conthold=f'{pv("area")}*{pv("cost")}*(1+{pv("cost_overrun")})*({cost_factor})*{pv("contingency")}*(1-{pv("cont_used")})',forward=f'IF({m}<{pv("end")},MAX(0,{futurecap}-{futureincome}-MAX(0,{pv("limit")}-{cum}-{draw})),0)',project_cum=f'{a("project_cf")}' if r==first else f'{pre("project_cum")}+{a("project_cf")}',equity_cum=f'{a("equity_cf")}' if r==first else f'{pre("equity_cum")}+{a("equity_cf")}')

 # V03 extension: staged treasury mobilisation and a separate repayable shareholder loan.
 weight='+'.join(f'IF({m}={pv("fund_month"+str(j))},{pv("fund_part"+str(j))},0)' for j in [1,2,3])
 d['equity']=f'({eq}-{pv("conversion")})*{pv("treasury_ratio")}*({weight})'
 d['loan_in']=f'IF({pv("funding_kind")}="Empréstimo",{a("equity")},0)'
 prevloan='0' if r==first else pre('shareloan')
 d['loan_interest']=f'({prevloan}+{a("loan_in")})*{pv("shareloan_rate")}/12'
 d['loan_principal']=f'IF({m}>={pv("shareloan_start")},MIN({prevloan}+{a("loan_in")},({prevloan}+{a("loan_in")})/MAX(1,{pv("shareloan_start")}+{pv("shareloan_term")}-{m})),0)'
 d['shareloan']=f'{prevloan}+{a("loan_in")}-{a("loan_principal")}'
 d['available']+=f'-{a("loan_interest")}-{a("loan_principal")}'
 d['service']+=f'+{a("loan_interest")}+{a("loan_principal")}'
 d['cash_check']+=f'+{a("loan_interest")}+{a("loan_principal")}'
 d['dividend']=d['dividend'].replace(f'{closing}<0.005',f'{closing}<0.005,{a("shareloan")}<0.005')
 return d

exec((ROOT/'src/v04_core.py').read_text(),globals())

def write_flow(s,result,first=5,**kw):
 for i,x in enumerate(result['flow']):
  r=first+i;text(s,r,0,i+1,'int');fs=flow_formulas(r,first,first+179,**kw)
  for k in FK[1:]:fw(s,r,FI[k],fs[k],x[k],'int' if k in ['year','trigger'] else 'mult' if k=='dscr' else 'eur')
s=sh('Fluxo 180M',FH,'Único motor bancário: desembolso por necessidade, juros sobre saldo usado, sem crédito automático além do limite. Défices permanecem visíveis.');write_flow(s,BASE);table(s,184,len(FK)-1,'Fluxo')
s.conditional_format('AF5:AG184',{'type':'cell','criteria':'>','value':.01,'format':F['bad']})

# Annual synthesis from same monthly engine.
AK=['year','receipts','construction','supervision','contingency','hotpay','equity','hotel_rev','hotel_ebitda','hotel_cf','res_cf','transfer','marketing','kyenu','draw','openingfee','broker','mobil','interest','principal','sweep','dividend','cash','debt','service','dscr','total_dscr','reserve','net_debt_ebitda','ltv','ffe_fund','res_fund','reserve_gap','gap']
AI={k:i for i,k in enumerate(AK)}
s=sh('Anual',['Ano','Recebimentos RE','Construção','Fiscalização','Contingência','HOTUSA pago','Equity','Receita Hotel','EBITDA Hotel','CFADS Hotel','Residual KK','Transmissão','Marketing','Kyenu','Desembolsos','Abertura','Broker','Mobilização','Juros','Capital obrigatório','Antecipações','Dividendos','Caixa final','Dívida final','Serviço obrigatório','DSCR Hotel','DSCR total','Reserva final','Dívida líq./EBITDA','LTV indicativo','Fundo FF&E Hotel','Fundos residencial','Falta reserva máxima','Défice máximo'],'Anos de projeto 1–15. DSCR anual apenas sobre serviço obrigatório; períodos sem serviço ficam sem rácio. Valorização não é garantia bancária.')
ANNUAL=[]
for y in range(1,16):
 r=y+4;fr=BASE['flow'][(y-1)*12:y*12];last=fr[-1];vals={k:sum(x[k] for x in fr) for k in AK if k in FI and k not in ['year','cash','debt','dscr','reserve','gap']};vals.update(year=y,cash=last['cash'],debt=last['debt'],reserve=last['reserve'],dscr=sum(x['hotel_cf'] for x in fr)/sum(x['service'] for x in fr) if sum(x['service'] for x in fr)>0 else None,total_dscr=(sum(x['hotel_cf']+x['res_cf'] for x in fr)/sum(x['service'] for x in fr)) if sum(x['service'] for x in fr)>0 else None,ffe_fund=BASE['hotel'][y*12-1]['ffe_fund'],res_fund=BASE['res'][y*12-1]['fund'],reserve_gap=max(x['reservegap'] for x in fr),gap=max(x['gap'] for x in fr));eb=sum(x['hotel_ebitda'] for x in fr);vals['net_debt_ebitda']=(last['debt']-last['cash'])/eb if eb>0 else None;vals['ltv']=last['debt']/(eb*DEFAULT['exit_mult']) if eb>0 else None;ANNUAL.append(vals)
 text(s,r,0,y,'int');start=5+(y-1)*12;end=start+11
 for k in AK[1:]:
  if k in ['cash','debt','reserve']:f=vr('Fluxo 180M',end,FI[k])
  elif k=='dscr':f=f'IF(Y{r}>0,J{r}/Y{r},"")'
  elif k=='total_dscr':f=f'IF(Y{r}>0,(J{r}+K{r})/Y{r},"")'
  elif k=='net_debt_ebitda':f=f'IF(I{r}>0,(X{r}-W{r})/I{r},"")'
  elif k=='ltv':f=f'IF(I{r}>0,X{r}/(I{r}*{pv("exit_mult")}),"")'
  elif k=='ffe_fund':f=vr('Hotel',end,HIDX['ffe_fund'])
  elif k=='res_fund':f=vr('Residencial',end,RIDX['fund'])
  elif k in ['reserve_gap','gap']:f=f'MAX(\'Fluxo 180M\'!{col(FI["reservegap" if k=="reserve_gap" else "gap"])}{start}:{col(FI["reservegap" if k=="reserve_gap" else "gap"])}{end})'
  else:f=f'SUM(\'Fluxo 180M\'!{col(FI[k])}{start}:{col(FI[k])}{end})'
  fw(s,r,AI[k],f,vals[k],'pct' if k=='ltv' else 'mult' if k in ['dscr','total_dscr','net_debt_ebitda'] else 'eur')
table(s,19,len(AK)-1,'Anual')

# Scenario calculation blocks: same financing formulas, selected common assumptions propagate.
sens=sh('Comparacao',['Equity M€','Hotel','Banco','Pico dívida','Dívida M30','Caixa M30','Caixa mínima','Falta reserva máx.','Défice máx.','Juros 180M','Dívida M180','Dividendos 180M','Estado'],'12 combinações equity × operação, mais comparação bancária e sensibilidade juros. Fórmulas recalculam com os pressupostos comuns; sem resultados fixos.')
sens.set_column('B:C',22)
SCEN=[]
for eq in [0,5e6,10e6,15e6,20e6]:
 for hotelname in HOTELS:SCEN.append(dict(equity=eq,hotel=hotelname,bank=None,label='Capital / operação'))
for bank in BANKS:SCEN.append(dict(equity=None,hotel=None,bank=bank,label='Negociação bancária'))
for rate in [.045,.05,.055,.06,.065,.07,.075,.08]:SCEN.append(dict(equity=None,hotel=None,bank='Personalizado',rate=rate,label='Juro '+str(rate)))
SCEN += [dict(equity=eq,hotel='Favorável',bank=None,equity_price=eq,label=f'{int(eq/1e6)}M por 40%') for eq in [15e6,20e6]]
for eq in [15e6,20e6]:
 for extra in [dict(delay=6,label='Atraso comercial +6M'),dict(overrun=.10,label='Sobrecusto +10%'),dict(hotel='Base',label='Hotel Base'),dict(hotel='Conservador',label='Hotel Conservador')]:
  SCEN.append(dict(equity=eq,hotel=extra.get('hotel','Favorável'),bank=None,equity_price=eq,**{k:v for k,v in extra.items() if k!='hotel'}))
for fee_kind in ['target','offer','limit','custom']:
 SCEN.append(dict(equity=None,hotel=None,bank=None,fee_kind=fee_kind,label='Operador '+fee_kind))
SCEN += [dict(equity=eq,equity_price=eq,equity_stake=stake,hotel='Favorável',bank=None,label=label) for label,eq,stake in [('KK-E15',15e6,.4),('KK-E20-A',20e6,.4),('KK-E20-B',20e6,.5),('KK-E25',25e6,.50)]]
for extra in [dict(delay=6,label='KK-E25 atraso +6M'),dict(overrun=.1,label='KK-E25 sobrecusto +10%'),dict(hotel='Base',label='KK-E25 Hotel Base'),dict(hotel='Conservador',label='KK-E25 Hotel Conservador')]:
 SCEN.append(dict(equity=25e6,equity_price=25e6,equity_stake=.50,bank=None,hotel=extra.pop('hotel','Favorável'),**extra))
SCRESULT=[]
for si,sc in enumerate(SCEN):
 row=si+5;eq=DEFAULT['equity'] if sc['equity'] is None else sc['equity'];hn=DEFAULT['hotel'] if sc['hotel'] is None else sc['hotel'];bn=DEFAULT['bank'] if sc['bank'] is None else sc['bank'];changes=dict(equity=eq,hotel=hn,bank=bn)
 if sc['equity'] is not None:changes['equity_mode']='Capital selecionado'
 if 'rate' in sc:changes['rate_custom']=sc['rate']
 if 'equity_price' in sc:changes['equity_price']=sc['equity_price']
 if 'equity_stake' in sc:changes['equity_stake']=sc['equity_stake']
 if 'delay' in sc:changes['commercial_delay']=DEFAULT['commercial_delay']+sc['delay']
 if 'overrun' in sc:changes['cost_overrun']=(1+DEFAULT['cost_overrun'])*(1+sc['overrun'])-1
 if 'fee_kind' in sc:
  changes['fee_mode']={'target':'Objetivo KK','offer':'Proposta simulada','limit':'Limites propostos','custom':'Personalizado'}[sc['fee_kind']]
 rr=run(config(**changes));SCRESULT.append(rr);sm=summary(rr)
 fw(sens,row,0,pv('equity')+'/1000000',eq/1e6,'num') if sc['equity'] is None else text(sens,row,0,eq/1e6,'num')
 fw(sens,row,1,pv('hotel'),hn,'text') if sc['hotel'] is None else text(sens,row,1,hn)
 fw(sens,row,2,pv('bank'),bn,'text') if sc['bank'] is None else text(sens,row,2,bn if 'rate' not in sc else sc['label'])
 # Each scenario has an operational Hotel block and financing block, hidden to keep review manageable.
 hsname=f'CalcH{si+1:02d}';fsname=f'CalcF{si+1:02d}'
 hs=sh(hsname,HHEAD,'Bloco de cenário; mesmas fórmulas do Hotel.');hs.hide()
 for i,x in enumerate(rr['hotel']):
  r=i+5;text(hs,r,0,i+1,'int');fs=hotel_formulas(r,vr('Comparacao',row,1,True),[(f'IF(ISNUMBER({vr("Condicoes Operador",j,3,True)}),{vr("Condicoes Operador",j,3,True)},{vr("Condicoes Operador",j,2,True)})' if sc['fee_kind']=='limit' else vr('Condicoes Operador',j,{'target':2,'offer':1,'custom':4}[sc['fee_kind']],True)) for j in [5,6,7]] if 'fee_kind' in sc else None)
  for k in HK[1:]:fw(hs,r,HIDX[k],fs[k],x[k],'pct' if k=='occ' else 'int' if k=='age' else 'eur')
 fs=sh(fsname,FH,'Bloco de cenário; mesmo motor de dívida.');fs.hide()
 bankexpr=None
 if sc['bank'] is not None:
  if sc['bank'] in BANKS:
   c=list(BANKS).index(sc['bank'])+1;bankexpr=[vr('Cenarios',r,c,True) for r in [5,6,7,8]]
  else:bankexpr=[str(sc['rate']),pv('grace_custom'),pv('term_custom'),pv('opening_custom')]
 # Scenario CAPEX/receipts wrappers preserve the single original engine and common base sheets.
 capsn='CAPEX';resn='Imobiliario';resisn='Residencial'
 if 'overrun' in sc:
  capsn=f'CalcC{si+1:02d}';cs=sh(capsn,['Mês','Base','Fiscalização','Contingência','Total'],'Stress adicional por fator; pesos e calendário preservados.');cs.hide()
  for mi,x in enumerate(rr['cap'],5):
   text(cs,mi,0,mi-4,'int')
   for cc,key in [(1,'base'),(2,'supervision'),(3,'cont'),(4,'total')]:fw(cs,mi,cc,vr('CAPEX',mi,cc)+f'*(1+{sc["overrun"]})',x[key])
 if 'delay' in sc:
  resn=f'CalcR{si+1:02d}';rs=sh(resn,['Mês']+list(REKEY[1:]),'Desfasamento comercial adicional; não altera a obra nem constitui aprovação contratual.');rs.hide()
  for mi in range(5,185):
   text(rs,mi,0,mi-4,'int')
   for cc in range(1,12):fw(rs,mi,cc,'0' if mi-5<sc['delay'] else vr('Imobiliario',mi-sc['delay'],cc),0)
 if 'delay' in sc:
  resisn=f'CalcA{si+1:02d}';ar=sh(resisn,['Mês']+RK[1:],'Disponibilidade por escritura no cenário de atraso.');ar.hide()
  for ai,ax in enumerate(rr['res']):
   arw=ai+5;text(ar,arw,0,ai+1,'int');arf=res_formulas(arw,sc['delay'])
   for ak in RK[1:]:fw(ar,arw,RIDX[ak],arf[ak],ax[ak],'eur')
 write_flow(fs,rr,residential_sn=resisn,equity_expr=f"'Comparacao'!$A${row}*1000000",hotel_sn=hsname,bank_override=bankexpr,capex_sn=capsn,re_sn=resn,cost_factor=str(1+sc.get('overrun',0)))
 for c,key in [(3,'peak'),(4,'debt30'),(5,'cash30'),(6,'cash_min'),(7,'reserve_gap'),(8,'funding_gap'),(9,'interest'),(10,'debt_end'),(11,'dividends')]:
  if key=='peak':f=f'MAX(\'{fsname}\'!AB5:AB184,\'{fsname}\'!O5:O184)' # gross intramonth peak separately formula SUM opening+draw helper not present
  elif key in ['debt30','cash30']:f=vr(fsname,34,FI['debt' if key=='debt30' else 'cash'])
  elif key=='cash_min':f=f'MIN(\'{fsname}\'!AE5:AE184)'
  elif key in ['reserve_gap','funding_gap']:f=f'MAX(\'{fsname}\'!{ "AG" if key=="reserve_gap" else "AF"}5:{"AG" if key=="reserve_gap" else "AF"}184)'
  elif key=='debt_end':f=vr(fsname,184,FI['debt'])
  else:f=f'SUM(\'{fsname}\'!{ "Y" if key=="interest" else "AD"}5:{"Y" if key=="interest" else "AD"}184)'
  if key=='peak':f='MAX('+','.join(f"'{fsname}'!O{r}+'{fsname}'!X{r}" for r in range(5,185))+')'
  fw(sens,row,c,f,sm[key])
 fw(sens,row,12,f'IF(OR(H{row}>0.01,I{row}>0.01),"Alerta liquidez","Conforme cenário")','Alerta liquidez' if sm['reserve_gap']>.01 or sm['funding_gap']>.01 else 'Conforme cenário','text')
text(sens,55,0,'Nota');text(sens,55,1,'Juros personalizados usam carência/prazo personalizados. Banco não renovável; défices não são financiados ficticiamente.');sens.set_row(54,30);alert(sens,f'M5:M{len(SCEN)+4}')

# KPI and shareholder sections completed below.
# Returns: 10 and 15 years, valuation explicitly separated from cash.
s=sh('Retornos',['Indicador','10 anos','15 anos','Definição / limitação'],'Saída estimada por múltiplo EBITDA Hotel, não contratualmente realizável. Caixa somada uma vez; dívida e reservas contingentes deduzidas.')
s.set_column('A:A',48);s.set_column('D:D',100)
RETROWS={'ebitda':5,'ev':6,'exit_cost':7,'debt':8,'cash':9,'contingent':10,'exit':11,'dividend':12,'share':13,'new_exit':14,'new_div':15,'new_total':16,'new_profit':17,'new_moic':18,'new_irr':19,'promoter_share':20,'promoter_total':21,'promoter_moic':22,'promoter_irr':23,'net_debt':24,'payout_capital':25,'npv_equity':26,'legacy_share_jscv':27,'legacy_share_jals':28,'feasible':29}
RET={}
for c,horizon in [(1,120),(2,180)]:
 sm=summary(BASE,horizon);fr=BASE['flow'];end=horizon+4;start=end-11;last=fr[horizon-1];B=lambda row:col(c)+str(row)
 vals={'ebitda':sum(x['ebitda'] for x in BASE['hotel'][horizon-12:horizon]),'ev':sm['ev'],'exit_cost':sm['ev']*DEFAULT['exit_cost'],'debt':last['debt'],'cash':last['cash'],'contingent':last['hothold']+last['conthold'],'exit':sm['exit_equity'],'dividend':sm['dividends'],'share':sm['share'],'new_exit':sm['exit_equity']*sm['share'],'new_div':sm['dividends']*sm['share'],'new_total':(sm['exit_equity']+sm['dividends'])*sm['share'],'new_profit':(sm['exit_equity']+sm['dividends'])*sm['share']-sm['equity'],'new_moic':sm['moic'],'new_irr':sm['equity_irr'],'promoter_share':1-sm['share'],'promoter_total':(sm['exit_equity']+sm['dividends'])*(1-sm['share']),'promoter_moic':(sm['exit_equity']+sm['dividends'])*(1-sm['share'])/DEFAULT['historic'],'promoter_irr':irr([-DEFAULT['historic']]+[x['dividend']*(1-sm['share'])+(sm['exit_equity']*(1-sm['share']) if x['m']==horizon else 0) for x in fr[:horizon]]),'net_debt':last['debt']-last['cash'],'payout_capital':0,'npv_equity':sum((-x['equity']+x['dividend']*sm['share'])/(1+DEFAULT['equity_discount'])**(x['m']/12) for x in fr[:horizon])+sm['exit_equity']*sm['share']/(1+DEFAULT['equity_discount'])**(horizon/12),'legacy_share_jscv':DEFAULT['legacy_jscv'],'legacy_share_jals':1-DEFAULT['legacy_jscv']-sm['share'],'feasible':'Conforme cenário'};RET[horizon]=vals
 formulas={'ebitda':f'SUM(\'Hotel\'!T{start}:T{end})','ev':f'{B(5)}*{pv("exit_mult")}','exit_cost':f'{B(6)}*{pv("exit_cost")}','debt':vr('Fluxo 180M',end,FI['debt'])+'+'+vr('Fluxo 180M',end,FI['shareloan']),'cash':vr('Fluxo 180M',end,FI['cash']),'contingent':f"'Fluxo 180M'!AP{end}+'Fluxo 180M'!AQ{end}",'exit':f'MAX(0,{B(6)}-{B(7)}-{B(8)}+{B(9)}-{B(10)})','dividend':f'SUM(\'Fluxo 180M\'!AD5:AD{end})','share':f'({pv("equity")}-{pv("conversion")})/{pv("equity_price")}*{pv("equity_stake")}','new_exit':f'{B(11)}*{B(13)}','new_div':f'{B(12)}*{B(13)}','new_total':f'{B(14)}+{B(15)}','new_profit':f'{B(16)}-({pv("equity")}-{pv("conversion")})','new_moic':f'IF({pv("equity")}>{pv("conversion")},{B(16)}/({pv("equity")}-{pv("conversion")}),"")','new_irr':f'IF({pv("equity")}>{pv("conversion")},IFERROR((1+IRR(\'Fluxos Retorno\'!{ "B" if horizon==120 else "C"}5:{"B" if horizon==120 else "C"}{horizon+5},0.005))^12-1,"TIR não definida"),"")','promoter_share':f'1-{B(13)}','promoter_total':f'({B(11)}+{B(12)})*{B(20)}','promoter_moic':f'IF({pv("historic")}>0,{B(21)}/{pv("historic")},"")','promoter_irr':f'IFERROR((1+IRR(\'Fluxos Retorno\'!{ "D" if horizon==120 else "E"}5:{"D" if horizon==120 else "E"}{horizon+5},0.005))^12-1,"TIR não definida")','net_debt':f'{B(8)}-{B(9)}','payout_capital':'0','npv_equity':f'NPV((1+{pv("equity_discount")})^(1/12)-1,\'Fluxos Retorno\'!{ "B" if horizon==120 else "C"}6:{"B" if horizon==120 else "C"}{horizon+5})','legacy_share_jscv':pv('legacy_jscv'),'legacy_share_jals':f'1-{pv("legacy_jscv")}-{B(13)}','feasible':f'IF(OR(MAX(\'Fluxo 180M\'!AF5:AG{end})>0.01,{pv("equity")}<{pv("conversion")}),"Alerta — retorno condicionado","Conforme cenário")'}
 for k,r in RETROWS.items():fw(s,r,c,formulas[k],vals[k],'pct' if k in ['share','new_irr','promoter_share','promoter_irr','legacy_share_jscv','legacy_share_jals'] else 'mult' if k.endswith('moic') else 'text' if k=='feasible' else 'eur')
labels=['EBITDA Hotel últimos 12 meses','Enterprise value Hotel','Custos estimados saída','Dívida no horizonte','Caixa no horizonte','Reservas contingentes deduzidas','Valor acionista estimado saída','Dividendos totais pagos','Participação novos investidores','Saída novos investidores','Dividendos novos investidores','Total novos investidores','Lucro novos investidores','MOIC novos investidores','TIR novos investidores','Participação promotores','Total promotores','MOIC promotores (base histórica proxy)','TIR promotores (data histórica proxy)','Dívida líquida','Reembolso contratual de capital','VPL novos investidores','JSCV pós entrada (ilustrativo)','JALS pós entrada (ilustrativo)','Condição de execução']
for r,label in enumerate(labels,5):text(s,r,0,label)
for r,note in {5:'Exclui residencial. Não usar avaliação contabilística como fluxo.',11:'EV líquido de custos − dívida + caixa − reservas contingentes. Fundos pool/FF&E não somados.',19:'TIR mensal anualizada; aporte mês 1, dividendos e saída no horizonte.',22:'4,5 M€ colocados artificialmente no mês 0: proxy comparativa, não retorno histórico certificado.',23:'Datas históricas reais indisponíveis; TIR não reproduz investimento efetivo do Grupo.',25:'Não existe reembolso garantido ou contratual modelado.',27:'Diluição proporcional ilustrativa; revisão de direitos e classes societárias pendente.',29:'Uma estimativa positiva não sana insuficiência de reservas nem constitui decisão de investimento.'}.items():text(s,r,3,note)
alert(s,'B29:C29')

s=sh('Fluxos Retorno',['Mês','Novos 10A','Novos 15A','Promotores 10A','Promotores 15A','Projeto sem saída','Projeto com saída','Equity agregado incl. histórico','Mudanças sinal projeto','Caixa acumulado novos sem saída'],'Vetores usados nas TIR/VPL. Valores de saída só entram no horizonte selecionado, sem alterar a tesouraria operacional.')
for m in range(0,181):
 r=m+5;text(s,r,0,m,'int');fr=BASE['flow'][m-1] if m else None
 for c,horizon,group in [(1,120,'new'),(2,180,'new'),(3,120,'promoter'),(4,180,'promoter')]:
  rc='B' if horizon==120 else 'C';share=f"'Retornos'!${rc}${13 if group=='new' else 20}";exit=f"'Retornos'!${rc}$11";sr=m+4
  if group=='new':f='0' if m==0 or m>horizon else f"-IF(A{r}=1,{pv('equity')}-{pv('conversion')},0)+'Fluxo 180M'!AD{sr}*{share}"+(f'+{exit}*{share}' if m==horizon else '')
  else:f=f'-{pv("historic")}' if m==0 else '0' if m>horizon else f"'Fluxo 180M'!AD{sr}*{share}"+(f'+{exit}*{share}' if m==horizon else '')
  v=0 if m>horizon else (-DEFAULT['historic'] if m==0 and group=='promoter' else 0 if m==0 else ((-fr['equity'] if group=='new' else 0)+fr['dividend']*RET[horizon]['share' if group=='new' else 'promoter_share']+(RET[horizon]['exit']*RET[horizon]['share' if group=='new' else 'promoter_share'] if m==horizon else 0)))
  fw(s,r,c,f,v)
 f='0' if m==0 else vr('Fluxo 180M',m+4,FI['project_cf']);fw(s,r,5,f,fr['project_cf'] if m else 0)
 terminal=f"+'Retornos'!$C$6-'Retornos'!$C$7-'Retornos'!$C$10" if m==180 else ''
 fw(s,r,6,f'F{r}'+terminal,(fr['project_cf'] if m else 0)+(RET[180]['ev']-RET[180]['exit_cost']-RET[180]['contingent'] if m==180 else 0))
 fw(s,r,7,f'-{pv("historic")}' if m==0 else f"'Fluxo 180M'!AK{m+4}"+("+'Retornos'!$C$11" if m==180 else ''),-DEFAULT['historic'] if m==0 else fr['equity_cf']+(RET[180]['exit'] if m==180 else 0))
 fw(s,r,8,'0' if m<=1 else f'IF(AND(F{r}<>0,F{r-1}<>0,SIGN(F{r})<>SIGN(F{r-1})),1,0)',0,'int')
 fw(s,r,9,'0' if m==0 else f'J{r-1}-IF(A{r}=1,{pv("equity")}-{pv("conversion")},0)+\'Fluxo 180M\'!AD{m+4}*\'Retornos\'!$C$13',0)

s=sh('Investidores',['Ticket €','Participação','Elegibilidade','Divid. acumulados 10A','Saída 10A','Total 10A','Lucro 10A','MOIC 10A','TIR 10A','Divid. acumulados 15A','Saída 15A','Total 15A','Lucro 15A','MOIC 15A','TIR 15A']+[f'Divid. ano {y}' for y in range(1,16)]+['Capital reembolsado'],'Tickets de 1 a 20 M€ como participações na ronda selecionada; não são quinze rondas independentes. Ticket superior ao capital da ronda fica indisponível.')
for i in range(1,21):
 r=i+4;ticket=i*1e6;text(s,r,0,ticket,'eur');share=ticket/DEFAULT['equity_price']*DEFAULT['equity_stake'];eligible=ticket<=DEFAULT['equity']-DEFAULT['conversion'];fw(s,r,1,f'A{r}/{pv("equity_price")}*{pv("equity_stake")}',share,'pct');fw(s,r,2,f'IF(A{r}<={pv("equity")}-{pv("conversion")},"Na ronda","Fora da ronda")','Na ronda' if eligible else 'Fora da ronda','text')
 for c,horizon,measure in [(3,120,'dividend'),(4,120,'exit'),(9,180,'dividend'),(10,180,'exit')]:
  rc='B' if horizon==120 else 'C';rr=12 if measure=='dividend' else 11;fw(s,r,c,f'IF(C{r}="Na ronda",\'Retornos\'!${rc}${rr}*B{r},"")',RET[horizon][measure]*share if eligible else None)
 for horizon,cols in [(120,[5,6,7,8]),(180,[11,12,13,14])]:
  totalc,profitc,moicc,irrc=cols;dc=3 if horizon==120 else 9;ec=dc+1
  for c,f,v,fmt in [(totalc,f'{col(dc)}{r}+{col(ec)}{r}',(RET[horizon]['dividend']+RET[horizon]['exit'])*share,'eur'),(profitc,f'{col(totalc)}{r}-A{r}',(RET[horizon]['dividend']+RET[horizon]['exit'])*share-ticket,'eur'),(moicc,f'{col(totalc)}{r}/A{r}',(RET[horizon]['dividend']+RET[horizon]['exit'])*share/ticket,'mult'),(irrc,f'\'Retornos\'!${"B" if horizon==120 else "C"}$19',RET[horizon]['new_irr'],'pct')]:fw(s,r,c,f'IF(C{r}="Na ronda",{f},"")',v if eligible else None,fmt)
 for y in range(1,16):fw(s,r,14+y,f'IF(C{r}="Na ronda",\'Anual\'!V{y+4}*B{r},"")',ANNUAL[y-1]['dividend']*share if eligible else None)
 text(s,r,30,0,'eur')
s.autofilter('A4:AE24')

s=sh('Societario',['Bloco','Antes','Depois ilustrativo','Estado / decisão necessária'],'Ponte de capital sem presumir manutenção automática de direitos. Classes históricas sem quantidades documentadas não são inventadas.')
s.set_column('A:A',28);s.set_column('D:D',110)
for r,label,old,f,v,note in [(5,'JSCV',.51,f'(1-\'Retornos\'!C13)*{pv("legacy_jscv")}',.306,'Diluição pro rata ilustrativa; validar deliberação e direitos.'),(6,'JALS',.49,f'(1-\'Retornos\'!C13)*(1-{pv("legacy_jscv")} )',.294,'Diluição pro rata ilustrativa; validar deliberação e direitos.'),(7,'Novos investidores',0,"'Retornos'!C13",.4,'Proporção negocial linear: 40% / 15 M€, não capitalização independente.')]:
 text(s,r,0,label);text(s,r,1,old,'pct');fw(s,r,2,f,v,'pct');text(s,r,3,note)
for r,block in enumerate(['Platinum','Gold','Silver Plus','Silver'],10):text(s,r,0,block);text(s,r,3,'Pendente de Validação Externa — quantidades, votos, preferências e conversibilidade carecem dos estatutos e deliberações.')
text(s,16,0,'Conversão em frações');text(s,16,3,'Alternativa de liquidez, sem garantia criada: abate recebimentos e aporte novo no cenário; exige acordo de cancelamento/transferência de ações.')

KPI=[]
def kp(key,label,f,val,fmt='eur',note=''):KPI.append((key,label,f,val,fmt,note))
SUM=summary(BASE);vgv=sum(u['price'] for u in UNITS);capbase=sum(x['total'] for x in BASE['cap']);costcom=sum(sum(BASE['re'][k]) for k in ['tax','marketing','kyenu']);proj=[0]+[x['project_cf'] for x in BASE['flow']];projexit=proj.copy();projexit[-1]+=RET[180]['ev']-RET[180]['exit_cost']-RET[180]['contingent']
kp('vgv','VGV 94 frações',"SUM('Unidades'!H5:H98)",vgv)
kp('area','Área vendável m²',"SUM('Unidades'!D5:D98)",sum(u['area'] for u in UNITS),'num')
kp('pm2','Preço médio €/m²',"B5/B6",vgv/sum(u['area'] for u in UNITS))
kp('commercial','Custos comerciais',"SUM('Imobiliario'!G5:I184)",costcom)
kp('netre','Vendas líquidas custos comerciais',"B5-B8",vgv-costcom)
kp('capex','CAPEX futuro pago',"SUM('CAPEX'!E5:E184)",capbase)
kp('rescapex','CAPEX imputado residencial',f'B10*{pv("res_capex_share")}',capbase*DEFAULT['res_capex_share'],note='Proxy por área; sem orçamento segmentado validado.')
kp('remargin','Resultado promoção antes finanças e imposto',"B9-B11",vgv-costcom-capbase*DEFAULT['res_capex_share'])
kp('remarginpct','Margem promoção proxy',"B12/B5",(vgv-costcom-capbase*DEFAULT['res_capex_share'])/vgv,'pct','Não inclui imposto de promoção nem alocação financeira; validar imputação.')
kp('peak','Pico intramês dívida','MAX('+','.join(f"'Fluxo 180M'!O{r}+'Fluxo 180M'!X{r}" for r in range(5,185))+')',SUM['peak'])
kp('debt30','Dívida mês 30',"'Fluxo 180M'!AB34",SUM['debt30'])
kp('cash30','Caixa mês 30',"'Fluxo 180M'!AE34",SUM['cash30'])
kp('debt42','Dívida mês 42',"'Fluxo 180M'!AB46",SUM['debt42'])
kp('cash42','Caixa mês 42',"'Fluxo 180M'!AE46",SUM['cash42'])
kp('interest','Juros 180 meses',"SUM('Fluxo 180M'!Y5:Y184)",SUM['interest'])
kp('fees','Comissões bancárias e broker',"SUM('Fluxo 180M'!T5:V184)",SUM['fees'])
kp('cashmin','Caixa mínima 180 meses',"MIN('Fluxo 180M'!AE5:AE184)",SUM['cash_min'])
kp('reservegap','Insuficiência máxima reserva',"MAX('Fluxo 180M'!AG5:AG184)",SUM['reserve_gap'])
kp('fundgap','Défice máximo não financiado',"MAX('Fluxo 180M'!AF5:AF184)",SUM['funding_gap'])
kp('debtend','Dívida final',"'Fluxo 180M'!AB184",SUM['debt_end'])
kp('cashend','Caixa final',"'Fluxo 180M'!AE184",SUM['cash_end'])
kp('npv','VPL projeto sem saída',f'NPV((1+{pv("discount")})^(1/12)-1,\'Fluxos Retorno\'!F6:F185)',SUM['npv_project'])
kp('npvexit','VPL projeto com saída',f'NPV((1+{pv("discount")})^(1/12)-1,\'Fluxos Retorno\'!G6:G185)',SUM['npv_project_exit'])
kp('projectirr','TIR projeto sem saída','IFERROR((1+IRR(\'Fluxos Retorno\'!F5:F185,0.01))^12-1,"TIR não definida")',irr(proj),'pct','Fluxos não convencionais: pode haver múltiplas raízes; teste numérico separado. Não usar isoladamente.')
kp('projectirrexit','TIR projeto com saída','IFERROR((1+IRR(\'Fluxos Retorno\'!G5:G185,0.01))^12-1,"TIR não definida")',irr(projexit),'pct','Inclui valor terminal estimado, não garantido.')
kp('equityirr','TIR novos investidores 15A',"'Retornos'!C19",RET[180]['new_irr'],'pct')
kp('equitymoic','MOIC novos investidores 15A',"'Retornos'!C18",RET[180]['new_moic'],'mult')
kp('equitynpv','VPL novos investidores 15A',"'Retornos'!C26",RET[180]['npv_equity'])
kp('dscrstub','DSCR ano inicial parcial',f'INDEX(\'Anual\'!$Z$5:$Z$19,ROUNDUP({pv("open")}/12,0))',ANNUAL[2]['dscr'],'mult','Ano 3 inclui juros de construção; analisar período parcial separadamente.')
full=[v['dscr'] for v in ANNUAL if v['year']>=4 and v['dscr'] is not None]
kp('dscrfull','DSCR mínimo anos completos operação',f'IF(COUNTIFS(\'Anual\'!$A$5:$A$19,">="&ROUNDUP(({pv("open")}+11)/12,0),\'Anual\'!$Y$5:$Y$19,">0")>0,_xlfn.MINIFS(\'Anual\'!$Z$5:$Z$19,\'Anual\'!$A$5:$A$19,">="&ROUNDUP(({pv("open")}+11)/12,0),\'Anual\'!$Y$5:$Y$19,">0"),"")',min(full) if full else None,'mult')
kp('headroom','Margem linha / pico',f'{pv("limit")}-B14',DEFAULT['limit']-SUM['peak'])
kp('safety','Margem segurança liquidez',"MIN('Fluxo 180M'!AE5:AE184-'Fluxo 180M'!Q5:Q184)",min(x['cash']-x['reserve'] for x in BASE['flow']),note='Caixa menos reserva; mínimo mensal. Não é margem comercial.')
kp('payback','Payback projeto em meses (última recuperação)',"IF('Fluxo 180M'!AS184<0,\"Não recuperado\",IFERROR(LOOKUP(2,1/('Fluxo 180M'!AS5:AS184<0),'Fluxo 180M'!A5:A184)+1,0))",next((x['m']+1 for x in reversed(BASE['flow']) if x['project_cum']<0),0),'int','Último mês com FCF acumulado negativo + 1; sem valor terminal, evita primeira recuperação transitória.')
kp('used','Utilização bancária acumulada',"MAX('Fluxo 180M'!AC5:AC184)",BASE['flow'][-1]['cumulative'])
kp('unused','Linha não utilizada (não renovável)',f'{pv("limit")}-MAX(\'Fluxo 180M\'!AC5:AC184)',DEFAULT['limit']-BASE['flow'][-1]['cumulative'],note='Este é o saldo realmente disponível da facilidade, sujeito à janela de utilização; não confundir com limite menos pico da dívida.')
kp('investor_payback','Payback novos investidores sem saída (mês)',"IF('Fluxos Retorno'!J185<0,\"Não recuperado\",IFERROR(LOOKUP(2,1/('Fluxos Retorno'!J6:J185<0),'Fluxos Retorno'!A6:A185)+1,0))",next((x['m'] for x in BASE['flow'] if sum(y['dividend']*SUM['share'] for y in BASE['flow'][:x['m']])>=DEFAULT['equity']),None),'int', 'Recuperação por dividendos apenas, sem venda estimada; participação da ronda selecionada.')
KROW={k:i+5 for i,(k,*_) in enumerate(KPI)}
s=sh('KPI',['Indicador','Resultado','Definição / limite'],'KPI com bases consistentes. TIR e valorização são estimativas de cenário; sustentabilidade deve cumprir reservas, dívida e cenários menos favoráveis.')
s.set_column('A:A',52);s.set_column('B:B',24);s.set_column('C:C',108)
for k,label,f,v,fmt,note in KPI:
 r=KROW[k];text(s,r,0,label)
 if k=='safety':
  # Avoid legacy dynamic-array formula: minimum of explicit scalar margins.
  f='MIN('+','.join(f"'Fluxo 180M'!AE{j}-'Fluxo 180M'!Q{j}" for j in range(5,185))+')'
 fw(s,r,1,f,v,fmt);text(s,r,2,note);s.write_comment(r-1,1,note or label+' — ligado ao motor mensal; euros nominais, não certificado externamente.')

# Operator fee matrix and economic counterfactual.
s=sh('Operador',['Fee base','Fee incentivo','Receitas 12M','GOP 12M','Gestão €','Sistemas €','EBITDA €','CFADS €','Gestão / CFADS','Sistemas / CFADS','Total / CFADS','Estado'],'Matriz 4×3. Usa primeiros 12 meses de operação; inclui fundo de maneio de arranque. Bases de fees dependem do HMA; contrafactual não garante valor acrescentado.')
firstop=int(DEFAULT['open'])+4;lastop=firstop+11;h12=BASE['hotel'][int(DEFAULT['open'])-1:int(DEFAULT['open'])+11];rev=sum(x['revenue'] for x in h12);gop=sum(x['gop'] for x in h12);ff=sum(x['ffe'] for x in h12);sp=sum(x['ffe_spent'] for x in h12);main=sum(x['maintenance'] for x in h12);dwc=sum(x['dwc'] for x in h12)
row=5
for basefee in [.04,.035,.03,.025]:
 for incentive in [.10,.08,.06]:
  r=row;row+=1;text(s,r,0,basefee,'pct');text(s,r,1,incentive,'pct');fw(s,r,2,f'SUMIFS(\'Hotel\'!$K$5:$K$184,\'Hotel\'!$A$5:$A$184,">="&{pv("open")},\'Hotel\'!$A$5:$A$184,"<"&({pv("open")}+12))',rev);fw(s,r,3,f'SUMIFS(\'Hotel\'!$P$5:$P$184,\'Hotel\'!$A$5:$A$184,">="&{pv("open")},\'Hotel\'!$A$5:$A$184,"<"&({pv("open")}+12))',gop)
  management=rev*basefee+max(0,gop-rev*basefee-rev*DEFAULT['systems'])*incentive;system=rev*DEFAULT['systems'];eb=gop-management-system;cf=eb-ff-main-max(0,eb-sp-main)*DEFAULT['tax']-dwc
  fw(s,r,4,f'C{r}*A{r}+MAX(0,D{r}-C{r}*A{r}-C{r}*{pv("systems")})*B{r}*IF(OR({pv("incentive_gate")}=0,AND(D{r}>={pv("gop_floor")},D{r}-C{r}*A{r}-C{r}*{pv("systems")}>={pv("ebitda_floor")})),1,0)',management);fw(s,r,5,f'C{r}*{pv("systems")}',system);fw(s,r,6,f'D{r}-E{r}-F{r}',eb)
  ffe=f'SUMIFS(\'Hotel\'!$U$5:$U$184,\'Hotel\'!$A$5:$A$184,">="&{pv("open")},\'Hotel\'!$A$5:$A$184,"<"&({pv("open")}+12))';maint=f'SUMIFS(\'Hotel\'!$X$5:$X$184,\'Hotel\'!$A$5:$A$184,">="&{pv("open")},\'Hotel\'!$A$5:$A$184,"<"&({pv("open")}+12))';wc=f'SUMIFS(\'Hotel\'!$AA$5:$AA$184,\'Hotel\'!$A$5:$A$184,">="&{pv("open")},\'Hotel\'!$A$5:$A$184,"<"&({pv("open")}+12))'
  fw(s,r,7,f'G{r}-{ffe}-{maint}-MAX(0,G{r}-{ffe}*{pv("ffe_spend")}-{maint})*{pv("tax")}-{wc}',cf)
  for c,num in [(8,f'E{r}'),(9,f'F{r}'),(10,f'E{r}+F{r}')]:fw(s,r,c,f'IF(H{r}>0,({num})/H{r},"")',management/cf if c==8 else system/cf if c==9 else (management+system)/cf,'pct')
  fw(s,r,11,f'IF(OR(H{r}<=0,I{r}>={pv("management_alert")},J{r}>={pv("systems_alert")},K{r}>={pv("total_alert")}),"Alerta","Conforme cenário")','Alerta' if cf<=0 or management/cf>=DEFAULT['management_alert'] or system/cf>=DEFAULT['systems_alert'] or (management+system)/cf>=DEFAULT['total_alert'] else 'Conforme cenário','text')
text(s,20,0,'Valor do operador');text(s,20,1,'Contrafactual hipotético; não decisão contratual. Gestão alternativa não orçamentada.');s.merge_range('B20:H20','Contrafactual hipotético; não decisão contratual. Gestão alternativa não orçamentada.',F['subtitle'])
text(s,21,0,'CFADS cenário com operador')
fw(s,21,1,f'SUMIFS(\'Hotel\'!$AB$5:$AB$184,\'Hotel\'!$A$5:$A$184,">="&({pv("open")}+12),\'Hotel\'!$A$5:$A$184,"<"&({pv("open")}+24))',sum(x['cf'] for x in BASE['hotel'][42:54]))
text(s,22,0,'Receita sem operador');fw(s,22,1,f'SUMIFS(\'Hotel\'!$K$5:$K$184,\'Hotel\'!$A$5:$A$184,">="&({pv("open")}+12),\'Hotel\'!$A$5:$A$184,"<"&({pv("open")}+24))*{pv("value_adr")}*{pv("value_occ")}',0)
text(s,23,0,'GOP sem operador');fw(s,23,1,f'SUMIFS(\'Hotel\'!$P$5:$P$184,\'Hotel\'!$A$5:$A$184,">="&({pv("open")}+12),\'Hotel\'!$A$5:$A$184,"<"&({pv("open")}+24))*{pv("value_adr")}*{pv("value_occ")}',0)
text(s,24,0,'CFADS sem operador proxy');fw(s,24,1,f'(B23-B22*({pv("ffe")}+{pv("maintenance")}))-MAX(0,B23-B22*({pv("ffe")}*{pv("ffe_spend")}+{pv("maintenance")}))*{pv("tax")}',0)
text(s,25,0,'Caixa incremental estimada');fw(s,25,1,'B21-B24',0)
text(s,26,0,'Limiar receita sem operador');text(s,26,1,'Comparação indicativa em ano maduro, sem variação FM contrafactual; custos de gestão alternativa ainda pendentes.')
alert(s,'L5:L16')
# Financial sensitivities and exit value scenarios, linked to active operations and debt.
s=sh('Sensibilidade',['ADR multiplicador','Ocupação multiplicador','Receita Hotel 12M','EBITDA Hotel 12M','CFADS anual proxy','DSCR ativo proxy','Estado'],'Sensibilidade ADR × ocupação, sem novo financiamento. Proxy de ano completo, FM estabilizado; reexecutar cenário completo ao alterar pressupostos.')
row=5
for adrm in [.9,1,1.1]:
 for occm in [.9,1,1.1]:
  r=row;row+=1;text(s,r,0,adrm,'mult');text(s,r,1,occm,'mult');fw(s,r,2,f'SUMIFS(\'Hotel\'!$K$5:$K$184,\'Hotel\'!$A$5:$A$184,">="&({pv("open")}+12),\'Hotel\'!$A$5:$A$184,"<"&({pv("open")}+24))*A{r}*B{r}',0)
  # All revenues and variable costs scale; no independent fixed-cost assertion.
  fw(s,r,3,f'SUMIFS(\'Hotel\'!$T$5:$T$184,\'Hotel\'!$A$5:$A$184,">="&({pv("open")}+12),\'Hotel\'!$A$5:$A$184,"<"&({pv("open")}+24))*A{r}*B{r}',0)
  fw(s,r,4,f'D{r}-C{r}*({pv("ffe")}+{pv("maintenance")})-MAX(0,D{r}-C{r}*({pv("ffe")}*{pv("ffe_spend")}+{pv("maintenance")}))*{pv("tax")}',0)
  fw(s,r,5,f'IF(SUMIFS(\'Fluxo 180M\'!$AH$5:$AH$184,\'Fluxo 180M\'!$A$5:$A$184,">="&({pv("open")}+12),\'Fluxo 180M\'!$A$5:$A$184,"<"&({pv("open")}+24))>0,E{r}/SUMIFS(\'Fluxo 180M\'!$AH$5:$AH$184,\'Fluxo 180M\'!$A$5:$A$184,">="&({pv("open")}+12),\'Fluxo 180M\'!$A$5:$A$184,"<"&({pv("open")}+24)),"")',0,'mult')
  fw(s,r,6,f'IF(OR(E{r}<0,AND(ISNUMBER(F{r}),F{r}<{pv("dscr_min")})),"Alerta","Conforme proxy")','Conforme proxy','text')
s.write_row(16,0,['Saída múltiplo','EV 10A','Equity value 10A','EV 15A','Equity value 15A'],F['head'])
for r,mult in enumerate([7,9,11],18):
 text(s,r,0,mult,'input_mult')
 for c,rc in [(1,'B'),(3,'C')]:
  fw(s,r,c,f'A{r}*\'Retornos\'!{rc}5',RET[120 if rc=='B' else 180]['ebitda']*mult)
  fw(s,r,c+1,f'MAX(0,{col(c)}{r}*(1-{pv("exit_cost")})-\'Retornos\'!{rc}8+\'Retornos\'!{rc}9-\'Retornos\'!{rc}10)',0)
alert(s,'G5:G13')

# Dashboard helper: dynamic chart references; no embedded pictures of KPIs.
def dash(name,note):
 s=sh(name,note=note);s.set_column('A:A',43);s.set_column('B:B',25);s.set_column('C:C',20);s.set_column('D:J',16);s.set_zoom(80);s.freeze_panes(4,0);s.set_tab_color(GREEN);return s

def metric(s,r,label,f,v=None,fmt='eur',note=''):
 text(s,r,0,label);fw(s,r,1,f,v,fmt);s.write_comment(r-1,1,note or 'Indicador dinâmico ligado ao motor de cálculo. Consultar KPI, Pressupostos e AUDITORIA.')
def chart(s,anchor,typ,title,categories,series,width=700,height=300):
 ch=wb.add_chart({'type':typ});ch.set_title({'name':title,'name_font':{'size':12,'color':NAVY}});ch.set_legend({'position':'bottom'});ch.set_size({'width':width,'height':height});ch.set_chartarea({'border':{'none':True}});ch.set_plotarea({'border':{'none':True}});ch.set_y_axis({'major_gridlines':{'visible':True,'line':{'color':'#E8EDF2'}},'num_format':'#,##0'});ch.set_x_axis({'name':'Ano' if categories[0]=='Anual' else 'Mês / categoria'});colors=[TEAL,'#5689C6','#E7A542','#87A83C','#9C6686']
 for j,(name,values) in enumerate(series):ch.add_series({'name':name,'categories':categories,'values':values,'line':{'color':colors[j%len(colors)],'width':2},'fill':{'color':colors[j%len(colors)]},'border':{'none':True}})
 s.insert_chart(anchor,ch)
def kref(k):return vr('KPI',KROW[k],1)

s=dash('DASH Real Estate','Promoção imobiliária | pricing e calendário de 94 frações | margens de imputação provisória | custos ligados ao caixa')
for r,(label,key,fmt) in enumerate([('VGV 94 frações','vgv','eur'),('Área vendável m²','area','num'),('Preço médio €/m²','pm2','eur'),('Custos comerciais','commercial','eur'),('Receitas líquidas','netre','eur'),('CAPEX residencial imputado','rescapex','eur'),('Resultado promoção proxy','remargin','eur'),('Margem antes finanças/imposto','remarginpct','pct')],5):metric(s,r,label,kref(key),fmt=fmt)
metric(s,14,'Recebimentos 30 meses',"SUM('Imobiliario'!F5:F34)");metric(s,15,'Recebimentos após mês 30',"SUM('Imobiliario'!F35:F184)");text(s,17,0,'Alocação custos residencial');fw(s,17,1,f'IF({pv("res_capex_share")}>0,"Pendente de Validação Externa","Alerta")','Pendente de Validação Externa','text');alert(s,'B17')
chart(s,'D5','line','Recebimentos e custos de obra — 30 meses',['Imobiliario',4,0,33,0],[('Recebimentos',['Imobiliario',4,5,33,5]),('CAPEX total projeto',['CAPEX',4,4,33,4])])
chart(s,'D22','line','Recuperação de caixa líquida de vendas',['Imobiliario',4,0,33,0],[('Acumulado líquido',['Imobiliario',4,10,33,10])])

s=dash('DASH Comercial','Filtros combináveis por ID, tipologia, piso, mês e trimestre. Selecionar Todos/Todas. Recebimentos, custos e gráficos reagem ao mesmo detalhe.')
# Filters vertical, avoiding ambiguity of merged title/header cells.
filters=[('ID','Todos'),('Tipologia','Todas'),('Piso','Todos'),('Mês venda','Todos'),('Trimestre','Todos')]
for r,(label,value) in enumerate(filters,5):text(s,r,0,label);text(s,r,1,value,'input_text');s.write_comment(r-1,1,'Filtro comercial. Escolher Todos/Todas para remover restrição; não altera pressupostos económicos.')
s.data_validation('B5',{'validate':'list','source':['Todos']+[u['id'] for u in UNITS]}) if False else None
# Excel validation list stored in hidden support cells to avoid 255-character list limits.
ls=sh('Listas',['ID','Tipologia','Piso','Meses','Trimestres'],'Apoio às validações dos filtros.');ls.hide()
lists=[['Todos']+[u['id'] for u in UNITS],['Todas']+sorted(set(u['tipo'] for u in UNITS)),['Todos']+sorted(set(u['piso'] for u in UNITS)),['Todos']+list(range(1,181)),['Todos']+list(range(1,61))]
for c,arr in enumerate(lists):
 for rr,v in enumerate(arr,5):ls.write(rr-1,c,v)
 s.data_validation(c+4,1,c+4,1,{'validate':'list','source':f"='Listas'!${col(c)}$5:${col(c)}${len(arr)+4}"})
# Detail table (filtered totals use explicit selection flag, also has normal Excel filters).
DETAIL_FIRST=52;DETAIL_LAST=145
s.write_row(50,0,['ID','Tipo','Piso','Área','Preço','Mês','Trimestre','Imposto','Kyenu','Marketing','Selecionada','Preço/m²'],F['head'])
for i,u in enumerate(UNITS):
 r=DETAIL_FIRST+i;ur=i+5
 for c,uc in [(0,0),(1,1),(2,2),(3,3),(4,7),(5,23),(6,9),(7,18),(8,20),(9,19)]:fw(s,r,c,vr('Unidades',ur,uc),u['id'] if c==0 else u['tipo'] if c==1 else u['piso'] if c==2 else 0,'text' if c<3 else 'eur' if c in [4,7,8,9] else 'num')
 fw(s,r,10,f'IF(AND(OR($B$5="Todos",A{r}=$B$5),OR($B$6="Todas",B{r}=$B$6),OR($B$7="Todos",C{r}=$B$7),OR($B$8="Todos",F{r}=$B$8),OR($B$9="Todos",G{r}=$B$9)),1,0)',1,'int')
 fw(s,r,11,f'IF(D{r}>0,E{r}/D{r},0)',u['price']/u['area'])
s.autofilter(50,0,144,11)
metric(s,11,'Frações selecionadas','SUM(K52:K145)',94,'int')
for r,label,letter,fmt in [(12,'VGV selecionado','E','eur'),(13,'Área selecionada m²','D','num'),(15,'Imposto transmissão','H','eur'),(16,'Comissão Kyenu','I','eur'),(17,'Marketing','J','eur')]:metric(s,r,label,f'SUMPRODUCT($K$52:$K$145,{letter}52:{letter}145)',fmt=fmt)
metric(s,14,'Preço médio €/m²','IF(B13>0,B12/B13,0)');metric(s,18,'Líquido selecionado','B12-SUM(B15:B17)')
# Dynamic 30-month receipts chart, with filters applied to each fraction.
s.write_row(149,0,['Mês','Vendas','Contrato 25%','Semestrais 55%','Escrituras 20%','Total','Acumulado','Trimestre','Vendas trimestre','Receb. trimestre'],F['head'])
for m in range(1,181):
 r=150+m;text(s,r,0,m,'int');letter=col(m)
 fw(s,r,1,f'SUMPRODUCT($K$52:$K$145,--($F$52:$F$145=A{r}),$E$52:$E$145)',sum(u['price'] for u in UNITS if u['sale']==m))
 fw(s,r,2,f'SUMPRODUCT($K$52:$K$145,--(\'Unidades\'!$X$5:$X$98=A{r}),\'Unidades\'!$N$5:$N$98)+SUMPRODUCT($K$52:$K$145,--(\'Unidades\'!$Y$5:$Y$98=A{r}),\'Unidades\'!$O$5:$O$98)',BASE['re']['deposit'][m-1]+BASE['re']['tranche'][m-1])
 fw(s,r,4,f'SUMPRODUCT($K$52:$K$145,--(\'Unidades\'!$Z$5:$Z$98=A{r}),\'Unidades\'!$Q$5:$Q$98)',BASE['re']['final'][m-1]);fw(s,r,5,f'SUMPRODUCT($K$52:$K$145,\'Recebimentos\'!{letter}$5:{letter}$98)',BASE['re']['total'][m-1]);fw(s,r,3,f'F{r}-C{r}-E{r}',BASE['re']['progress'][m-1]);fw(s,r,6,f'SUM(F$151:F{r})',sum(BASE['re']['total'][:m]));fw(s,r,7,f'ROUNDUP(A{r}/3,0)',math.ceil(m/3),'int')
 if m<=60:
  fw(s,r,8,f'SUMIF($H$151:$H$330,A{r},$B$151:$B$330)',0);fw(s,r,9,f'SUMIF($H$151:$H$330,A{r},$F$151:$F$330)',0)
# Type/floor summaries react to current filters.
s.write_row(33,0,['Grupo','VGV','Área','Imposto','Kyenu','Marketing'],F['head'])
for r,(group,kind) in enumerate([(t,'B') for t in sorted(set(u['tipo'] for u in UNITS))]+[(p,'C') for p in sorted(set(u['piso'] for u in UNITS))],35):
 text(s,r,0,group)
 for c,letter in [(1,'E'),(2,'D'),(3,'H'),(4,'I'),(5,'J')]:fw(s,r,c,f'SUMPRODUCT($K$52:$K$145,--(${kind}$52:${kind}$145=A{r}),${letter}$52:${letter}$145)',0,'num' if c==2 else 'eur')
chart(s,'D5','column','Recebimentos selecionados — 30 meses',['DASH Comercial',150,0,179,0],[('25% inicial',['DASH Comercial',150,2,179,2]),('55% semestral',['DASH Comercial',150,3,179,3]),('20% escritura',['DASH Comercial',150,4,179,4])])
chart(s,'H22','line','Curva acumulada selecionada',['DASH Comercial',150,0,179,0],[('Acumulado',['DASH Comercial',150,6,179,6])],600,260);chart(s,'H38','column','Vendas e recebimentos trimestrais',['DASH Comercial',150,0,159,0],[('Vendas',['DASH Comercial',150,8,159,8]),('Recebimentos',['DASH Comercial',150,9,159,9])],600,230)
s.print_area('A1:Q50')

s=dash('DASH Hotel MICE','Hotel de 200 quartos + MICE de 1.500 pessoas | DRE própria | operação Favorável selecionável | sem incorporar receita de 294 unidades')
metric(s,5,'Cenário ativo',pv('hotel'),DEFAULT['hotel'],'text');metric(s,6,'ADR inicial',"'Cenarios'!B19",DEFAULT['adr'] if 'adr' in DEFAULT else 210);metric(s,7,'Ocupação inicial',"'Cenarios'!C19",.64,'pct')
for r,label,hc in [(8,'Receita ano completo','K'),(9,'GOP ano completo','P'),(10,'EBITDA ano completo','T'),(11,'Gestão e incentivo','Q'),(12,'Sistemas','R'),(13,'CFADS ano completo','AB')]:
 f=f'SUMIFS(\'Hotel\'!${hc}$5:${hc}$184,\'Hotel\'!$A$5:$A$184,">="&({pv("open")}+12),\'Hotel\'!$A$5:$A$184,"<"&({pv("open")}+24))'
 if r==11:f+='+'+f.replace('$Q','$S')
 metric(s,r,label,f)
metric(s,14,'Gestão / CFADS','IF(B13>0,B11/B13,"")',fmt='pct');metric(s,15,'Sistemas / CFADS','IF(B13>0,B12/B13,"")',fmt='pct');metric(s,16,'Total operador / CFADS','IF(B13>0,(B11+B12)/B13,"")',fmt='pct');metric(s,17,'Estado encargos',f'IF(OR(B13<=0,B14>={pv("management_alert")},B15>={pv("systems_alert")},B16>={pv("total_alert")}),"Alerta","Conforme cenário")',fmt='text');metric(s,19,'DSCR ano parcial',kref('dscrstub'),fmt='mult');metric(s,20,'DSCR anos completos',kref('dscrfull'),fmt='mult');metric(s,21,'Residual residencial incluído',pv('res_include'),1,'int');alert(s,'B17')
chart(s,'D5','column','Hotel + MICE — desempenho anual',['Anual',4,0,18,0],[('Receitas',['Anual',4,7,18,7]),('EBITDA',['Anual',4,8,18,8]),('CFADS',['Anual',4,9,18,9])]);chart(s,'D22','line','Cobertura do serviço obrigatório',['Anual',4,0,18,0],[('DSCR Hotel',['Anual',4,25,18,25])])

s=dash('DASH Acionistas','Retorno económico a 10 e 15 anos | valor de saída hipotético | ponte societária ilustrativa | sem reembolso ou yield garantido')
for r,label,f,fmt in [(5,'Capital novo líquido',f'{pv("equity")}-{pv("conversion")}','eur'),(6,'Participação novos',"'Retornos'!C13",'pct'),(7,'Participação promotores',"'Retornos'!C20",'pct'),(8,'TIR novos 10A',"'Retornos'!B19",'pct'),(9,'TIR novos 15A',"'Retornos'!C19",'pct'),(10,'MOIC novos 10A',"'Retornos'!B18",'mult'),(11,'MOIC novos 15A',"'Retornos'!C18",'mult'),(12,'Dividendos novos 15A',"'Retornos'!C15",'eur'),(13,'Saída novos 15A estimada',"'Retornos'!C14",'eur'),(14,'Múltiplo saída',pv('exit_mult'),'mult'),(15,'EV Hotel 15A estimado',"'Retornos'!C6",'eur'),(16,'VPL novos investidores',"'Retornos'!C26",'eur'),(18,'Condição cenário',"'Retornos'!C29",'text')]:metric(s,r,label,f,fmt=fmt)
chart(s,'D5','column','Dividendos distribuídos por ano',['Anual',4,0,18,0],[('Dividendos totais',['Anual',4,21,18,21])]);chart(s,'D22','column','Valorização estimada por múltiplo',['Sensibilidade',17,0,19,0],[('Equity value 10A',['Sensibilidade',17,2,19,2]),('Equity value 15A',['Sensibilidade',17,4,19,4])]);alert(s,'B18')

s=dash('DASH Consolidado','180 meses | 15 anos | cenário bancário Limite aceitável por referência | caixa, dívida e reservas ligados ao mesmo motor')
for r,label,key,fmt in [(5,'CAPEX futuro','capex','eur'),(7,'Pico de dívida','peak','eur'),(8,'Dívida mês 30','debt30','eur'),(9,'Caixa mês 30','cash30','eur'),(10,'Dívida mês 42','debt42','eur'),(11,'Caixa mínima','cashmin','eur'),(12,'Falta máxima de reserva','reservegap','eur'),(13,'Défice não financiado','fundgap','eur'),(14,'Juros acumulados','interest','eur'),(15,'Comissões / broker','fees','eur'),(16,'VPL projeto sem saída','npv','eur'),(17,'TIR projeto sem saída','projectirr','pct'),(18,'VPL projeto com saída','npvexit','eur'),(19,'TIR novos investidores','equityirr','pct')]:metric(s,r,label,kref(key),fmt=fmt)
metric(s,6,'Capital próprio novo',pv('equity'));metric(s,23,'Linha não utilizada',kref('unused'));metric(s,24,'Utilização acumulada',kref('used'));metric(s,21,'Estado liquidez',f'IF(OR({kref("reservegap")}>0.01,{kref("fundgap")}>0.01),"Alerta","Conforme cenário")',fmt='text');alert(s,'B21')
chart(s,'D5','line','Dívida e caixa mensais',['Fluxo 180M',4,0,183,0],[('Dívida',['Fluxo 180M',4,27,183,27]),('Caixa',['Fluxo 180M',4,30,183,30]),('Reserva',['Fluxo 180M',4,16,183,16])]);chart(s,'D22','column','Serviço obrigatório e amortização voluntária',['Anual',4,0,18,0],[('Juros',['Anual',4,18,18,18]),('Capital obrigatório',['Anual',4,19,18,19]),('Antecipações',['Anual',4,20,18,20])])

# Audit sheet contains live reconciliations and documentary pending items.
s=sh('AUDITORIA',['Controlo','Resultado','Estado','Critério / ação'],'Conforme = teste técnico de cenário. Alerta = desvio. Pendente de Validação Externa = falta contrato, fonte ou decisão; não significa aprovação.')
s.set_column('A:A',48);s.set_column('B:B',26);s.set_column('C:C',38);s.set_column('D:D',110)
AUD=[]
def aud(label,f,criterion,note,fmt='eur',external=False):AUD.append((label,f,criterion,note,fmt,external))
aud('94 IDs únicos',"COUNTA('Unidades'!A5:A98)",'B{r}=94','Contagem e unicidade também testadas automaticamente.','int')
aud('VGV vs referência',"SUM('Unidades'!H5:H98)-50656336.37",'ABS(B{r})<0.01','Preço individual arredondado a cêntimos; preço origem preservado.')
aud('Recebimentos vs cobrável',"SUM('Imobiliario'!F5:F184)-SUM('Unidades'!V5:V98)",'ABS(B{r})<0.01','100% por unidade; última prestação/escritura absorve arredondamento.')
aud('Reconciliação individual máxima',"MAX('Unidades'!R5:R98)",'ABS(B{r})<0.01','Testes externos verificam também mínimo absoluto e cada unidade.')
aud('Percentuais contrato somam 100%',f'{pv("deposit")}+{pv("tranche")}+{pv("progress")}+{pv("final")}-1','ABS(B{r})<0.000001','Saldo escritura é complemento, mas percentuais definidos têm de reconciliar.','pct')
aud('Incidências orçamento',"SUM('Obra'!C5:C34)-1",'ABS(B{r})<0.000001','Sem normalização oculta.','pct')
aud('Calendários capítulos',"SUMPRODUCT(ABS('Obra'!AI5:AI34-1))",'ABS(B{r})<0.000001','Cada capítulo tem de somar 100%.','pct')
aud('Área individual reconciliada',"SUM('Unidades'!D5:D98)-SUM('Unidades'!E5:F98)",'ABS(B{r})<0.01','Área total = interior + varanda.','num')
aud('Check caixa máximo','MAX('+','.join(f"ABS('Fluxo 180M'!AL{r})" for r in range(5,185))+')','ABS(B{r})<0.01','Identidade de caixa todos os meses.')
aud('Check dívida máximo','MAX('+','.join(f"ABS('Fluxo 180M'!AM{r})" for r in range(5,185))+')','ABS(B{r})<0.01','Dívida inicial + desembolso − principal − antecipação = dívida final.')
aud('Limite não renovável',f'MAX(\'Fluxo 180M\'!AC5:AC184)-{pv("limit")}','B{r}<=0.01','Desembolsos acumulados nunca superam linha.')
aud('Défice não financiado',kref('fundgap'),'B{r}<0.01','Défice não é coberto por fonte fictícia.')
aud('Insuficiência reservas',kref('reservegap'),'B{r}<0.01','Inclui tesouraria mínima, obra, pré-abertura e riscos reservados.')
aud('Marketing mínimo',pv('marketing'),'B{r}>=0.01','Mínimo 1%.','pct')
aud('DSCR ano parcial',kref('dscrstub'),f'B{{r}}>={pv("dscr_min")}','Ano de abertura inclui juros de construção; não confundir com ano completo.','mult')
aud('DSCR anos completos',kref('dscrfull'),f'OR(B{{r}}="",B{{r}}>={pv("dscr_min")})','Somente serviço obrigatório; voluntário excluído.','mult')
aud('Mês final construção',pv('end'),'B{r}=30','Cronograma por capítulos é de 30 meses; revisão estrutural necessária para outra duração.','int')
aud('Abertura após obra',f'{pv("open")}-{pv("end")}','B{r}>=1','Convenção relativa. Antecipar abertura requer faseamento de obra e licenças.','int')
aud('Yield residencial não coberto',"SUM('Residencial'!N5:N184)",'B{r}<0.01','Yield é objetivo sem garantia; falta de cobertura deve permanecer visível.')
aud('Fee operador / caixa',"'DASH Hotel MICE'!B16",f'AND(ISNUMBER(B{{r}}),B{{r}}<{pv("total_alert")})','Alerta ao atingir ou superar 50%.','pct')
aud('Conversão dentro da fração',f'{pv("conversion")}-SUMIF(\'Unidades\'!A5:A98,{pv("conversion_id")},\'Unidades\'!H5:H98)','B{r}<=0.01','Conversão não pode superar preço da fração selecionada.')
aud('Conversão dentro capital novo',f'{pv("conversion")}-{pv("equity")}','B{r}<=0.01','Sem capital ou cobranças negativos.')
aud('Escrituras 504 e 810',pv('validate_deeds'),'B{r}=1','Mês 29 herdado é apenas hipótese, sem validação contratual.','int',True)
for label,note in [('Imposto transmissão / IVA','Confirmar sujeito passivo, incidência, IVA de obra, vendas e exploração e recuperabilidade; 1,5% é hipótese.'),('HOTUSA','Reserva de risco 1 M€ no cenário prudente; não reconhece dívida certa nem pagamento obrigatório.'),('Contrato Operador Internacional','Fees, base de incentivo, distribuição/sistemas, custos e performance test precisam de HMA.'),('Financiamento bancário','Term sheet, condições de utilização, carência, comissões e antecipação ainda a negociar.'),('Forecast Hotel e MICE','Cenários editoriais; categorias MICE repartidas por proxy, sem contratos/eventos que validem volume.'),('Direitos residual residencial','45% do excedente KK antes de imposto, hipótese herdada; contratos sell & leaseback e waterfall por validar.'),('Valorização e preço de saída','Múltiplos EBITDA editáveis são estimativas, não comprador ou preço contratualmente realizável.'),('Sociedade e classes de ações','Ponte JSCV/JALS pro rata ilustrativa; direitos Platinum, Gold, Silver Plus e Silver dependem de estatutos.'),('Fiscalidade e custos promoção','Imputação residencial por área é proxy; imposto promoção, perdas fiscais, depreciações e taxas locais carecem parecer.'),('Investimento histórico e retorno Grupo','4,5 M€ sem datas efetivas; TIR promotores é proxy com aporte no mês 0, não TIR histórica comprovada.')]:aud(label,'"Por confirmar"','FALSE',note,'text',True)
for r,(label,f,criterion,note,fmt,external) in enumerate(AUD,5):
 text(s,r,0,label);fw(s,r,1,f,None,fmt)
 state=f'IF({criterion.format(r=r)},"Conforme","Pendente de Validação Externa")' if external else f'IF({criterion.format(r=r)},"Conforme","Alerta")'
 fw(s,r,2,state,'Pendente de Validação Externa' if external else 'Conforme','text');text(s,r,3,note)
alert(s,f'C5:C{4+len(AUD)}');s.autofilter(3,0,3+len(AUD),3)

s=sh('Origem e Alteracoes',['Módulo V45','Tratamento','Motivo / efeito'],'Mapa de migração: original integral preservado fora do novo motor. Não se mantêm motores antigos concorrentes no ficheiro final.')
s.set_column('A:A',36);s.set_column('B:B',42);s.set_column('C:C',110)
changes=[('Unidades Pricing','Unidades + Recebimentos','94 IDs, tipos, pisos, áreas e preços origem preservados; pricing operacional a cêntimos; datas editáveis.'),('Orcamento Capitulos','Obra + CAPEX','Incidências e 30 pesos por capítulo preservados; histórico não sai de caixa novamente.'),('V22 / KK01 / versões históricas','Fluxo 180M único','Elimina múltiplos motores concorrentes e indicadores desligados; caixa negativa não é escondida.'),('V23 / V25 / V27','Hotel + Operador','200 quartos próprios; designação exclusiva Operador Internacional; MICE e F&B separados.'),('Abertura mês 30 V45','Mês 31 no cenário reconstruído','Convenção de operação após conclusão de obra; editável e exige validação do calendário operacional.'),('HOTUSA tratada como saída','Reserva selecionável, pagamento opcional','Não presumir dívida certa. Reserva não reduz caixa por transferência interna.'),('Contingência obra','Dotação × realização','Pagamento e reserva não utilizada distinguidos; referência prudente 100% realizada.'),('Caixa Hotel MAX(0,...)','CFADS sem piso zero','Perdas e consumo de fundo maneio visíveis; colchão pré-abertura acrescido à reserva.'),('Juros e carência','Saldo usado e amortização contratual','Juros sobre abertura + desembolso do mês; álgebra evita circularidade de draw/juros.'),('Broker e comissões','Só na 1ª utilização efetiva','Success fee não incide sobre crédito nunca obtido; base sobre facilidade é hipótese a negociar.'),('DSCR e dívida residual','Serviço obrigatório / antecipações separados','Sem rácios sentinela 999; ausência de serviço deixa célula vazia.'),('Waterfall residencial','Residual KK integrado uma vez','Yield não garantido; OPEX, fees fixos, FF&E e waterfall separados; novo FM explícito.'),('504 / 810 escritura mês 29','Datas preservadas com pendência','Não validar contrato por alterar silenciosamente datas.'),('TIR e valor terminal','Vetores explícitos 10 e 15 anos','Valor terminal por EBITDA Hotel, dívida/caixa reconciliadas; histórico em vetor distinto.'),('Dashboards antigos','Cinco dashboards novos','Filtros comerciais, gráficos nativos e indicadores apontam ao mesmo motor.'),('Fórmulas de auditoria inválidas','Controlos refeitos e testados','V45 tem 5 IF(=...) e 12 expressões rejeitadas pelo parser; caches antigos não provam cálculo correto.')]
for r,row in enumerate(changes,5):
 for c,v in enumerate(row):text(s,r,c,v)

s=sh('Mapa V45',['Folha origem','Linhas','Colunas','Fórmulas','Dependências','Gráficos','Validações'],'Inventário integral da V45; CSV/JSON adicionais incluem lista de erros e mapa de dependências. Caches antigos são evidência, não recálculo certificado.')
s.set_column('A:A',36);s.set_column('E:E',100)
for r,row in enumerate(SOURCE['sheets'],5):
 for c,k in enumerate(['sheet','rows','columns','formulas','dependencies','charts','validations']):text(s,r,c,', '.join(row[k]) if isinstance(row[k],list) else row[k],'text' if c in [0,4] else 'int')
s.autofilter('A4:G74')
exec((ROOT/'src/v02_panels.py').read_text(),globals())
exec((ROOT/'src/v03_panels.py').read_text(),globals())
from scenario_seller_formulas import seller_updates
for sn,changes in seller_updates(SCMAP,PR).items():
 for addr,f in changes.items():S[sn].write_formula(addr,'='+f,F['eur'],'')

exec((ROOT/'src/v04_panels.py').read_text(),globals())

# Navigation links and consistent print settings.
visible=[name for name,sheet in S.items() if not name.startswith('Calc') and name!='Listas']
for sheet in S.values():
 if not sheet.name.startswith('Calc') and sheet.name not in ['LEIA-ME','EXEC KK01']:sheet.write_url('K1',"internal:'LEIA-ME'!A1",F['link'],'Início')
 if not sheet.name.startswith('Calc'):sheet.set_header('&LKK Hotelaria e Turismo, S.A.&R180 meses / 15 anos')
for i,name in enumerate(visible,27):read.write_url(i-1,0,f"internal:'{name}'!A1",F['link'],name)
# Notes on every blue input, including original physical data and chapter weights.
input_formats={id(v) for k,v in F.items() if k.startswith('input_')}
for ws in S.values():
 for row,cells in list(ws.table.items()):
  for c,cell in list(cells.items()):
   if hasattr(cell,'format') and id(cell.format) in input_formats and c not in ws.comments.get(row,{}):
    ws.write_comment(row,c,'Célula editável de cenário. Preservar fonte e suporte contratual; alterações propagam-se às fórmulas. Consultar Pressupostos, AUDITORIA e documentação técnica.',{'author':'KK-01'})
# Put the five dashboards at the front after the guide and inputs.
front=['EXEC KK01','LEIA-ME','Pressupostos','DASH Consolidado','DASH Real Estate','DASH Comercial','DASH Hotel MICE','DASH Acionistas','AUDITORIA']
wb.worksheets_objs.sort(key=lambda ws:(0,front.index(ws.name)) if ws.name in front else (1,list(S).index(ws.name)))
for i,ws in enumerate(wb.worksheets_objs):ws.index=i
S['EXEC KK01'].activate();S['EXEC KK01'].set_first_sheet()
wb.close()
(ROOT/'src/build_metadata.json').write_text(json.dumps({'param_rows':PR,'kpi_rows':KROW,'flow_columns':FI,'hotel_columns':HIDX,'res_columns':RIDX,'scenarios':SCEN,'source_sha256':SOURCE['source_sha256'],'v02':True},ensure_ascii=False,indent=2))
print('GRAVADO',OUT,OUT.stat().st_size,'bytes;',len(S),'folhas;',len(SCEN),'cenários completos')
