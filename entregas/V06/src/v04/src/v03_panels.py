"""V03 additive panels. Never applied to V02 source or workbook."""
s=sh('Capital V03',['Métrica','Cenário modelado','Estado / critério'],'Dados modelados não são comprovativos de realização. Montantes documentados em branco = indisponíveis; não presumir equity realizado.')
s.set_column('A:A',46);s.set_column('B:B',26);s.set_column('C:C',100)
for r,label,f,fmt in [(5,'Cenário selecionado',rawpv('capital_scenario'),'text'),(6,'Preço contratado / subscrito',f'IF({rawpv("equity_mode")}="Combinação tickets",SUM(\'Blocos V03\'!E5:E8),IF(B5="KK-E15",15000000,IF(OR(B5="KK-E20-A",B5="KK-E20-B"),20000000,25000000)))','eur'),(7,'Participação subscrita',f'IF({rawpv("equity_mode")}="Combinação tickets",SUM(\'Blocos V03\'!F5:F8),IF(B5="KK-E20-B",50%,IF(B5="KK-E25",50%,40%)))','pct'),(8,'JSCV após aquisição modelada',f'IF({rawpv("equity_mode")}="Combinação tickets",51%-11%*{rawpv("n_platinum")},IF(OR(B5="KK-E25",B5="KK-E20-B"),40%,51%))','pct'),(9,'JALS após aquisição modelada','1-B7-B8','pct'),(10,'Recebido investidores (cenário)',f'B6*{rawpv("received_ratio")}','eur'),(11,'Retido vendedores (cenário)',f'B10*(1-{rawpv("treasury_ratio")})','eur'),(12,'Transferido KK (cenário)',f'B10*{rawpv("treasury_ratio")}','eur'),(13,'Diferença reconciliação','B10-B11-B12','eur'),(14,'Natureza da transferência',rawpv('funding_kind'),'text'),(15,'Obrigação de reembolso',f'IF(B14="Empréstimo","Sim — ver Fluxo 180M AU:AX","Não, apenas hipótese não formalizada")','text'),(16,'Soma tranches',f'{rawpv("fund_part1")}+{rawpv("fund_part2")}+{rawpv("fund_part3")}','pct'),(17,'Estado jurídico',f'IF({rawpv("legal_confirmed")}=1,"Formalização assinalada — consultar documento","Pendente de Validação Externa")','text'),(18,'Unpaid / contratado','B6-B10','eur')]:
 text(s,r,0,label);fw(s,r,1,f,None,fmt)
text(s,20,0,'DADOS REALIZADOS / DOCUMENTADOS');text(s,21,0,'Compromisso assinado (€)');text(s,22,0,'Recebido investidores (€)');text(s,23,0,'Retido pelos vendedores (€)');text(s,24,0,'Transferido efetivamente KK (€)')
for r in range(21,25):s.write_blank(r-1,1,None,F['input_eur']);text(s,r,2,'Indisponível até entrega de comprovativo; não alimenta automaticamente o cenário.')
fw(s,25,1,'IF(COUNT(B22:B24)=3,B22-B23-B24,"Indisponível")',None,'text');text(s,25,0,'Reconciliação dos valores realizados')
s.merge_range('A28:C30','Ações subscritas não equivalem a ações liquidadas. Com received_ratio < 100%, a participação é expectativa contratada e o saldo por pagar fica visível. Não usar como prova de propriedade ou de equity bancário realizado.',F['warn'])
s.merge_range('A32:C34','Origem das ações KK-E25: 11.000 da JSCV e 39.000 da JALS. Hipótese de alocação: Platinum recebe 11.000 JSCV + 9.000 JALS; restantes blocos provenientes da JALS. Validar transmissão e titularidade. KK-E20-B é apenas comparação hipotética.',F['warn'])

s=sh('Blocos V03',['Categoria','Ticket unitário','Quota por bloco','Subscritos cenário','Capital contratado','Quota subscrita','Ações subscritas','Disponíveis iniciais','Por subscrever','Ações JSCV cedidas','Ações JALS cedidas'],'Cenário combinação: editar quantidades em Pressupostos e selecionar Combinação tickets. Apenas blocos subscritos recebem participação económica no modelo.')
CATS=[('Platinum',10e6,.20,'n_platinum',1,11000),('Gold',5e6,.10,'n_gold',1,0),('Silver',1e6,.02,'n_silver',10,0)]
for r,(label,ticket,stake,key,qty,jscv) in enumerate(CATS,5):
 for c,v in [(0,label),(1,ticket),(2,stake),(7,qty)]:text(s,r,c,v,'text' if c==0 else 'pct' if c==2 else 'eur' if c==1 else 'int')
 fw(s,r,3,rawpv(key),DEFAULT[key],'int')
 for c,f,fmt in [(4,f'B{r}*D{r}','eur'),(5,f'C{r}*D{r}','pct'),(6,f'ROUND(F{r}*{pv("shares_total")},0)','int'),(8,f'H{r}-D{r}','int'),(9,f'D{r}*{jscv}','int'),(10,f'G{r}-J{r}','int')]:fw(s,r,c,f,None,fmt)
for c in [3,4,5,6,7,8,9,10]:fw(s,10,c,f'SUM({col(c)}5:{col(c)}8)',None,'total_pct' if c==5 else 'total_eur' if c==4 else 'total_int')
s.autofilter('A4:K8');text(s,13,0,'Limites');fw(s,13,1,'IF(AND(MIN(D5:D8)>=0,MIN(I5:I8)>=0,SUMPRODUCT(D5:D8-INT(D5:D8),D5:D8-INT(D5:D8))=0),"Conforme","Alerta")',None,'text')

s=sh('Matriz Societaria V03',['Acionista','Platinum','Gold','Silver','Blocos totais','Quota alvo','Ações alvo','Quota ativa','Ações ativas'],'Estrutura alvo E25; direitos económicos separados de governação. Ações ainda não subscritas permanecem com os vendedores.')
for r,label,counts in [(5,'JSCV',[2,0,0]),(6,'JALS',[0,0,5]),(7,'Novos investidores',[1,1,10])]:
 text(s,r,0,label)
 for c,v in enumerate(counts,1):text(s,r,c,v,'int')
 fw(s,r,4,f'SUM(B{r}:D{r})',None,'int');fw(s,r,5,f'B{r}*20%+C{r}*10%+D{r}*2%',None,'pct');fw(s,r,6,f'ROUND(F{r}*{pv("shares_total")},0)',None,'int')
 fw(s,r,7,pv('legacy_jscv') if r==5 else "'Capital V03'!B9" if r==6 else "'Capital V03'!B7",None,'pct');fw(s,r,8,f'ROUND(H{r}*{pv("shares_total")},0)',None,'int')
for c in range(1,9):fw(s,9,c,f'SUM({col(c)}5:{col(c)}7)',None,'total_pct' if c in [5,7] else 'total_int')
for r,label,f in [(12,'Quotas alvo = 100%','ABS(F9-1)<0.0000001'),(13,'Ações alvo = 100.000','G9=100000'),(14,'Blocos totais = 19','SUM(B9:D9)=19'),(15,'Blocos investidores = 12','SUM(B7:D7)=12'),(16,'Tickets investidores = 25 M€',"SUMPRODUCT('Blocos V03'!B5:B8,'Blocos V03'!H5:H8)=25000000"),(17,'Quotas ativas = 100%','ABS(H9-1)<0.0000001'),(18,'Ações ativas inteiras = 100.000','I9=100000')]:text(s,r,0,label);fw(s,r,1,f'IF({f},"Conforme","Alerta")',None,'text')

s=S['Ronda Investidores'];fw(s,20,1,"'Capital V03'!B6",25e6);fw(s,22,1,"'Capital V03'!B12",25e6);fw(s,23,1,"'Capital V03'!B11",0);fw(s,24,1,"'Capital V03'!B7",.50,'pct');fw(s,27,1,"'Blocos V03'!B13",None,'text')
for r in range(5,9):
 for c,letter in [(1,'B'),(2,'C'),(3,'D'),(4,'E'),(5,'F'),(6,'G')]:fw(s,r,c,f"'Blocos V03'!{letter}{r}",None,'pct' if c in [2,5] else 'int' if c in [3,6] else 'eur')
s.merge_range('A41:G43','V03: 25 M€ / 50%, com 12 blocos investidores. Transferência à KK condicional; ver Capital V03. Os quadros de benefícios V02 abaixo são históricos e substituídos pela matriz Beneficios V03.',F['warn'])
# Correct economic bridge and previous fixed audit constraints.
s=S['Societario'];fw(s,5,1,'51%',.51,'pct');fw(s,6,1,'49%',.49,'pct');fw(s,5,2,pv('legacy_jscv'),.4,'pct');fw(s,6,2,"'Capital V03'!B9",.11,'pct')
text(s,5,3,'KK-E25: cede 11.000 ações, conserva 2 Platinum = 40%. Direitos estratégicos por formalizar.');text(s,6,3,'KK-E25: cede 39.000 ações, conserva 10%.');text(s,7,3,'KK-E25: 50.000 ações; 12 blocos. Recebimento e transferência são operações separadas.')
for r,f in [(19,"'Capital V03'!B8*100000"),(20,"'Capital V03'!B9*100000"),(21,"'Capital V03'!B7*100000")]:fw(s,r,1,f,None,'int')
for c in [1,2]:
 fw(S['Retornos'],27,c,pv('legacy_jscv'),None,'pct');fw(S['Retornos'],28,c,"'Capital V03'!B9",None,'pct')
# Finance blocks already contain loan balance; deduct from terminal equity in all scenario exits.
for item in SCMAP:
 r=item['row'];sc=item['scenario'];stake=str(sc.get('equity_stake',.4)) if sc.get('equity') is not None else pv('equity_stake')
 fw(resheet,r,4,f'B{r}/D{r}*({stake})',None,'pct')
 fw(resheet,r,2,f"SUM('{item['flow']}'!C5:C184)",None)
 for sheet in [resheet,vec]:
  for rr,cells in list(sheet.table.items()):
   for cc,cell in list(cells.items()):
    if cell.__class__.__name__=='Formula' and f"'{item['flow']}'!AB" in cell.formula:
     import re
     formula=re.sub(r"-('"+re.escape(item['flow'])+r"'!AB(\d+))",lambda m:'-'+m.group(1)+"-'"+item['flow']+"'!AX"+m.group(2),cell.formula)
     fw(sheet,rr+1,cc,formula.lstrip('='),None)
# Hotel benefits cost: deducted after incentive, before tax, preserving GOP incentive base.
benefit='(3*'+pv('benefit_platinum')+'+'+pv('benefit_gold')+'+15*'+pv('benefit_silver')+')/12'
for name,hs in S.items():
 if name=='Hotel' or name.startswith('CalcH'):
  for r in range(5,185):
   old=hs.table[r-1][HIDX['ebitda']].formula
   fw(hs,r,HIDX['ebitda'],old.lstrip('=')+f'-IF(A{r}>={pv("open")},{benefit},0)',None)

s=sh('Governacao V03',['Matéria','Proposta para estudo','Proteção restantes acionistas','Estado jurídico'],'40% económicos e dois Platinum não conferem controlo absoluto. Não existem votos múltiplos ou vetos aprovados neste modelo.')
items=[('Direitos económicos','Proporcionais às ações adquiridas, sem preferência aprovada','Igualdade de tratamento; regras transparentes de dividendos'),('Votos','Uma ação / um voto como hipótese de análise; confirmar Estatutos','Sem votos múltiplos presumidos'),('Nomeações','Estudar representação estratégica da JSCV / Platinum no conselho','Representação investidores e minoritários; independentes'),('Matérias reservadas','Definir orçamento, endividamento, alienações e alteração do objeto','Quóruns e maiorias sujeitos a parecer jurídico; evitar bloqueios'),('Direitos especiais Platinum','Proposta de concentração de determinadas funções estratégicas','Não atribuir veto absoluto; cessação e transmissibilidade por definir'),('Conflitos de interesses','Operações relacionadas com aprovação independente','Abstenção, divulgação e avaliação independente'),('Informação','Relatório mensal e contas anuais auditadas','Acesso proporcional e confidencialidade'),('Saída / liquidez','Estudar tag-along, drag-along e preferência','Sem recompra obrigatória ou liquidez garantida')]
for r,row in enumerate(items,5):
 for c,v in enumerate((*row,'Pendente de Validação Externa')):text(s,r,c,v,'input_text' if c==1 else 'text')
s.set_column('A:A',30);s.set_column('B:C',70);s.set_column('D:D',40)

s=sh('Beneficios V03',['Categoria','Custo anual por bloco','N.º total alvo','Custo anual KK','Benefícios comerciais em estudo','Direitos económicos preferenciais'],'Todos os custos são hipóteses editáveis; zero não afirma gratuitidade. Custo de todos os blocos promotores + investidores deduzido do EBITDA/CFADS; sem receita artificial para investidor.')
for r,(label,ticket,stake,key,qty,jscv) in enumerate(CATS,5):
 text(s,r,0,label);fw(s,r,1,pv('benefit_'+key[2:]),0);text(s,r,2,[3,1,15][r-5],'int');fw(s,r,3,f'B{r}*C{r}',0);text(s,r,4,'Alojamento preferencial, MICE, instalações, fidelização — proposta');text(s,r,5,'Não aprovados; nenhuma yield, recompra ou preferência nos dividendos')
text(s,11,0,'Governança / institucionais');text(s,11,1,'Informação reforçada, reuniões estratégicas, futuras oportunidades — ver Governacao V03')
text(s,12,0,'Custo anual incremental KK');fw(s,12,1,'SUM(D5:D8)',0)
s.merge_range('A15:F17','Simular custo económico líquido (despesa + margem perdida) nos pressupostos por categoria. Não adicionar valor subjetivo de hospitalidade à TIR financeira. Preferências de dividendos ou liquidez obrigatória requerem decisão expressa e modelação adicional antes de contratar.',F['warn'])

s=sh('Comparativo Capital V03',['Cenário','Preço contratado','Mobilizado KK','Quota novos','Pico dívida','Usado linha','Juros','Comissões','Caixa mínima','Caixa M30','Dívida M30','DSCR parcial','Dividendos','TIR 15A','MOIC 15A','VPL 15A','Saída equity','Retorno promotores','Estado'],'Quatro cenários recalculados no mesmo motor 180M. Não confundir maior equity com maior TIR; E20-B hipotético.')
CAPMAP={}
for r,label in enumerate(['KK-E15','KK-E20-A','KK-E20-B','KK-E25'],5):
 item=next(x for x in SCMAP if x['scenario']['label']==label);sr=item['row'];CAPMAP[label]=item
 text(s,r,0,label)
 for c,sourcec in enumerate([1,2,4,6,5,7,8,10,13,15,17,19,20,21,22,25,23,26],1):fw(s,r,c,vr('Resultados Cenarios',sr,sourcec),None,'pct' if c in [3,13] else 'mult' if c in [11,14] else 'text' if c==18 else 'eur')
chart(s,'A12','column','Capital / dívida / juros',['Comparativo Capital V03',4,0,7,0],[(label,['Comparativo Capital V03',4,c,7,c]) for label,c in [('Capital KK',2),('Pico dívida',4),('Usado linha',5)]])
# Returns by category, per single block on active operating/capital scenario.
s=sh('Retorno Categorias V03',['Categoria','Ticket','Quota','Preço por ponto percentual','Dividendos 10A','Saída 10A','MOIC 10A','TIR 10A','VPL 10A','Dividendos 15A','Saída 15A','MOIC 15A','TIR 15A','VPL 15A','Payback dividendos','Condição'],'Três categorias ativas; preço uniforme por ponto percentual. Simulação por bloco. Só KK-E25 / Combinação tickets corresponde à nova oferta comercial.')
v=sh('Vet Categorias V03',['Mês'],'Fluxos mensais por categoria; entradas investidor nas tranches e dividendos/saída por ações.');v.hide()
for m in range(181):text(v,m+5,0,m,'int')
for r,(label,ticket,stake,key,qty,jscv) in enumerate(CATS,5):
 text(s,r,0,label);text(s,r,1,ticket,'eur');text(s,r,2,stake,'pct');fw(s,r,3,f'B{r}/(C{r}*100)',None)
 vc=(r-5)*2+1
 for h,j in [(120,0),(180,1)]:
  for m in range(181):
   mr=m+5;fr=m+4;weight='+'.join(f'IF(A{mr}={pv("fund_month"+str(k))},{pv("fund_part"+str(k))},0)' for k in [1,2,3]);rc='B' if h==120 else 'C'
   f='0' if m==0 or m>h else f"-'Retorno Categorias V03'!$B${r}*{rawpv('received_ratio')}*({weight})+'Fluxo 180M'!AD{fr}*'Retorno Categorias V03'!$C${r}"
   if m==h:f+=f"+'Retornos'!{rc}11*'Retorno Categorias V03'!$C${r}"
   fw(v,mr,vc+j,f,None)
  base=4 if h==120 else 9;rc='B' if h==120 else 'C';vl=col(vc+j)
  fw(s,r,base,f"'Retornos'!{rc}12*C{r}",None);fw(s,r,base+1,f"'Retornos'!{rc}11*C{r}",None)
  fw(s,r,base+2,f'IF(B{r}*{rawpv("received_ratio")}>0,({col(base)}{r}+{col(base+1)}{r})/(B{r}*{rawpv("received_ratio")}),"")',None,'mult')
  fw(s,r,base+3,f'IFERROR((1+IRR(\'Vet Categorias V03\'!{vl}5:{vl}{h+5},0.005))^12-1,"TIR não definida")',None,'pct')
  fw(s,r,base+4,f'NPV((1+{pv("equity_discount")})^(1/12)-1,\'Vet Categorias V03\'!{vl}6:{vl}{h+5})',None)
 # Payback requires cumulative paid investment vs dividends, excludes hypothetical exit.
 terms=[]
 for y in range(1,16):
  c=16+y;fw(s,r,c,f"'Anual'!V{y+4}*C{r}",None);terms.append(f'SUM(R{r}:{col(c)}{r})')
  if r==5:text(s,4,c,'Dividendos ano '+str(y))
 for c in range(33,213):
  m=c-32;fw(s,r,c,f'IF(SUM(\'Fluxo 180M\'!$AD$5:AD{m+4})*C{r}>=B{r}*{rawpv("received_ratio")},{m},1E9)',None,'int')
 fw(s,r,14,f'IF(MIN(AH{r}:HE{r})<1E9,MIN(AH{r}:HE{r}),"Não recuperado")',None,'text')
 fw(s,r,15,f'IF(OR(MAX(\'Fluxo 180M\'!AF5:AG184)>0.01,{rawpv("received_ratio")}<1),"Alerta — retorno condicionado","Cenário, não garantia")',None,'text')
s.set_column('R:AF',18);s.set_column('AH:HE',None,None,{'hidden':True});s.autofilter('A4:AF8')

# Update executive caption, decision, extra indicators and actual chart.
s=S['EXEC KK01'];fw(s,3,0,'"V03 • "&\'Capital V03\'!B5&" • "&TEXT(\'Capital V03\'!B12/1000000,"0.0")&" M€ mobilizados / "&TEXT(\'Capital V03\'!B7,"0%")&" • Hipóteses sujeitas a formalização"',None,'text')
s.write('A37','V03: comparar segurança financeira e retorno por categoria. 40% económicos e dois Platinum não garantem controlo jurídico.',F['warn'])
s.write_url('A40',"internal:'Comparativo Capital V03'!A1",F['link'],'↔ Quatro cenários capital V03')
for r,label,f,fmt in [(43,'Capacidade bancária remanescente',kref('unused'),'eur'),(44,'JSCV / dois Platinum',"'Capital V03'!B8",'pct'),(45,'JALS',"'Capital V03'!B9",'pct'),(46,'CFADS Hotel ano 4',"'Anual'!J8",'eur'),(47,'Retorno promotores 15A',"'Retornos'!C21",'eur')]:metric(s,r,label,f,fmt=fmt)
chart(s,'A50','column','Comparação de capital e dívida',['Comparativo Capital V03',4,0,7,0],[(label,['Comparativo Capital V03',4,c,7,c]) for label,c in [('Capital KK',2),('Pico dívida',4)]])
s.print_area('A1:P68')
for name in ['Capital V03','Blocos V03','Matriz Societaria V03','Governacao V03','Beneficios V03','Comparativo Capital V03','Retorno Categorias V03']:S[name].write_url('L1',"internal:'EXEC KK01'!A1",F['link'],'← Executivo')
S['DASH Acionistas'].write_url('A26',"internal:'Retorno Categorias V03'!A1",F['link'],'↗ TIR / MOIC por categoria');S['DASH Riscos'].write_url('A20',"internal:'Governacao V03'!A1",F['link'],'↗ Direitos estratégicos pendentes')
# V02 historical text is superseded, not erased from preserved V02. Clear misleading fixed assertions.
a=S['AUDITORIA']
for rr,cells in list(a.table.items()):
 cell=cells.get(1)
 if cell and cell.__class__.__name__=='Formula':
  f=cell.formula
  if 'C5-51%' in f:fw(a,rr+1,1,"'Societario'!C5-'Capital V03'!B8",0)
 status=cells.get(2)
 if status and status.__class__.__name__=='Formula':
  f=status.formula.replace('<=40%','<=50%').replace('<=20000000','<=25000000');fw(a,rr+1,2,f.lstrip('='),None,'text')
# Add explicit V03 audit rows.
start=max(a.table)+3
for r,label,test,external in [(start,'Ações e quotas reconciliadas',"AND('Matriz Societaria V03'!H9=1,'Matriz Societaria V03'!I9=100000)",False),(start+1,'Tickets e blocos disponíveis',"'Blocos V03'!B13=\"Conforme\"",False),(start+2,'Fundos recebido = retido + KK',"ABS('Capital V03'!B13)<0.01",False),(start+3,'Tranches somam 100%',"ABS('Capital V03'!B16-1)<0.0000001",False),(start+4,'Controlo Platinum aprovado','FALSE',True),(start+5,'Transmissão e fiscalidade aprovadas',f'{pv("legal_confirmed")}=1',True)]:
 text(a,r,0,label);fw(a,r,1,f'IF({test},0,1)',None);fw(a,r,2,f'IF({test},"Conforme","'+('Pendente de Validação Externa' if external else 'Alerta')+'")',None,'text')
(ROOT/'src/v03_metadata.json').write_text(json.dumps({'capital_scenarios':CAPMAP,'new_sheets':['Capital V03','Blocos V03','Matriz Societaria V03','Governacao V03','Beneficios V03','Comparativo Capital V03','Retorno Categorias V03']},ensure_ascii=False,indent=2))
# Mobilisation dates also determine investor disbursements; seller loan receipts are promoter flows.
s=S['Fluxos Retorno']
for r in range(5,186):
 m=r-5;fr=m+4
 for c,h in [(1,120),(2,180)]:
  rc='B' if h==120 else 'C';weight='+'.join(f'IF(A{r}={pv("fund_month"+str(j))},{pv("fund_part"+str(j))},0)' for j in [1,2,3])
  f='0' if m==0 or m>h else f'-({pv("equity")}-{pv("conversion")})*{rawpv("received_ratio")}*({weight})+\'Fluxo 180M\'!AD{fr}*\'Retornos\'!{rc}13'
  if m==h:f+=f"+'Retornos'!{rc}11*'Retornos'!{rc}13"
  fw(s,r,c,f,None)
 for c,h in [(3,120),(4,180)]:
  rc='B' if h==120 else 'C'
  if 1<=m<=h:
   old=s.table[r-1][c].formula.lstrip('=');old+=f"+'Fluxo 180M'!AV{fr}+'Fluxo 180M'!AW{fr}"
   if m==h:old+=f"+'Fluxo 180M'!AX{fr}"
   fw(s,r,c,old,None)
for c,h in [(1,120),(2,180)]:
 rc=col(c);end=h+4
 old=S['Retornos'].table[20][c].formula.lstrip('=');fw(S['Retornos'],21,c,old+f"+SUM('Fluxo 180M'!AV5:AW{end})+'Fluxo 180M'!AX{end}",None)
 fw(S['Retornos'],18,c,f'IF(({pv("equity")}-{pv("conversion")})*{rawpv("received_ratio")}>0,{rc}16/(({pv("equity")}-{pv("conversion")})*{rawpv("received_ratio")}),"")',None,'mult')
# Suppress misleading V02-specific notes and assertions in the active V03.
text(S['LEIA-ME'],8,1,'V03: KK-E25 25 M€ / 50%; JSCV 40%, JALS 10%. Selecionar cenário em Pressupostos. Realização efetiva por documentar.')
text(S['Retornos'],27,3,'JSCV/JALS ligadas ao cenário V03; 40%/10% na referência E25. Ver Matriz Societaria V03.');text(S['Retornos'],28,3,'Retorno promotores inclui reembolso/juros de eventual empréstimo dos vendedores; alocação individual do crédito pendente.')
text(S['AUDITORIA'],start-9,3,'Controlo herdado V02; limites atualizados V03. Consultar testes novos.') if False else None
s=S['Ronda Investidores'];s.write('A29','V03: ver Capital V03. Não reembolsável é hipótese; Empréstimo gera dívida, juros, amortizações e bloqueia dividendos até liquidação.',F['warn']);text(s,34,1,'V03: JSCV = 2 Platinum');text(s,35,1,'V03: JALS = 5 Silver');s.write('A37','Direitos e benefícios V03 em Governacao V03 / Beneficios V03; sujeitos a validação. Estrutura V02 preservada no ficheiro anterior.',F['warn'])
for name in ['Blocos V03','Capital V03','Comparativo Capital V03','Retorno Categorias V03','Governacao V03','Beneficios V03']:
 S['DASH Acionistas'].write_url('D'+str(28+list(['Blocos V03','Capital V03','Comparativo Capital V03','Retorno Categorias V03','Governacao V03','Beneficios V03']).index(name)*2),"internal:'"+name+"'!A1",F['link'],name)
# Full note on source-of-shares allocation and superseded legacy generic investor table.
S['Investidores'].write('A2','Tabela linear histórica de sensibilidade — não representa os preços das categorias V03. Usar Retorno Categorias V03 para a oferta de blocos.',F['warn'])
# Scenario return vectors follow the same staged payments and promoter creditor cash flows.
for si,item in enumerate(SCMAP):
 sr=item['row'];fsn=item['flow'];vc=si*4+1
 for m in range(1,181):
  r=m+5;fr=m+4;weight='+'.join(f'IF(A{r}={pv("fund_month"+str(j))},{pv("fund_part"+str(j))},0)' for j in [1,2,3])
  for j,h in [(0,120),(1,180)]:
   if m<=h:
    old=vec.table[r-1][vc+j].formula.lstrip('=');old=old.replace(f'IF(A{r}=1,\'Resultados Cenarios\'!$B${sr},0)',f'\'Resultados Cenarios\'!$B${sr}*{rawpv("received_ratio")}*({weight})');fw(vec,r,vc+j,old,None)
  for j,h in [(2,120),(3,180)]:
   if m<=h:
    old=vec.table[r-1][vc+j].formula.lstrip('=')+f"+'{fsn}'!AV{fr}+'{fsn}'!AW{fr}"
    if m==h:old+=f"+'{fsn}'!AX{fr}"
    fw(vec,r,vc+j,old,None)
 for c,h in [(21,180),(27,120)]:
  old=resheet.table[sr-1][c].formula.lstrip('=');old=old.replace(f'/B{sr}',f'/(B{sr}*{rawpv("received_ratio")})');fw(resheet,sr,c,old,None,'mult')
 old=resheet.table[sr-1][23].formula.lstrip('=');fw(resheet,sr,23,old+f"+SUM('{fsn}'!AV5:AW184)+'{fsn}'!AX184",None)
s=S['Anual']
for c,label in [(34,'Juros vendedores'),(35,'Capital vendedores'),(36,'Saldo empréstimo vendedores')]:text(s,4,c,label,'head')
for y in range(1,16):
 r=y+4;st=(y-1)*12+5;en=st+11
 for c,f in [(34,f"SUM('Fluxo 180M'!AV{st}:AV{en})"),(35,f"SUM('Fluxo 180M'!AW{st}:AW{en})"),(36,f"'Fluxo 180M'!AX{en}")]:fw(s,r,c,f,None)
# Seller receipts are only amounts actually received and retained, at the mobilisation dates.
s=S['Fluxos Retorno']
for m in range(181):
 r=m+5;fr=m+4;weight='+'.join(f'IF(A{r}={pv("fund_month"+str(j))},{pv("fund_part"+str(j))},0)' for j in [1,2,3])
 for c,h in [(3,120),(4,180)]:
  rc='B' if h==120 else 'C'
  f=f'-{pv("historic")}' if m==0 else '0' if m>h else f"'Fluxo 180M'!AD{fr}*'Retornos'!{rc}20+'Capital V03'!B11*({weight})+'Fluxo 180M'!AV{fr}+'Fluxo 180M'!AW{fr}"
  if m==h:f+=f"+'Retornos'!{rc}11*'Retornos'!{rc}20+'Fluxo 180M'!AX{fr}"
  fw(s,r,c,f,None)
for c,h in [(1,120),(2,180)]:
 rc=col(c);end=h+4;fw(S['Retornos'],21,c,f"({rc}11+{rc}12)*{rc}20+'Capital V03'!B11+SUM('Fluxo 180M'!AV5:AW{end})+'Fluxo 180M'!AX{end}",None)
# Source allocation applied to retained proceeds; the legal instrument remains pending.
source_ratio="IF('Capital V03'!B7>0,(51%-'Capital V03'!B8)/'Capital V03'!B7,0)"
for r,ratio in [(5,source_ratio),(6,'(1-'+source_ratio+')')]:fw(S['Retorno Socios'],r,4,f"('Capital V03'!B11+SUM('Fluxo 180M'!AV5:AW184)+'Fluxo 180M'!AX184)*({ratio})",None)

# Three-category decision confirmed by user; legacy Silver Plus controls locked to zero in Excel.
for key in ['n_silver_plus','benefit_silver_plus']:
 row=PR[key];text(S['Pressupostos'],row,0,'Histórico desativado — '+key);S['Pressupostos'].write_formula(row-1,1,'=0',F['num'],0);text(S['Pressupostos'],row,3,'Silver Plus retirado da oferta V03 por instrução expressa. Preservado apenas como rasto histórico; sem efeito económico.')
# Remove category-four rows in legacy active round and category matrices.
for r in [8]:
 for c in range(7):S['Ronda Investidores'].write_blank(r-1,c,None,F['text'])
for c in range(17,32):text(S['Retorno Categorias V03'],4,c,'Dividendos ano '+str(c-16),'head')
for c in [1,2]:
 rc=col(c);fw(S['Retornos'],17,c,f'{rc}16-({pv("equity")}-{pv("conversion")})*{rawpv("received_ratio")}',None)

S['Retorno Socios'].write('A2','V03: JSCV40%, JALS10%, novos50% na referência. Custos históricos individuais desconhecidos; não inferir TIR.',F['subtitle'])

for key,r in [('equity',6),('equity_price',6),('equity_stake',7),('legacy_jscv',8)]:fw(S['Pressupostos'],PR[key],1,"'Capital V03'!B"+str(r),None,'pct' if key in ['equity_stake','legacy_jscv'] else 'eur')
