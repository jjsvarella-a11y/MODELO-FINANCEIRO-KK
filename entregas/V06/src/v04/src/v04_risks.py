# Stress calculations reuse the same hotel,residential,CAPEX,bank and waterfall functions.
RISK_DEFS=[('R00','Referência',{}),('R01','Sem novos contratosRE apósM24',{'sales_stop':24}),('R02','Sem novas subscrições apósM24',{'capital_stop':24}),('R03','Investidores +12M',{'capital_delay':12}),('R04','Vendas e escrituras +6M',{'commercial_delay':6}),('R05','Investidores+12M e vendas+6M',{'capital_delay':12,'commercial_delay':6}),('R06','CAPEX+10% e conclusão+6M',{'cost_overrun':.10,'construction_delay':6}),('R07','Hotel Conservador e fees maiores',{'hotel':'Conservador','fee_mode':'Proposta simulada'}),('R08','Escrituras+6M eADRresidencial−10%',{'deed_delay':6,'res_adr_multiplier':.9}),('R09','Apenas60%crédito mobilizável',{'bank_available':.6}),('R10','Apenas75%captação prevista',{'capital_fraction':.75}),('R11','Linha20M€',{'limit':20e6}),('R12','Linha25M€',{'limit':25e6}),('R13','Linha30M€',{'limit':30e6})]
RISK_META=[];RISK_RESULTS=[]
basepv=pv
for ri,(code,label,extra) in enumerate(RISK_DEFS):
 changes={};ov={};sd=extra.get('commercial_delay',0);dd=extra.get('deed_delay',0);fd=extra.get('capital_delay',0)
 for k,value in extra.items():
  if k in ['capital_delay','commercial_delay','construction_delay','deed_delay']:
   changes[k]=DEFAULT[k]+value;ov[k]=f'({rawpv(k)}+{value})'
  elif k in ['sales_stop','capital_stop']:
   changes[k]=min(DEFAULT[k],value);ov[k]=f'MIN({rawpv(k)},{value})'
  elif k in ['cost_overrun']:
   changes[k]=(1+DEFAULT[k])*(1+value)-1;ov[k]=f'((1+{rawpv(k)})*(1+{value})-1)'
  elif k in ['bank_available','capital_fraction']:
   changes[k]=DEFAULT[k]*value;ov[k]=f'({rawpv(k)}*{value})'
  elif k=='res_adr_multiplier':changes['res_adr']=DEFAULT['res_adr']*value;ov['res_adr']=f'({rawpv("res_adr")}*{value})'
  elif k=='limit':changes[k]=value;ov[k]=str(value)
  else:changes[k]=value
 if 'construction_delay' in extra:ov['end']='(30+'+ov['construction_delay']+')';ov['open']='(30+'+ov['construction_delay']+'+1+'+rawpv('opening_delay')+')'
 if 'cost_overrun' in extra:
  if DEFAULT['etc_amount'] is not None:changes['etc_amount']=DEFAULT['etc_amount']*(1+extra['cost_overrun'])
  ov['etc_amount']=f'IF(ISNUMBER({rawpv("etc_amount")}),{rawpv("etc_amount")}*(1+{extra["cost_overrun"]}),"")'
 rr=run(config(**changes));RISK_RESULTS.append(rr)
 if ri==0:RISK_META.append({'code':code,'label':label,'row':5,'flow':'Fluxo 180M','hotel':'Hotel','capex':'CAPEX','waterfall':'V04 Waterfall180','policy':'V04 Politica','params':extra});continue
 def pv(k):return ov.get(k,basepv(k))
 try:
  # Risk-specific pipeline retains independently documented fields in the main register.
  psn=f'V04 {code}Pipeline';ps=sh(psn,['Código'],'Pipeline de cenário,não execução.');ps.hide()
  for i,plan in enumerate(pipeline(rr['p'])):
   r=i+5
   for c,letter in [(0,'A'),(1,'B'),(5,'F'),(7,'H'),(8,'I')]:fw(ps,r,c,f"'V04 Pipeline'!{letter}{r}",None,'text' if c<2 else 'eur')
   fw(ps,r,9,f"'V04 Pipeline'!J{r}+{fd}",plan['pay_month'],'int')
   # Strip main cutoff/fraction before applying risk cutoff/fraction, never fabricate commitments.
   scale=f'IF({pv("equity_mode")}="Combinação tickets",1,MAX(0,{pv("equity")}-{pv("conversion")})/25000000)'
   fw(ps,r,19,f'IF(I{r}<={pv("capital_stop")},F{r}*H{r}*({scale})*{pv("received_ratio")}*{pv("capital_fraction")},0)',plan['planned_received']);fw(ps,r,22,f'IF(F{r}*({scale})>0,MIN(1,T{r}/(F{r}*({scale}))),0)',None,'pct')
  capital=f'V04 {code}Capital';cs=sh(capital,['Mês','Peso total','Peso ativo','Recebido'],'Capital efetivo do cenário, sem duplicar preço e transferência.');cs.hide()
  for m in range(1,181):
   r=m+4;text(cs,r,0,m,'int');fw(cs,r,3,f'SUMIF(\'{psn}\'!$J$5:$J$16,A{r},\'{psn}\'!$T$5:$T$16)',investor_outlay(rr['p'],m));fw(cs,r,2,f'IF({pv("equity")}>{pv("conversion")},D{r}/({pv("equity")}-{pv("conversion")}),0)',None,'pct');fw(cs,r,1,f'C{r}',None,'pct')
  # CaseCAPEX: immutable original chapter profile, scenario calendar and ETC.
  capsn='CAPEX'
  if 'construction_delay' in extra or 'cost_overrun' in extra:
   rawsn=f'V04 {code}CB';raws=sh(rawsn,['Mês','Base','Fiscalização','Contingência','Total'],'Perfil orçamento do cenário.');raws.hide();capsn=f'V04 {code}CAPEX';caps=sh(capsn,['Mês','Base','Fiscalização','Contingência','Total'],'Forecast sujeito aETC.');caps.hide()
   for m in range(1,181):
    r=m+4;text(raws,r,0,m,'int');text(caps,r,0,m,'int');src=f'IF(A{r}<{pv("pause_month")},A{r},IF(A{r}>={pv("pause_month")}+{pv("construction_delay")},A{r}-{pv("construction_delay")},0))';factor=1+extra.get('cost_overrun',0);bc=f'IF(AND(({src})>=1,({src})<=30),SUMPRODUCT(\'Obra\'!$D$5:$D$34,INDEX(\'Obra\'!$E$5:$AH$34,0,MAX(1,MIN(30,{src}))))*{factor},0)';hold=f'IF(AND({pv("construction_delay")}>0,A{r}>={pv("pause_month")},A{r}<{pv("pause_month")}+{pv("construction_delay")}),{pv("delay_cost")},0)'
    fw(raws,r,1,bc+'+'+hold,None);fw(raws,r,2,f'({bc})*{pv("supervision")}',None);fw(raws,r,3,f'({bc})*{pv("contingency")}*{pv("cont_used")}',None);fw(raws,r,4,f'SUM(B{r}:D{r})',None)
    future=f'SUMIFS(\'{rawsn}\'!$E$5:$E$184,\'{rawsn}\'!$A$5:$A$184,">"&{pv("status_month")})';adj=f'IF(AND({pv("etc_override")}=1,ISNUMBER({pv("etc_amount")}),A{r}>{pv("status_month")}),IF({future}>0,{pv("etc_amount")}/{future},0),1)'
    for c in range(1,5):fw(caps,r,c,f"'{rawsn}'!{col(c)}{r}*({adj})",rr['cap'][m-1][['','base','supervision','cont','total'][c]])
  resn=write_case_re(code,rr)
  # Operating blocks and the same financing motor.
  hsn=f'V04 {code}Hotel';hs=sh(hsn,HHEAD,'Hotel200+MICE do cenário.');hs.hide();hn='"'+extra['hotel']+'"' if 'hotel' in extra else None;fees=[vr('Condicoes Operador',j,1,True) for j in [5,6,7]] if 'fee_mode' in extra else None
  for i,x in enumerate(rr['hotel']):
   r=i+5;text(hs,r,0,i+1,'int');hf=hotel_formulas(r,hn,fees);cost='(2*'+pv('benefit_platinum')+'+3*'+pv('benefit_gold')+'+15*'+pv('benefit_silver')+')/12';hf['ebitda']+=f'-IF(A{r}>={pv("open")},{cost},0)'
   for k in HK[1:]:fw(hs,r,HIDX[k],hf[k],x[k],'pct' if k=='occ' else 'eur')
  asn=f'V04 {code}Resid';asheet=sh(asn,['Mês']+RK[1:],'Pool residencial elegível por escritura.');asheet.hide()
  for i,x in enumerate(rr['res']):
   r=i+5;text(asheet,r,0,i+1,'int');rf=res_formulas(r)
   for k in RK[1:]:fw(asheet,r,RIDX[k],rf[k],x[k],'eur')
  fsn=f'V04 {code}Fluxo';fs=sh(fsn,FH,'Mesmo motor bancário180M; condições do cenário.');fs.hide();write_flow(fs,rr,hotel_sn=hsn,capex_sn=capsn,re_sn=resn,residential_sn=asn,capital_sn=capital)
  pol=f'V04 {code}Politica';write_policy(pol,rr,fsn,hsn);wf=f'V04 {code}WF';write_waterfall(wf,pol,rr,fsn,0,hsn,capital_sn=capital,pipeline_sn=psn)
  RISK_META.append({'code':code,'label':label,'row':ri+5,'flow':fsn,'hotel':hsn,'capex':capsn,'waterfall':wf,'policy':pol,'params':extra,'capital':capital,'pipeline':psn})
 finally:pv=basepv

s=sh('V04 Riscos',['Cenário','Funding gap','Falta reservas','Caixa mínima','Pico dívida','Usado linha','Juros','Completion coverage projetadaM24','DSCR anualmín','Dividendos','TIR novos — política','MOIC novos','Saída equity','Estado funding','Metas convencionais','Método / confirmação'],'Sensibilidades combinadas ligadas aos mesmos motores. Coverageprojetada não ésecured; TIRpositiva não resolve défice ou violação de reservas.30M€éback-up estratégico, não recomendação de utilização integral.')
for item,rr in zip(RISK_META,RISK_RESULTS):
 r=item['row'];fs=item['flow'];wf=item['waterfall'];pol=item['policy'];hs=item['hotel'];caps=item['capex'];text(s,r,0,item['label']);Rv=lambda k,i:vr(fs,i,FI[k]);ranges=lambda k:f"'{fs}'!{col(FI[k])}5:{col(FI[k])}184"
 for c,f,fmt in [(1,'MAX('+ranges('gap')+')','eur'),(2,'MAX('+ranges('reservegap')+')','eur'),(3,'MIN('+ranges('cash')+')','eur'),(4,'MAX('+','.join(Rv('debt_open',j)+'+'+Rv('draw',j) for j in range(5,185))+')','eur'),(5,'MAX('+ranges('cumulative')+')','eur'),(6,'SUM('+ranges('interest')+')','eur'),(9,'SUM('+ranges('dividend')+')','eur'),(10,f'IFERROR((1+IRR(\'{wf}\'!V5:V184,0.005))^12-1,"TIR não definida")','pct'),(11,f'IF(SUM(\'{wf}\'!U5:U184)>0,SUM(\'{wf}\'!T5:T184)/SUM(\'{wf}\'!U5:U184),"")','mult'),(12,f"'{pol}'!B11",'eur')]:fw(s,r,c,f,None,fmt)
 # Forecast coverageuses case-specific dates and availability; actual evidence remains central.
 end_expr=f'{rr["p"]["open"]}+({pv("construction_delay")}-{DEFAULT["construction_delay"]})+({pv("opening_delay")}-{DEFAULT["opening_delay"]})';last=f'({end_expr}+4)';cut='24';futureEq=f'SUM(\'{fs}\'!C29:INDEX(\'{fs}\'!C5:C184,{end_expr}))';futureRE=f'SUM(\'{item.get("flow")}\'!D29:INDEX(\'{fs}\'!D5:D184,{end_expr}))-SUM(\'{fs}\'!E29:INDEX(\'{fs}\'!E5:E184,{end_expr}))-SUM(\'{fs}\'!F29:INDEX(\'{fs}\'!F5:F184,{end_expr}))-SUM(\'{fs}\'!G29:INDEX(\'{fs}\'!G5:G184,{end_expr}))';lim=str(item['params'].get('limit',DEFAULT['limit'])) if 'limit' in item['params'] else pv('limit');avail=f'{pv("bank_available")}*{item["params"].get("bank_available",1)}';bank=f'IF(AND({pv("bank_cp")}=1,24<{pv("draw_end")}),MAX(0,({lim})*({avail})-\'{fs}\'!AC28),0)';bank=projected_bank_formula(fs,caps,pv,end_expr,lim,avail);source=f"'{fs}'!AE28+({futureEq})+({futureRE})+({bank})"
 charges='+'.join(f'SUM(\'{fs}\'!{col(FI[k])}29:INDEX(\'{fs}\'!{col(FI[k])}5:{col(FI[k])}184,{end_expr}))' for k in ['interest','principal','openingfee','broker','mobil','loan_interest','loan_principal','commitment','hotpay']);needs=f'SUM(\'{caps}\'!E29:INDEX(\'{caps}\'!E5:E184,{end_expr}))+({charges})+MAX(0,-SUM(\'{fs}\'!L29:INDEX(\'{fs}\'!L5:L184,{end_expr}))-SUM(\'{fs}\'!M29:INDEX(\'{fs}\'!M5:M184,{end_expr})))+INDEX(\'{fs}\'!Q5:Q184,{end_expr})';fw(s,r,7,f'IF(({needs})>0,({source})/({needs}),"")',None,'mult')
 vals=[];cnt=[]
 for y in range(1,16):
  st=(y-1)*12+5;en=st+11;serv=f'SUM(\'{fs}\'!AH{st}:AH{en})';cond=f'AND({y*12}>={end_expr},{serv}>0.000001)';vals.append(f'IF({cond},SUM(\'{fs}\'!L{st}:L{en})/{serv},1E99)');cnt.append(f'IF({cond},1,0)')
 fw(s,r,8,f'IF(SUM({",".join(cnt)})>0,MIN({",".join(vals)}),"")',None,'mult');fw(s,r,13,f'IF(OR(B{r}>0.01,C{r}>0.01,H{r}<1),"Alerta — funding/reservas","Projetado coberto; secured pendente")',None,'text');fw(s,r,14,f'IF(\'{pol}\'!B30,"Pesos constantes exatos","ConsultarLP / desvios preferência")',None,'text');text(s,r,15,'Hipótese de esforço; parâmetros e datas não aprovados')
s.autofilter(3,0,len(RISK_META)+3,15);alert(s,f'N5:N{len(RISK_META)+4}')
# Financial ImpactView: explicit changes in money, not invented performance.
s=sh('V04 Impactos',['Cenário / marco','Δfunding gap','Δcaixa mínima','Δpico dívida','Δjuros','Δdividendos','ΔTIR novos','Ação / decisão'],'FINANCIAL IMPACT VIEW: diferenças contraR00; não são desvios realizados. Reprogramar,capitalizar,negociarcrédito ou corrigir custos conforme diagnóstico.')
for r,item in enumerate(RISK_META[1:],5):
 sr=item['row'];text(s,r,0,item['label'])
 for c,sc in [(1,1),(2,3),(3,4),(4,6),(5,9),(6,10)]:fw(s,r,c,f'\'V04 Riscos\'!{col(sc)}{sr}-\'V04 Riscos\'!{col(sc)}5',None,'pct' if c==6 else 'eur')
 s.write_url(r-1,7,"internal:'V04 Decisoes Board'!A5",F['link'],'Rever decisão, documentos e plano de mitigação')
