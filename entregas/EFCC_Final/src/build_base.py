from core import *
import csv,shutil,hashlib,os
OUT=Path(__file__).resolve().parents[1];WORK=Path(os.environ.get('KK01_WORKDIR','/workspace/kk01_efcc'));WORK.mkdir(exist_ok=True)
def build():
 wb=xlsxwriter.Workbook(str(WORK/'V06_modulos.xlsx'));wb.set_calc_mode('auto');wb.set_properties({'title':'KK-01 V06 Executive Master — Candidate','company':'JSCV','comments':'Native Excel 365 certification pending. V04 engine retained as legacy reference.'})
 palette={'charcoal':'252932','red':'B31F2B','text':'202833','grey':'69717C','light':'F5F6F8'}
 fm={}
 for key,extra in {'num':{'num_format':'#,##0.00;[Red](#,##0.00);–'},'int':{'num_format':'0'},'pct':{'num_format':'0.00%'},'text':{},'head':{'bold':True,'font_color':'FFFFFF','bg_color':'252932','text_wrap':True},'input':{'font_color':'1565A0','bg_color':'EAF2FA','num_format':'#,##0.00'},'title':{'font_size':23,'bold':True},'sub':{'font_size':10,'font_color':'69717C'},'section':{'font_size':14,'bold':True,'bottom':1,'bottom_color':'B31F2B'},'card':{'font_size':23,'bold':True,'bg_color':'F5F6F8','num_format':'#,##0;[Red](#,##0);–'},'cardpct':{'font_size':23,'bold':True,'bg_color':'F5F6F8','num_format':'0.00%'},'cardx':{'font_size':23,'bold':True,'bg_color':'F5F6F8','num_format':'0.00"×"'},'label':{'font_size':10,'font_color':'69717C','bg_color':'F5F6F8','text_wrap':True},'note':{'font_size':10,'text_wrap':True,'valign':'top'},'side':{'bg_color':'252932','font_color':'FFFFFF','font_size':10,'text_wrap':True,'valign':'vcenter'},'active':{'bg_color':'B31F2B','font_color':'FFFFFF','font_size':10,'bold':True,'text_wrap':True,'valign':'vcenter'}}.items():fm[key]=wb.add_format({'font_name':'Aptos','font_size':11,'font_color':'202833',**extra})
 dashboards=[('EXEC V06','Visão Executiva','Executive Overview'),('DASH 01 V06','Investimento e Construção','Investment & Construction'),('DASH 02 V06','Fiscalidade, Incentivos e Seguros','Tax, Incentives & Insurance'),('DASH 03 V06','Consolidado do Projeto','Project Consolidated'),('DASH 04A V06','Promoção Imobiliária','Real Estate Development'),('DASH 04B V06','Residencial Sell & Leaseback','Residential Sell & Leaseback'),('DASH 05 V06','Hotel e MICE','Hotel & MICE'),('DASH 06 V06','Acionistas e Blocos','Shareholding & Investment Blocks'),('DASH 07 V06','Retornos e Waterfall','Equity Returns & Waterfall')]
 sheets={n:wb.add_worksheet(n) for n,_,_ in dashboards}
 # Central single input; all dashboard controls navigate here.
 settings=wb.add_worksheet('V06 Controlos');settings.set_column('A:A',36);settings.set_column('B:B',24);settings.set_column('C:C',13);settings.set_column('D:D',65);settings.set_column('E:E',38);settings.freeze_panes(4,1)
 settings.write_row(3,0,['Parâmetro PT','Valor / Value','Unidade','Fonte / Source / Limitation','English'],fm['head'])
 params={};paramrows={}
 lists={'language':['PT','EN'],'hotel_scenario':list(legacy.HOTELS),'div_policy':['Auto','Minimum','Intermediate','Maximum','Partial','Suspended'],'hot_state':['Pending','Pay','Released'],'profit_mode':['Proxy','Documented'],'convention':['Em preparação','Submetida','Em negociação','Aprovada','Não aprovada']}
 for r,(key,pt,en,value,unit,note) in enumerate(DEFAULTS,5):
  paramrows[key]=r;settings.write(r-1,0,pt);settings.write(r-1,1,value,fm['input']);settings.write(r-1,2,unit);settings.write(r-1,3,note);settings.write(r-1,4,en)
  settings.write_comment(r-1,1,note+'\nBase setting: '+str(value))
  if key in lists:settings.data_validation(r-1,1,r-1,1,{'validate':'list','source':lists[key]})
  elif unit.startswith('%'):settings.data_validation(r-1,1,r-1,1,{'validate':'decimal','criteria':'between','minimum':0,'maximum':1})
  elif unit in ['month','year','0/1','count']:settings.data_validation(r-1,1,r-1,1,{'validate':'integer','criteria':'between','minimum':0,'maximum':180 if unit!='0/1' else 1})
  else:settings.data_validation(r-1,1,r-1,1,{'validate':'decimal','criteria':'>=','value':0})
  params[key]=X(f"'V06 Controlos'!$B${r}",value)
 # Reject unsafe parameters, with business-specific bounds.
 for key,lo,hi in [('vat_lag',0,60),('tax_lag',1,12),('dep_life',1,100),('cve',1,1000),('rate',0,.3),('term',1,360),('coverage',1,10),('buffer',0,3),('construction_delay',0,60)]:
  r=paramrows[key]-1;settings.data_validation(r,1,r,1,{'validate':'decimal' if key in ['rate','coverage','cve'] else 'integer','criteria':'between','minimum':lo,'maximum':hi})
 settings.write_url('A1',"internal:'EXEC V06'!A1",string='← Executive Overview')
 # Immutable base settings retained alongside editable controls, no fake reset button.
 settings.write_row(3,6,['Key','Base setting'],fm['head'])
 for r,item in enumerate(DEFAULTS,4):settings.write(r,6,item[0]);settings.write(r,7,item[3])
 model=Model(wb,fm,params,[],legacy.CHAPTERS)
 units=model.sheet('V06 Unidades',['ID','Tipo','Piso','Área m²','Interior','Varanda','Preço origem EUR','Preço contrato EUR','Mês venda','Estado documental','Recebido efetivo EUR','Mês escritura','Preço/m²','Prémio vista x','Estado comercial','Data efetiva escritura','ITI comprador EUR','ITI KK EUR','Atos KK EUR'])
 dataset=[]
 for r,u in enumerate(legacy.UNITS,5):
  d={}
  for c,k in enumerate(['id','tipo','piso','area','interior','varanda','price_source','price','sale']):
   units.write(r-1,c,u[k],fm['input'] if k in ['price','sale'] else fm['text']);d[k]=X(f"'V06 Unidades'!${col(c)}${r}",u[k])
  units.write(r-1,9,'Pending external validation');units.write(r-1,10,None,fm['input']);units.write(r-1,11,u['deed'],fm['input']);d['deed']=X(f"'V06 Unidades'!$L${r}",u['deed']);dataset.append(d)
  model.cell(units,r,12,d['price']/d['area']);units.write(r-1,13,1,fm['input']);units.write(r-1,14,'Planned');units.write(r-1,15,None,fm['input']);model.cell(units,r,16,d['price']*params['transfer_rate']*(1-params['transfer_kk']));model.cell(units,r,17,d['price']*params['transfer_rate']*params['transfer_kk']);model.cell(units,r,18,(params['deed_cve']+params['cert_cve'])/params['cve']*params['deed_kk'])
 units.autofilter(3,0,97,18);model.units=dataset
 curve=model.sheet('V06 Curva',['Mês fonte V04','Peso mensal V04']);
 for i in range(30):curve.write(i+4,0,i+1);curve.write(i+4,1,sum(c['share']*c['weights'][i] for c in legacy.CHAPTERS))
 chapters=model.sheet('V06 Capitulos',['ID','Capítulo','Peso','Orçamento base','Contratado documentado','Realizado documentado','ETC previsão','EAC','Desvio EAC','Fonte/estado'])
 for r,c in enumerate(legacy.CHAPTERS,5):
  chapters.write(r-1,0,c['id']);chapters.write(r-1,1,c['name']);chapters.write(r-1,2,c['share'],fm['pct']);budget=model.cell(chapters,r,3,asx(c['share'])*56400000*(1+params['overrun']));chapters.write(r-1,4,None,fm['input']);chapters.write(r-1,5,None,fm['input']);etc=model.cell(chapters,r,6,budget-X(f"N('V06 Capitulos'!F{r})",0));eac=model.cell(chapters,r,7,etc+X(f"N('V06 Capitulos'!F{r})",0));model.cell(chapters,r,8,eac-budget);chapters.write(r-1,9,'V04 weights; no actual evidence supplied')
 capital=model.sheet('V06 Capital',['Bloco','Classe','Ticket EUR','Participação','Mês planeado','Fração primária KK','Fração recebida cenário','Recebido efetivo EUR','Mês efetivo','Natureza','Subscrito documental EUR','Recebido cenário EUR','Transferência KK EUR','Secundário EUR','Estado'])
 tickets=[]
 for r,(code,cat,amount,share,month) in enumerate(TICKETS,5):
  vals=[code,cat,amount,share,month,1,1,None,None,'Primary',None]
  refs=[]
  for c,v in enumerate(vals):capital.write(r-1,c,v,fm['input'] if c in [4,5,6,7,8,10] else fm['pct'] if c==3 else fm['text']);refs.append(X(f"'V06 Capital'!${col(c)}${r}",v or 0))
  tickets.append((code,cat,refs[2],refs[3],refs[4],refs[5],refs[6],refs[7],refs[8]))
  planned=model.cell(capital,r,11,refs[2]*refs[6]*params['capital_fraction']);model.cell(capital,r,12,planned*refs[5]);model.cell(capital,r,13,planned*(1-refs[5]));capital.write(r-1,14,'Assumed calendar; actual cells blank')
 model.tickets=tickets
 # Reference model plus independent, live Excel scenario engines.
 main=model.calculate('V06 Motor')
 scenario_defs=[('Base',{}),('VAT partial',{'vat_recovery':.5}),('VAT delayed',{'vat_lag':18}),('VAT none',{'vat_recovery':0}),('IRPC 3 years',{'exempt_years':3}),('IRPC 5 years',{'exempt_years':5}),('IRPC full',{'exempt_years':15}),('Insurance +25%',{'ins_factor':1.25}),('Insurance +50%',{'ins_factor':1.5}),('Operator 4/3/10',{'base_fee':.04,'incentive':.10}),('Div Minimum',{'div_policy':'Minimum'}),('Div Intermediate',{'div_policy':'Intermediate'}),('Div Maximum',{'div_policy':'Maximum'}),('Div Partial',{'div_policy':'Partial'}),('Div Suspended',{'div_policy':'Suspended'}),('Exit 10x',{'exit_multiple':10}),('Exit 15x',{'exit_multiple':15}),('Exit M120',{'exit_month':120}),('Combined stress',{'capital_delay':6,'sales_delay':6,'vat_lag':18,'ins_factor':1.5,'overrun':.1,'rate':.08,'hotel_scenario':'Conservador'}),('Capital delayed',{'capital_delay':6}),('Sales delayed',{'sales_delay':6}),('Construction +6M',{'construction_delay':6}),('Interest 8%',{'rate':.08}),('Hotel conservative',{'hotel_scenario':'Conservador'}),('Capital 50%',{'capital_fraction':.5})]
 scenario_models=[main]
 for i,(label,overrides) in enumerate(scenario_defs[1:],1):
  print('Building scenario',i,label,flush=True);scenario_models.append(model.calculate(f'V06 S{i:02}',overrides))
 metrics={};ambiguities=[];summary=model.sheet('V06 Cenarios',['Cenário','Caixa M30','Caixa M42','Caixa M54','Dividendos M54','Pico dívida','Défice máximo','Falta reserva','IVA pico','Equity Value','TIR novos','MOIC novos','VPL novos','EBITDA Ano15','Dividendos totais','Juros','IRPC pago','Seguros','Retorno condicional','VPL projeto','TIR projeto'])
 # Individual block cash flows; every scenario has its own genuine 181-period vectors.
 flows=model.sheet('V06 Fluxos Invest',['Month']+[t[0] for t in TICKETS]+['New investors','Project unlevered','M120 alternative'])
 blockmetrics=[]
 def getmetric(mod,label,rowindex):
  rows=mod['rows'];p=mod['params'];sn=mod['sheet'];c=mod['columns'];last=int(p['exit_month'].v);f=rows[last-1]
  def periodrange(key,first=1,last=180):return rng([x[key] for x in rows[first-1:last]])
  # INDEX allows changing the active exit month after generation.
  def at(key):return X(f'INDEX(\'{sn}\'!${col(c[key])}$5:${col(c[key])}$184,{p["exit_month"].f})',f[key].v)
  year15=periodrange('ebitda',169,180);exit_ebitda=X(f'SUM(INDEX(\'{sn}\'!${col(c["ebitda"])}$5:${col(c["ebitda"])}$184,{p["exit_month"].f}-11):INDEX(\'{sn}\'!${col(c["ebitda"])}$5:${col(c["ebitda"])}$184,{p["exit_month"].f}))',sum(x['ebitda'].v for x in rows[last-12:last]))
  ev=exit_ebitda*p['exit_multiple'];exitvalue=mx(0,ev*(1-p['exit_cost'])-at('debt')+at('cash')-at('hotusa_reserve')-at('tax_payable'))
  # Dedicated scenario cash-flow columns are compact and separately auditable.
  scflow=wb.add_worksheet('V06 F'+str(rowindex).zfill(2));scflow.write_row(3,0,['Mês','Fluxo investidores','Projeto sem dívida','Saída equity','Saída projeto'],fm['head']);newcf=[];pcf=[]
  for m in range(0,181):
   rr=m+5;scflow.write(rr-1,0,m)
   if m==0:ncf=asx(0);projectcf=asx(0);ex=asx(0);pex=asx(0)
   else:
    d=rows[m-1];ex=IF(asx(m).eq(p['exit_month']),exitvalue,0);pex=IF(asx(m).eq(p['exit_month']),mx(0,ev*(1-p['exit_cost'])-at('tax_payable')-at('hotusa_reserve')),0)
    new_dividend=d['dividend']*IF(d['dividend_target']>=4150000,1900000/4150000,IF(d['dividend_target']>=3575000,1650000/3575000,1400000/3000000));ncf=IF(asx(m)<=p['exit_month'],new_dividend-d['investor_outlay']+ex*.5,0);projectcf=IF(asx(m)<=p['exit_month'],d['project_cf']+pex,0)
   newcf.append(model.cell(scflow,rr,1,ncf));pcf.append(model.cell(scflow,rr,2,projectcf));model.cell(scflow,rr,3,ex);model.cell(scflow,rr,4,pex)
  cf_range=f"'{scflow.name}'!$B$5:$B$185";prange=f"'{scflow.name}'!$C$5:$C$185"
  ni=legacy.irr([x.v for x in newcf]);proots=legacy.irr_roots([x.v for x in pcf]);pi=min(proots,key=lambda a:abs(math.log1p(a)/12-math.log1p(.01))) if proots else None;irr=X(f'IFERROR((1+IRR({cf_range},0.005))^12-1,"Indisponível")',ni if ni is not None else 'Indisponível');projectirr=X(f'IFERROR((1+IRR({prange},0.01))^12-1,"Indisponível")',pi if pi is not None else 'Indisponível')
  if len(proots)>1:ambiguities.append({'scenario':label,'annual_roots':proots,'selected_root':pi,'criterion':'Root nearest Excel monthly IRR guess 1%; prefer NPV when roots are multiple'})
  paid=sum(max(0,-x['investor_outlay'].v) for x in []) # replaced below by gross paid capital
  paid_expr=X(f'SUM(INDEX(\'{sn}\'!${col(c["investor_outlay"])}$5:${col(c["investor_outlay"])}$184,1):INDEX(\'{sn}\'!${col(c["investor_outlay"])}$5:${col(c["investor_outlay"])}$184,{p["exit_month"].f}))',sum(x['investor_outlay'].v for x in rows[:last]))
  divstoexit=X(f'SUM(INDEX(\'{sn}\'!${col(c["dividend"])}$5:${col(c["dividend"])}$184,1):INDEX(\'{sn}\'!${col(c["dividend"])}$5:${col(c["dividend"])}$184,{p["exit_month"].f}))',sum(x['dividend'].v for x in rows[:last]))
  moic=IF(paid_expr>0,(rng(newcf)+paid_expr)/mx(.00001,paid_expr),0)
  npv=SUM(x/(1+p['discount'])**(m/12) for m,x in enumerate(newcf))
  npv=X(f'NPV((1+{p["discount"].f})^(1/12)-1,\'{scflow.name}\'!$B$6:$B$185)+\'{scflow.name}\'!B5',npv.v)
  pnpv=SUM(x/(1+p['project_discount'])**(m/12) for m,x in enumerate(pcf));pnpv=X(f'NPV((1+{p["project_discount"].f})^(1/12)-1,\'{scflow.name}\'!$C$6:$C$185)',pnpv.v)
  peak=lambda key:X(f'MAX(\'{sn}\'!${col(c[key])}$5:${col(c[key])}$184)',max(x[key].v for x in rows))
  result={'cash30':rows[29]['cash'],'cash42':rows[41]['cash'],'cash54':rows[53]['cash'],'div54':rows[53]['dividend'],'debt':peak('debt'),'gap':peak('funding_gap'),'reservegap':peak('reserve_gap'),'vatcredit':peak('vat_credit'),'exit':exitvalue,'irr':irr,'moic':moic,'npv':npv,'ebitda15':year15,'dividends':divstoexit,'interest':periodrange('interest'),'tax':periodrange('tax_paid'),'insurance':periodrange('construction_insurance')+periodrange('operating_insurance'),'conditional':IF(OR(peak('funding_gap')>.01,peak('reserve_gap')>.01,paid_expr<25000000),'Conditional','Study'),'pnpv':pnpv,'pirr':projectirr,'ev':ev,'paid':paid_expr}
  summary.write(rowindex+3,0,label)
  for j,key in enumerate(['cash30','cash42','cash54','div54','debt','gap','reservegap','vatcredit','exit','irr','moic','npv','ebitda15','dividends','interest','tax','insurance','conditional','pnpv','pirr'],1):result[key]=model.cell(summary,rowindex+4,j,result[key],'pct' if key in ['irr','pirr'] else 'text' if key=='conditional' else 'num')
  return result
 for i,(mod,(label,_)) in enumerate(zip(scenario_models,scenario_defs),1):metrics[label]=getmetric(mod,label,i)
 mainmetrics=metrics['Base'];rows=main['rows'];p=params
 # Ownership and returns for all fifteen blocks, preserving promoter historical-cost uncertainty.
 blocks=[('JSCV P1','JSCV','Platinum',.20,None),('JSCV P2','JSCV','Platinum',.20,None)]+[(f'JALS S{i}','JALS','Silver',.02,None) for i in range(1,6)]+[(t[0],'New',t[1],t[3],i) for i,t in enumerate(TICKETS)]
 own=model.sheet('V06 Blocos',['Bloco','Titular','Classe','Participação','Investido EUR','Dividendos EUR','Saída EUR','Recebido total EUR','MOIC','TIR anual','VPL','Cash-on-cash anual médio','Base histórica'])
 allblockflows=[]
 flows.write_row(3,1,[x[0] for x in blocks],fm['head'])
 flows.write_row(3,20,[x[0]+' dividend' for x in blocks],fm['head'])
 flows.write_row(3,36,[x[0]+' outlay' for x in blocks],fm['head'])
 for bi,(code,holder,cat,share,ti) in enumerate(blocks):
  cf=[];divs=[];invested=[];exit_b=mainmetrics['exit']*share
  for m in range(0,181):
   rr=m+5
   if m==0:cfv=asx(0);divv=asx(0);iv=asx(0)
   else:
    d=rows[m-1];target=d['dividend_target'];idx={'Platinum':0,'Gold':1,'Silver':2}[cat];targets=TARGETS[cat]
    blocktarget=IF(target>=4150000,targets[2],IF(target>=3575000,targets[1],targets[0]));denom=IF(target>=4150000,4150000,IF(target>=3575000,3575000,3000000));divv=IF(asx(m)<=p['exit_month'],d['dividend']*blocktarget/denom,0)
    if ti is not None:
     t=tickets[ti];paym=IF(t[8]>0,t[8],t[4]+p['capital_delay']);amount=IF(t[8]>0,t[7],t[2]*t[6]*p['capital_fraction']);iv=IF(AND(asx(m).eq(paym),asx(m)<=p['exit_month']),amount,0)
    else:iv=asx(0)
    cfv=divv-iv+IF(asx(m).eq(p['exit_month']),exit_b,0)
   dest=flows
   # 15 block vectors use columns B:P, aggregate Q, project R.
   divref=model.cell(dest,rr,20+bi,divv);ivref=model.cell(dest,rr,36+bi,iv);ref=model.cell(dest,rr,bi+1,divref-ivref+(IF(asx(m).eq(p['exit_month']),exit_b,0) if m else 0));cf.append(ref);divs.append(divref);invested.append(ivref)
   if bi==0:dest.write(rr-1,0,m)
  br=bi+5;paid=rng(invested);div=rng(divs);moic=IF(paid>0,(div+exit_b)/mx(.00001,paid),0);ir=legacy.irr([v.v for v in cf]);npv=SUM(x/(1+p['discount'])**(m/12) for m,x in enumerate(cf))
  own.write_row(br-1,0,[code,holder,cat]);own.write(br-1,3,share,fm['pct']);model.cell(own,br,4,paid);model.cell(own,br,5,div);model.cell(own,br,6,exit_b);model.cell(own,br,7,div+exit_b)
  model.cell(own,br,8,moic if ti is not None else asx('Unknown historical cost'),'num' if ti is not None else 'text')
  xirr=X(f'IFERROR((1+IRR(\'V06 Fluxos Invest\'!{col(bi+1)}5:{col(bi+1)}185,0.005))^12-1,"Indisponível")',ir if ir is not None else 'Indisponível')
  model.cell(own,br,9,xirr if ti is not None else asx('Unknown historical dates'),'pct' if ti is not None else 'text');npv=X(f'NPV((1+{p["discount"].f})^(1/12)-1,\'V06 Fluxos Invest\'!{col(bi+1)}6:{col(bi+1)}185)',npv.v);model.cell(own,br,10,npv if ti is not None else asx('Not computed'),'num' if ti is not None else 'text');model.cell(own,br,11,IF(paid>0,div/mx(.00001,paid)/(p['exit_month']/12),0),'pct');own.write(br-1,12,'Individual promoter costs/dates unavailable' if ti is None else 'Scenario subscription, not confirmed receipt');allblockflows.append(cf)
 for m in range(181):
  model.cell(flows,m+5,16,SUM(allblockflows[i][m] for i in range(7,15)))
 # Tax, annual and quarterly reporting reconcile directly to the current engine.
 annual=model.sheet('V06 Anual',['Ano','Receita Hotel','EBITDA','IRPC apurado','IRPC pago','Seguros','Dividendos','Caixa fecho','Dívida fecho','Yield proprietários','CFADS','Serviço dívida','DSCR'])
 for y in range(1,16):
  r=y+4;rr=rows[(y-1)*12:y*12];annual.write(r-1,0,y)
  for j,k in enumerate(['revenue','ebitda','annual_tax','tax_paid','operating_insurance','dividend'],1):model.cell(annual,r,j,rng([x[k] for x in rr]))
  model.cell(annual,r,7,rr[-1]['cash']);model.cell(annual,r,8,rr[-1]['debt']);model.cell(annual,r,9,rng([x['owner_distribution'] for x in rr]));a=model.cell(annual,r,10,rng([x['cfads'] for x in rr]));b=model.cell(annual,r,11,rng([x['debt_service'] for x in rr]));model.cell(annual,r,12,IF(b>0,a/mx(.00001,b),0))
 quarterly=model.sheet('V06 Trimestral',['Trimestre','CAPEX','Capital KK','Banco','Recebimentos','IVA','Reembolso','Seguro obra','IRPC pago','Caixa fecho'])
 for q in range(1,61):
  rr=rows[(q-1)*3:q*3];r=q+4;quarterly.write(r-1,0,q)
  for j,k in enumerate(['capex','equity','draw','receipts','vat_incurred','vat_refund','construction_insurance','tax_paid'],1):model.cell(quarterly,r,j,rng([x[k] for x in rr]))
  model.cell(quarterly,r,9,rr[-1]['cash'])
 # Bilingual dictionary uses short definitions; detailed provenance is retained in source and manual.
 dictionary=model.sheet('V06 Dicionario',['KPI / sigla','Nome PT','Name EN','Definição PT','Definition EN','Método / Method','Fonte / Source','Unidade','Interpretação PT','Interpretation EN','Direção / Direction','Limite / Threshold','Limitações / Limitations'])
 definitions=[('IRR','TIR','IRR','Taxa que anula o VPL dos fluxos do investidor.','Discount rate making investor NPV zero.','(1+IRR(monthly cash flows))^12-1','V06 Fluxos Invest','%','Comparar ao retorno exigido.','Compare with required return.','Higher','Discount rate','May have multiple roots; contingent on exit'),('NPV','VPL','NPV','Valor descontado dos fluxos de caixa.','Discounted cash-flow value.','Sum CF/(1+annual rate)^(m/12)','V06 Cenarios','EUR','Negativo indica retorno inferior à taxa exigida.','Negative means return below discount rate.','Higher','0','Nominal, pre-investor tax'),('MOIC','Múltiplo capital','MOIC','Recebimentos totais / capital investido.','Total proceeds / invested capital.','(Dividends+exit)/capital','V06 Blocos','x','Não mede tempo.','Does not measure timing.','Higher','1','Estimated exit included'),('GDV','VGV','GDV','Soma dos preços de 94 frações.','Sum of 94 apartment prices.','SUM(contract price)','V06 Unidades','EUR','Vendas previstas, não recebidas.','Planned sales, not actual receipts.','Context','None','Commercial validation pending'),('CAPEX','Investimento futuro','Future CAPEX','Obra, fiscalização, contingência e ETC adicional.','Construction, supervision, contingency and incremental ETC.','56.4m+1.128m+5.64m+ETC','V06 Motor','EUR','Histórico excluído.','Historical investment excluded.','Budget','63168000 base','VAT and insurance shown separately'),('EBITDA','Resultado operacional','EBITDA','GOP menos fees e seguros operacionais.','GOP less operator fees and operating insurance.','GOP-fees-insurance','V06 Motor','EUR','Antes de depreciação, juros e IRPC.','Before depreciation, interest and corporate tax.','Higher','0','Forecast, not guaranteed'),('ADR','Tarifa média diária','ADR','Preço médio de quarto vendido.','Average sold room rate.','V04 scenario × growth','V06 Motor','EUR/night','Avaliar com ocupação.','Assess with occupancy.','Context','None','Commercial assumption'),('RevPAR','Receita por quarto disponível','RevPAR','ADR multiplicado pela ocupação.','ADR multiplied by occupancy.','ADR × occupancy','V06 Motor','EUR/night','Não inclui receitas MICE.','Excludes MICE revenue.','Higher','None','Room revenue only'),('DSCR','Cobertura dívida','DSCR','CFADS Hotel / serviço obrigatório.','Hotel CFADS / mandatory debt service.','CFADS / interest+principal+commitment','V06 Anual','x','Não inclui amortizações voluntárias.','Excludes voluntary sweep.','Higher','1.30 indicative','Hotel allocation excludes consolidated corporate tax; see manual'),('CFADS','Caixa antes serviço dívida','CFADS','EBITDA menos FF&E, reposições, FM e IPI.','EBITDA less FF&E, replacement, working capital and property tax.','EBITDA-FFE-maintenance-dWC-IPI','V06 Motor','EUR','Imposto consolidado apresentado à parte.','Consolidated tax shown separately.','Higher','0','Corporate tax allocation pending'),('EAC','Custo previsto conclusão','Estimate at Completion','Realizado documentado mais ETC.','Documented actual plus ETC.','Actual + ETC','V06 Capitulos','EUR','Células vazias não são pagamentos comprovados.','Blank cells are not verified payments.','Budget','Approved budget','Forecast proxy'),('ETC','Custo por concluir','Estimate to Complete','Previsão de custo restante.','Forecast remaining cost.','Forecast EAC - actual','V06 Capitulos','EUR','Separar compromisso e realizado.','Separate commitments from actuals.','Context','None','Not an accounting ledger'),('Equity','Valor de saída equity','Equity Value','EV menos dívida, custos e obrigações mais caixa elegível.','EV less debt, costs, obligations plus eligible cash.','EV-cost-debt+cash-HOTUSA-tax payable','V06 Cenarios','EUR','Reserva protegida excluída.','Protected legal reserve excluded.','Higher','None','Not contracted transaction'),('EV','Valor da empresa','Enterprise Value','EBITDA últimos 12 meses × múltiplo.','Trailing 12-month EBITDA × multiple.','EBITDA × 10/12/15','V06 Cenarios','EUR','Exclui apartamentos vendidos.','Excludes sold apartments.','Context','None','Uncontracted exit'),('Sweep','Amortização voluntária','Cash Sweep','Caixa livre usada para reduzir dívida.','Free cash used to reduce debt.','MIN(debt,excess cash)','V06 Motor','EUR','Não repõe limite de crédito.','Does not replenish facility.','Context','Non-revolving','Subject to reserves'),('VAT','Crédito IVA','VAT Credit','IVA elegível menos reembolsos simulados.','Eligible VAT less simulated refunds.','Prior credit+eligible-refund','V06 Motor','EUR','Recuperação é hipótese de estudo.','Recovery is a study assumption.','Lower exposure','0','No automatic cross-entity offset'),('Yield','Rendimento cash-on-cash','Cash-on-cash Yield','Dividendos recebidos / capital investido.','Cash dividends / invested capital.','Annual dividends / capital','V06 Blocos','%','Não confundir com TIR ou yield garantido.','Not IRR or a guaranteed yield.','Higher','Class objectives','Pre-investor tax')]
 dictrows={}
 for r,definition in enumerate(definitions,5):dictionary.write_row(r-1,0,definition);dictrows[definition[0]]=r;dictionary.set_row(r-1,45)
 dictionary.set_column('D:M',35);dictionary.autofilter(3,0,len(definitions)+3,12)
 # Central translations are formulas, with no impact on finance.
 translations=model.sheet('V06 Traducoes',['Key','PT','EN','Ativo / Active']);trcount=4
 def tr(key,pt,en):
  nonlocal trcount
  trcount+=1;r=trcount;translations.write(r-1,0,key);translations.write(r-1,1,pt);translations.write(r-1,2,en);translations.write_formula(r-1,3,f'=IF({params["language"].f}="PT",B{r},C{r})',fm['text'],pt);return f"'V06 Traducoes'!D{r}"
 nav=[tr('nav'+str(i),pt,en) for i,(_,pt,en) in enumerate(dashboards)]
 opening=lambda key:X("INDEX('V06 Motor'!$"+col(main['columns'][key])+"$5:$"+col(main['columns'][key])+"$184,31+"+p['construction_delay'].f+")",rows[30][key].v)
 card_defs={
 'EXEC V06':[('VGV','GDV',X("SUM('V06 Unidades'!H5:H98)",sum(u['price'] for u in legacy.UNITS)),'GDV','card'),('CAPEX futuro','Future CAPEX',rng([x['capex'] for x in rows]),'CAPEX','card'),('Pico dívida','Peak bank debt',mainmetrics['debt'],'DSCR','card'),('TIR novos investidores','New investor IRR',mainmetrics['irr'],'IRR','cardpct'),('MOIC novos investidores','New investor MOIC',mainmetrics['moic'],'MOIC','cardx'),('VPL investidores','Investor NPV',mainmetrics['npv'],'NPV','card')],
 'DASH 01 V06':[('Histórico, sem saída futura','Historical, no future cash outflow',asx(4500000),'CAPEX','card'),('CAPEX futuro','Future CAPEX',rng([x['capex'] for x in rows]),'CAPEX','card'),('Construção e FF&E inicial','Construction and initial FF&E',rng([x['construction'] for x in rows]),'CAPEX','card'),('Fiscalização','Supervision',rng([x['supervision'] for x in rows]),'EAC','card'),('Contingência utilizada','Contingency used',rng([x['contingency_paid'] for x in rows]),'ETC','card'),('ETC adicional','Incremental ETC',p['eac_override'],'ETC','card')],
 'DASH 02 V06':[('IVA suportado','VAT incurred',rng([x['vat_incurred'] for x in rows]),'VAT','card'),('Pico crédito IVA','Peak VAT credit',mainmetrics['vatcredit'],'VAT','card'),('IRPC pago','Corporate tax paid',mainmetrics['tax'],'NPV','card'),('Seguro de obra','Construction insurance',rng([x['construction_insurance'] for x in rows]),'CAPEX','card'),('Seguro operação total','Total operating insurance',rng([x['operating_insurance'] for x in rows]),'EBITDA','card'),('ITI KK total','KK transfer tax',rng([x['transfer_kk'] for x in rows]),'CAPEX','card')],
 'DASH 03 V06':[('Caixa M30','Cash M30',mainmetrics['cash30'],'CFADS','card'),('Caixa M42','Cash M42',mainmetrics['cash42'],'CFADS','card'),('Caixa M54','Cash M54',mainmetrics['cash54'],'CFADS','card'),('Défice de caixa máximo','Maximum funding gap',mainmetrics['gap'],'DSCR','card'),('Insuficiência reserva','Reserve shortfall',mainmetrics['reservegap'],'CFADS','card'),('Juros totais','Total interest',mainmetrics['interest'],'DSCR','card')],
 'DASH 04A V06':[('Frações','Apartments',asx(94),'GDV','card'),('VGV','GDV',X("SUM('V06 Unidades'!H5:H98)",sum(u['price'] for u in legacy.UNITS)),'GDV','card'),('Recebimentos 180M','180M receipts',rng([x['receipts'] for x in rows]),'GDV','card'),('Custos comerciais','Sales costs',rng([x['sales_cost'] for x in rows]),'GDV','card'),('Atos e propriedade horizontal','Deeds and horizontal property',rng([x['deeds_cost'] for x in rows]),'CAPEX','card'),('Margem promoção proxy','Development margin proxy',rng([x['sales_profit'] for x in rows]),'EBITDA','card')],
 'DASH 04B V06':[('Receita pool 15A','Pool revenue 15Y',rng([x['res_revenue'] for x in rows]),'Yield','card'),('Rendas proprietários','Owner distributions',rng([x['owner_distribution'] for x in rows]),'Yield','card'),('Yield não coberto','Uncovered yield target',rng([x['owner_shortfall'] for x in rows]),'Yield','card'),('Residual KK','KK residual',rng([x['res_kk'] for x in rows]),'CFADS','card'),('FF&E pool','Pool FF&E',rng([x['res_ffe'] for x in rows]),'CAPEX','card'),('Fundo adicional pool','Additional pool reserve',rng([x['res_fund'] for x in rows]),'CAPEX','card')],
 'DASH 05 V06':[('ADR abertura','Opening ADR',opening('adr'),'ADR','card'),('Ocupação abertura','Opening occupancy',opening('occupancy'),'ADR','cardpct'),('RevPAR abertura','Opening RevPAR',opening('adr')*opening('occupancy'),'RevPAR','card'),('EBITDA Ano15','Year15 EBITDA',mainmetrics['ebitda15'],'EBITDA','card'),('Fees operador 15A','Operator fees 15Y',rng([x['base_fee'] for x in rows])+rng([x['systems_fee'] for x in rows])+rng([x['incentive_fee'] for x in rows]),'CFADS','card'),('CFADS 15A','CFADS 15Y',rng([x['cfads'] for x in rows]),'CFADS','card')],
 'DASH 06 V06':[('Blocos totais','Total blocks',asx(15),'MOIC','card'),('Blocos novos','New blocks',asx(8),'MOIC','card'),('Capital alvo','Capital target',asx(25000000),'MOIC','card'),('Capital KK de cenário','Scenario KK capital',rng([x['equity'] for x in rows]),'MOIC','card'),('JSCV','JSCV',asx(.4),'MOIC','cardpct'),('JALS','JALS',asx(.1),'MOIC','cardpct')],
 'DASH 07 V06':[('Dividendos M54','M54 dividends',mainmetrics['div54'],'Yield','card'),('Dividendos até saída','Dividends through exit',mainmetrics['dividends'],'Yield','card'),('Equity Value','Equity Value',mainmetrics['exit'],'Equity','card'),('TIR investidores','Investor IRR',mainmetrics['irr'],'IRR','cardpct'),('MOIC investidores','Investor MOIC',mainmetrics['moic'],'MOIC','cardx'),('VPL investidores','Investor NPV',mainmetrics['npv'],'NPV','card')]
 }
 for di,(name,pt,en) in enumerate(dashboards):
  s=sheets[name];s.hide_gridlines(2);s.set_zoom(80);s.freeze_panes(5,4);s.set_column('A:A',2);s.set_column('B:C',13);s.set_column('D:D',2);s.set_column('E:P',12);s.set_column('Q:Q',2)
  for rr in range(0,62):s.set_row(rr,22);s.write_blank(rr,0,None,fm['side']);s.write_blank(rr,1,None,fm['side']);s.write_blank(rr,2,None,fm['side']);s.write_blank(rr,3,None,fm['side'])
  s.merge_range('B2:C3','JSCV',fm['active']);s.merge_range('B4:C5','KK-01\nV06 EXECUTIVE',fm['side'])
  for ni,(target,_,_) in enumerate(dashboards):
   rr=7+ni*4;s.merge_range(rr-1,1,rr+1,2,'',fm['active'] if ni==di else fm['side']);f='HYPERLINK("#\''+target+'\'!A1",'+nav[ni]+')';s.write_formula(rr-1,1,'='+f,fm['active'] if ni==di else fm['side'],dashboards[ni][1])
  s.merge_range('E2:P3','',fm['title']);s.write_formula('E2','='+nav[di],fm['title'],pt)
  s.merge_range('E4:L4','KK HOTEL & RESIDENCES  ·  PRAIA  ·  EUR  ·  180M',fm['sub'])
  s.merge_range('M4:P4','',fm['sub']);s.write_url('M4',"internal:'V06 Controlos'!B5",fm['sub'],string='PT / EN  ·  Controlos / Controls')
  s.merge_range('E5:P5','',fm['sub']);status=tr('status'+str(di),'Estudo financeiro · cenário editável · Excel 365: teste nativo pendente','Financial study · editable scenario · Excel 365: native test pending');s.write_formula('E5','='+status,fm['sub'],'Estudo financeiro · Excel 365 pendente')
  for ci,(lpt,len_,value,dkey,fmt) in enumerate(card_defs[name]):
   rr=7+(ci//3)*6;cc=4+(ci%3)*4;label=tr('card'+str(di)+'_'+str(ci),lpt,len_)
   s.merge_range(rr-1,cc,rr,cc+2,'',fm['label']);s.write_formula(rr-1,cc,'='+label,fm['label'],lpt);s.write_url(rr-1,cc+3,f"internal:'V06 Dicionario'!A{dictrows[dkey]}",fm['label'],string='ⓘ')
   s.merge_range(rr+1,cc,rr+3,cc+3,'',fm[fmt]);s.write_formula(rr+1,cc,'='+value.f,fm[fmt],value.v)
  sect=tr('analysis'+str(di),'ANÁLISE FINANCEIRA E CENÁRIOS','FINANCIAL ANALYSIS & SCENARIOS');s.merge_range('E20:P20','',fm['section']);s.write_formula('E20','='+sect,fm['section'],'ANÁLISE FINANCEIRA E CENÁRIOS')
  # Chart series use genuine computed values and translated titles.
  def chart(keys,charttype='line'):
   ch=wb.add_chart({'type':charttype});
   for key,pt_,en_ in keys:
    title=tr('series'+str(di)+key,pt_,en_);c=main['columns'][key];ch.add_series({'name':'='+title,'categories':['V06 Motor',4,0,183,0],'values':['V06 Motor',4,c,183,c],'line':{'width':1.8}})
   ch.set_legend({'position':'bottom','font':{'size':10}});ch.set_chartarea({'border':{'none':True},'fill':{'color':'FFFFFF'}});ch.set_plotarea({'border':{'none':True}});ch.set_y_axis({'num_format':'0,," M€"','major_gridlines':{'visible':True,'line':{'color':'E5E7EB'}}});ch.set_x_axis({'name':'Month / Mês','label_interval':12});ch.set_size({'width':540,'height':270});return ch
  keysets=[ [('cash','Caixa','Cash'),('reserve','Reserva protegida','Protected reserve')], [('capex','CAPEX','CAPEX'),('construction_insurance','Seguro obra','Construction insurance')], [('vat_credit','Crédito IVA','VAT credit'),('tax_paid','IRPC pago','Tax paid')], [('debt','Dívida','Debt'),('cash','Caixa','Cash')], [('receipts','Recebimentos','Receipts'),('deed_sales','Escrituras','Deed sales')], [('owner_distribution','Renda suportável','Affordable owner distribution'),('owner_target','Renda alvo','Owner target')], [('ebitda','EBITDA','EBITDA'),('revenue','Receita','Revenue')], [('equity','Capital KK','KK capital'),('investor_outlay','Investimento','Investor outlay')], [('dividend','Dividendos','Dividends'),('cash','Caixa','Cash')] ]
  s.insert_chart('E22',chart(keysets[di]),{'x_scale':1.55,'y_scale':1.05})
  rr=36;s.merge_range('E36:P36','',fm['section']);tabletitle=tr('table'+str(di),'CONTROLOS E COMPARATIVOS','CONTROLS & COMPARISONS');s.write_formula('E36','='+tabletitle,fm['section'],'CONTROLOS E COMPARATIVOS')
  # Contextual tables and direct technical navigation.
  if di in [0,2,3,8]:
   labels=[('Base','Base'),('VAT delayed','IVA diferido'),('IRPC 3 years','IRPC isenção 3A'),('Insurance +50%','Seguros +50%'),('Combined stress','Stress combinado')]
   heads=['Cenário / Scenario','TIR / IRR','MOIC','VPL / NPV']
   for ci,h in enumerate(heads):s.merge_range(rr,4+ci*3,rr,6+ci*3,h,fm['head'])
   for ri,(key,label) in enumerate(labels,rr+1):
    s.merge_range(ri,4,ri,6,label,fm['text'])
    for ci,mk in enumerate(['irr','moic','npv'],1):s.merge_range(ri,4+ci*3,ri,6+ci*3,'',fm['pct'] if mk=='irr' else fm['num']);s.write_formula(ri,4+ci*3,'='+metrics[key][mk].f,fm['pct'] if mk=='irr' else fm['num'],metrics[key][mk].v)
  else:
   target={1:'V06 Capitulos',4:'V06 Unidades',5:'V06 Anual',6:'V06 Anual',7:'V06 Capital'}[di]
   s.merge_range('E38:P39','',fm['note']);s.write_url('E38',f"internal:'{target}'!A1",fm['note'],string='Abrir detalhe financeiro e inputs / Open financial details and inputs → '+target)
   for j,(pt_,en_,expr) in enumerate([('Caixa M30','Cash M30',mainmetrics['cash30']),('Caixa M42','Cash M42',mainmetrics['cash42']),('Caixa M54','Cash M54',mainmetrics['cash54'])]):
    lab=tr('detail'+str(di)+str(j),pt_,en_);s.write_formula(40+j,4,'='+lab,fm['text'],pt_);s.write_formula(40+j,10,'='+expr.f,fm['num'],expr.v)
  s.merge_range('E46:P46','',fm['section']);interpret=tr('interpret'+str(di),'INTERPRETAÇÃO EXECUTIVA','EXECUTIVE INTERPRETATION');s.write_formula('E46','='+interpret,fm['section'],'INTERPRETAÇÃO EXECUTIVA')
  msg=tr('risk'+str(di),'Confirmar capital, vendas, impostos e contratos. Dividendos e saídas são simulações, não garantias.','Confirm capital, sales, taxes and contracts. Dividends and exit values are simulations, not guarantees.')
  s.merge_range('E48:P50','',fm['note']);s.write_formula('E48','='+msg,fm['note'],'Confirmar capital, vendas, impostos e contratos. Dividendos e saídas são simulações, não garantias.')
  s.merge_range('E52:P54','',fm['note']);s.write_formula('E52',f'=IF({mainmetrics["gap"].f}>0,IF({params["language"].f}="PT","ALERTA: défice de financiamento não coberto.","WARNING: unfunded financing gap."),IF({mainmetrics["npv"].f}<0,IF({params["language"].f}="PT","ALERTA: VPL negativo à taxa de retorno exigida.","WARNING: negative NPV at the required return."),IF({params["language"].f}="PT","Rever pressupostos e validações externas.","Review assumptions and external validations.")))',fm['note'],'Review assumptions')
  if di==2:
   s.merge_range('E56:P59','Convenção de Estabelecimento Turístico / Tourism Investment Convention: candidatura em preparação; incentivos pretendidos sujeitos a negociação, aprovação e formalização. Nenhum benefício concedido. / Application under preparation; proposed incentives require negotiation, approval and formalisation. No benefits granted.',fm['note'])
  else:s.merge_range('E56:P59','HOTUSA: reserva jurídica protegida de 1 M€ no cenário base; não é dívida reconhecida nem pagamento. / Protected legal contingency reserve in the base case; not a recognised debt or payment.',fm['note'])
  s.merge_range('E61:P62','Fonte / Source: V04 Fixed + mandato V06 + matriz V05 · 22/10/2026 · Estado / Status: CANDIDATE · Não certificado Excel 365 / Not Excel 365 certified',fm['sub'])
  s.set_portrait();s.set_paper(8);s.fit_to_pages(1,1);s.print_area('A1:Q62');s.set_margins(.2,.2,.25,.25)
  s.set_footer('&LKK-01 · V06 · Candidate&R&P / &N')
 sheets['EXEC V06'].activate();sheets['EXEC V06'].set_first_sheet()
 # Persist actual executed Python reference results and reproducible build metadata.
 reference={'metrics':{label:{key:v.v for key,v in vals.items()} for label,vals in metrics.items()},'main_rows':[{key:v.v for key,v in d.items()} for d in main['rows']],'scenarios':{label:[{key:v.v for key,v in d.items()} for d in mod['rows']] for mod,(label,_) in zip(scenario_models,scenario_defs)},'param_rows':paramrows,'main_columns':main['columns'],'dashboards':[x[0] for x in dashboards],'assumptions':['Eight calendars reconstructed as P01 M1, G01 M2, G02 M3, S01 M3, S02 M4, S03 M4, S04 M5, S05 M6; not recoverable exact V05 dates.','Zero property tax/customs/stamp inputs are unresolved exposures, not granted exemptions.','Project fiscal years, three-month tax payment lag and depreciation are study conventions.','V04 sheets retained as historical reference; V06 dashboard results use only V06 Motor.','No native Excel Desktop available.']}
 reference['multiple_irr_cases']=ambiguities
 wb.close();(OUT/'Referencia_Python_V06.json').write_text(json.dumps(reference,ensure_ascii=False,indent=2));print('V06 modules saved',flush=True)
 return reference
if __name__=='__main__':build()
