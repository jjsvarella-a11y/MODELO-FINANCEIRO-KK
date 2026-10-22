"""Institutional paginated report, using existing report expressions and V06 outputs."""
from xlsxwriter.utility import xl_col_to_name as col
from report_final import SECTIONS
from premium_ui import LANG,CENTER

def build_reports(wb,F,src,ref,data,chartcols):
 def tr(a,b):return f'IF({LANG}="PT","{a.replace(chr(34),chr(34)*2)}","{b.replace(chr(34),chr(34)*2)}")'
 def put(s,area,f,v,kind='text'):
  if area.startswith('I') and kind=='ratio':kind='ratio_small'
  s.merge_range(area,'',F[kind]);s.write_formula(area.split(':')[0],'='+f,F[kind],v)
 def text(s,area,a,b,kind='text'):put(s,area,tr(a,b),a,kind)
 def link(s,area,sh,cell,a,b):put(s,area,f'HYPERLINK("#\'{sh}\'!{cell}",{tr(a,b)})',a,'nav')
 def oldcell(sn,a):return src.label(sn,a)
 def structured(sn,page):
  off=page*50;cards=[];records=[]
  for r in [14,19,24]:
   for c in ['A','G']:
    f,v=src.get(sn,f'{c}{off+r+2}')
    if f is not None and isinstance(v,(float,int)):
     lab=src.label(sn,f'{c}{off+r}');cards.append((lab,(f,v),src.card_kind(sn,f'{c}{off+r+2}')))
  for r in range(25,46):
   f,v=src.get(sn,f'I{off+r}')
   if f is not None:records.append((src.label(sn,f'A{off+r}'),(f,v),src.format_kind(sn,f'I{off+r}') if isinstance(v,(int,float)) else 'text'))
  notes=[]
  for r in range(26,46):
   f,v=src.get(sn,f'A{off+r}')
   if f and isinstance(v,str) and len(v)>80 and not any(x[0][0]==f for x in records):notes.append((f,v))
  return cards,records,notes
 # Selected-case chart projection: same monthly financial outputs, selected globally.
 selectedcols={};offset=150;case="'V06 Hipoteses'!$D$27";cc=ref['main_columns'];rows=ref['main_rows'];rr=ref['scenarios']['Reference Case'];stocks={'cash','debt','reserve','vat_credit','ffe_restricted_fund','res_fund'}
 for key in ['capex','construction','supervision','contingency_paid','receipts','deed_sales','cash','debt','reserve','res_revenue','owner_distribution','revenue','ebitda','tax_paid','operator_fees','dividend','equity','investor_outlay']:
  c=offset+len(selectedcols);selectedcols[key]=c;data.write(3,c,key)
  for y in range(1,16):
   a=col(cc[key]);active=f"'V06 Motor'!{a}{y*12+4}" if key in stocks else f"SUM('V06 Motor'!{a}{(y-1)*12+5}:{a}{y*12+4})";frozen=active.replace('V06 Motor','V06 Referencia');v=rows[y*12-1][key] if key in stocks else sum(r[key] for r in rows[(y-1)*12:y*12]);data.write_formula(y+3,c,f'=IF({case}="Reference Case",{frozen},{active})',F['num'],v)
 qcols={'capex':202,'receipts':203}
 for q in range(10):
  data.write(q+4,201,q+1)
  for key,c in qcols.items():
   a=col(cc[key]);active=f"SUM('V06 Motor'!{a}{q*3+5}:{a}{q*3+7})";frozen=active.replace('V06 Motor','V06 Referencia');data.write_formula(q+4,c,f'=IF({case}="Reference Case",{frozen},{active})',F['num'],sum(r[key] for r in rows[q*3:q*3+3]))
 def plot(s,anchor,keys,width=626,height=178):
  ch=wb.add_chart({'type':'line'})
  for j,(key,pt,en) in enumerate(keys):
   c=qcols[key] if section==1 else selectedcols[key];nr=550+j+len(s.name);data.write_formula(nr,c,'='+tr(pt,en),F['text'],pt);ch.add_series({'name':['V06 Visual',nr,c],'categories':['V06 Visual',4,201 if section==1 else 0,13 if section==1 else 18,201 if section==1 else 0],'values':['V06 Visual',4,c,13 if section==1 else 18,c],'line':{'color':['142337','C8101E','1765A8'][j%3],'width':2}})
  ch.set_size({'width':width,'height':height});ch.set_chartarea({'border':{'none':True}});ch.set_legend({'position':'bottom','font':{'size':9,'name':'Aptos'}});ch.set_x_axis({'num_font':{'size':8}});ch.set_y_axis({'num_format':'0,," M€"','num_font':{'size':8},'major_gridlines':{'visible':True,'line':{'color':'E7EDF3'}}});s.insert_chart(anchor,ch)
 pagekeys=[['capex','equity'],['capex','receipts'],['cash','debt','reserve'],['receipts','deed_sales'],['res_revenue','owner_distribution'],['revenue','ebitda'],['tax_paid','operator_fees'],['dividend','investor_outlay'],['revenue','ebitda'],['cash','debt']]
 labels={'capex':('CAPEX','CAPEX'),'equity':('Capital KK','KK equity'),'receipts':('Recebimentos','Collections'),'cash':('Caixa','Cash'),'debt':('Dívida','Debt'),'reserve':('Reservas','Reserves'),'deed_sales':('Escrituras','Deed sales'),'res_revenue':('Receitas residenciais','Residential revenue'),'owner_distribution':('Rendas suportáveis','Supported owner income'),'revenue':('Receita','Revenue'),'ebitda':('EBITDA','EBITDA'),'tax_paid':('IRPC','Corporate tax'),'operator_fees':('Fees','Fees'),'dividend':('Dividendos','Dividends'),'investor_outlay':('Capital investidores','Investor equity')}
 bindings=[]
 for sn,pages in [('DASH 10 V06',10),('V06 Brief',2)]:
  s=wb.add_worksheet(sn);s.hide_gridlines(2);s.hide_row_col_headers();s.set_column('A:L',7);s.set_column('M:P',18);s.set_default_row(12);s.set_portrait();s.set_paper(9);s.set_print_scale(100);s.set_margins(.3,.3,.3,.3);s.print_area(0,0,pages*60-1,11);s.set_h_pagebreaks([60*i for i in range(1,pages)]);s.freeze_panes(7,0);s.set_zoom(100);s.set_tab_color('142337');s.set_footer('&LKK-01 · EFCC · Simulated scenario / Cenário simulado&R&P / &N')
  for page in range(pages):
   off=page*60;section=page if sn=='DASH 10 V06' else (0 if page==0 else 9)
   def area(c1,r,c2,r2):return f'{c1}{off+r}:{c2}{off+r2}'
   s.merge_range(area('A',1,'I',3),'KK-01 | HOTEL & RESIDENCES',F['header']);link(s,area('J',1,'L',3),'V06 Identidade','B8','Logótipo','Logo');text(s,area('A',4,'L',6),f'{section+1:02} — {SECTIONS[section][0]}',f'{section+1:02} — {SECTIONS[section][1]}','section')
   link(s,area('M',2,'M',3),'EXEC V06','A1','Início','Home');link(s,area('N',2,'N',3),CENTER,'A1','Voltar','Back');link(s,area('O',2,'O',3),CENTER,'H11','Parâmetros','Assumptions');link(s,area('P',2,'P',3),'V06 Imprimir','A4','Imprimir / PDF','Print / PDF')
   put(s,area('A',8,'L',8),f"'V06 Hipoteses'!D29&\" · \"&'V06 Hipoteses'!D27&\" · \"&'V06 Hipoteses'!D28",'EFCC-20261022-01 · Active Scenario · 2026-10-22','sub')
   cards,records,notes=structured(sn,page)
   if len(cards)>4:
    selected=[0,1,2,4] if section==0 else [0,1,2,4]
    spare=[x for i,x in enumerate(cards) if i not in selected];cards=[cards[i] for i in selected];records=[(a,b,'ratio' if kind=='ratio' else 'percent' if kind=='pct' else 'num') for a,b,kind in spare]+records
   if sn=='DASH 10 V06' and section==0:
    reference_row=next(int(a[1:]) for a in src.cells('V06 Cenarios') if a.startswith('A') and src.get('V06 Cenarios',a)[1]=='Reference Case')
    def selected_metric(column,key):
     return f'IF({case}="Reference Case",\'V06 Cenarios\'!{column}{reference_row},\'V06 Cenarios\'!{column}5)',ref['metrics']['Base'][key]
    reserve_records=records[:2]
    records=[((tr('VAL do projeto','Project NPV'),'VAL do projeto'),selected_metric('T','pnpv'),'num')]+reserve_records+[((tr('Capital KK previsto','Planned KK equity'),'Capital KK previsto'),(f'IF({case}="Reference Case",SUM(\'V06 Referencia\'!M5:M184),SUM(\'V06 Motor\'!M5:M184))',sum(r['equity'] for r in rows)),'num'),((tr('Pico de dívida','Peak debt'),'Pico de dívida'),selected_metric('F','debt'),'num')]+[r for r in records[2:] if r[2]=='percent']
   if sn=='DASH 10 V06' and section==1:
    c=col(cc['capex']);records.append(((tr('CAPEX após M30','CAPEX after M30'),'CAPEX após M30'),(f'IF({case}=\"Reference Case\",SUM(\'V06 Referencia\'!{c}35:{c}184),SUM(\'V06 Motor\'!{c}35:{c}184))',sum(r['capex'] for r in rows[30:])),'num'))
   if sn=='DASH 10 V06' and section==2:
    def monthmetric(key,m):
     c=col(cc[key]);return f'IF({case}="Reference Case",\'V06 Referencia\'!{c}{m+4},\'V06 Motor\'!{c}{m+4})',rows[m-1][key]
    def sumvalue(key):
     c=col(cc[key]);return f'IF({case}="Reference Case",SUM(\'V06 Referencia\'!{c}5:{c}184),SUM(\'V06 Motor\'!{c}5:{c}184))',sum(r[key] for r in rows)
    records=[((tr('Reserva protegida M42','Protected reserves M42'),'Reserva protegida M42'),monthmetric('reserve',42),'num'),((tr('Dívida M180','M180 debt'),'Dívida M180'),monthmetric('debt',180),'num'),((tr('DSCR Hotel M42 · não covenant','Hotel DSCR M42 · not a covenant'),'DSCR Hotel M42 · não covenant'),monthmetric('dscr',42),'ratio'),((tr('Cash sweep total','Total cash sweep'),'Cash sweep total'),sumvalue('sweep'),'num')]+[r for r in records if any(k in str(r[0][1]).lower() for k in ['cobertura','juros'])]
   if sn=='DASH 10 V06' and section==5:
    year=f'IF({case}="Reference Case",5,\'DASH 05 V06\'!$N$23)';selected=[r for r in rows if r['op_age']==4 and r['available_room_nights']>0]
    def annual_metric(key):
     c=col(cc[key]);active=f'SUMIFS(\'V06 Motor\'!{c}5:{c}184,\'V06 Motor\'!E5:E184,({year})-1,\'V06 Motor\'!DL5:DL184,">0")';frozen=active.replace('V06 Motor','V06 Referencia');return f'IF({case}="Reference Case",{frozen},{active})',sum(r[key] for r in selected)
    def ratio_metric(a,b):
     af,av=a;bf,bv=b;return f'IF(({bf})>0,({af})/({bf}),0)',av/bv if bv else 0
    actcount=f'COUNTIFS(\'V06 Motor\'!E5:E184,({year})-1,\'V06 Motor\'!DL5:DL184,">0")';frozen=actcount.replace('V06 Motor','V06 Referencia');count=(f'IF({case}="Reference Case",{frozen},{actcount})',len(selected))
    cards=[((tr('ADR · ano selecionado','ADR · selected year'),'ADR · ano selecionado'),ratio_metric(annual_metric('adr'),count),'money'),((tr('Ocupação · ano selecionado','Occupancy · selected year'),'Ocupação · ano selecionado'),ratio_metric(annual_metric('occupancy'),count),'pct'),((tr('RevPAR · ano selecionado','RevPAR · selected year'),'RevPAR · ano selecionado'),ratio_metric(annual_metric('room_revenue'),annual_metric('available_room_nights')),'money'),((tr('EBITDA · ano selecionado','EBITDA · selected year'),'EBITDA · ano selecionado'),annual_metric('ebitda'),'money')]
    records=[((tr('Índice de utilização MICE','MICE utilisation index'),'Índice de utilização MICE'),(f'IF({case}="Reference Case",0.5,\'DASH 05 V06\'!N25)',.5),'percent')]
    for key,a,b in [('revenue','Receita total · ano','Total revenue · year'),('gop','GOP · ano','GOP · year')]:records.append(((tr(a,b),a),annual_metric(key),'num'))
    records.append(((tr('Margem EBITDA · ano','EBITDA margin · year'),'Margem EBITDA · ano'),ratio_metric(annual_metric('ebitda'),annual_metric('revenue')),'percent'))
    for key,a,b in [('operator_fees','Fees operadora · ano','Operator fees · year'),('ffe','Dotação FF&E · ano','FF&E contribution · year')]:records.append(((tr(a,b),a),annual_metric(key),'num'))
   if sn=='DASH 10 V06' and section==6:
    cards[0]=((tr('IRPC nominal de estudo','Nominal study corporate tax'),'IRPC nominal de estudo'),(f'IF({case}="Reference Case",0.2,\'V06 Controlos\'!B9)',.2),'pct');cards[1]=((tr('IVA de estudo','Study VAT rate'),'IVA de estudo'),(f'IF({case}="Reference Case",0.15,\'V06 Controlos\'!B6)',.15),'pct')
    feescol=col(cc['res_fees']);revcol=col(cc['res_revenue']);resratio=f'IF(SUM(\'V06 Motor\'!{revcol}5:{revcol}184)>0,SUM(\'V06 Motor\'!{feescol}5:{feescol}184)/SUM(\'V06 Motor\'!{revcol}5:{revcol}184),0)';frozen=resratio.replace('V06 Motor','V06 Referencia');grcol=col(cc['general_reserve'])
    records=[((tr('Reembolso IVA · meses','VAT refund · months'),'Reembolso IVA · meses'),(f'IF({case}="Reference Case",6,\'V06 Controlos\'!B8)',6),'number')]+records[:3]+[((tr('Fees residenciais / receita','Residential fees / revenue'),'Fees residenciais / receita'),(f'IF({case}="Reference Case",{frozen},{resratio})',sum(r['res_fees'] for r in rows)/sum(r['res_revenue'] for r in rows)),'percent'),((tr('Reserva geral ativa','Active general reserve'),'Reserva geral ativa'),(f'IF({case}="Reference Case",\'V06 Referencia\'!{grcol}5,\'V06 Motor\'!{grcol}5)',rows[0]['general_reserve']),'num')]
   if section==9:
    reference_row=next(int(a[1:]) for a in src.cells('V06 Cenarios') if a.startswith('A') and src.get('V06 Cenarios',a)[1]=='Reference Case')
    def decision_metric(column,key):return f'IF({case}="Reference Case",\'V06 Cenarios\'!{column}{reference_row},\'V06 Cenarios\'!{column}5)',ref['metrics']['Base'][key]
    cards=[((tr(a,b),a),decision_metric(c,k),kind) for a,b,c,k,kind in [('VAL novos investidores','New investor NPV','M','npv','money'),('TIR novos investidores','New investor IRR','K','irr','pct'),('MOIC novos investidores','New investor MOIC','L','moic','ratio'),('VAL do projeto','Project NPV','T','pnpv','money')]]
    records=[]
    records=[((tr('Promotor · VAL do projeto','Developer · project NPV'),'Promotor · VAL do projeto'),decision_metric('T','pnpv'),'num'),((tr('Investidores · VAL equity','Investors · equity NPV'),'Investidores · VAL equity'),decision_metric('M','npv'),'num')]
    for key,a,b,stock in [('funding_gap','Banca · défice acumulado','Lenders · cumulative funding gap',False),('owner_shortfall','Compradores · yield alvo não coberto','Buyers · uncovered target income',False),('operator_fees','Operadora · fees no horizonte','Operator · fees over horizon',False),('reserve','Reservas protegidas · fim','Protected reserves · end',True)]:
     c=col(cc[key]);active=f"'V06 Motor'!{c}184" if stock else f"SUM('V06 Motor'!{c}5:{c}184)";frozen=active.replace('V06 Motor','V06 Referencia');value=rows[-1][key] if stock else sum(r[key] for r in rows);records.append(((tr(a,b),a),(f'IF({case}="Reference Case",{frozen},{active})',value),'num'))

   if not cards:
    cards=[(a,b,'money') for a,b,_ in records[:4]];records=records[4:]
   for n,(lab,value,kind) in enumerate(cards[:4]):
    r=10+(n//2)*5;c='A' if n%2==0 else 'G';end='F' if n%2==0 else 'L';put(s,area(c,r,end,r+1),*lab,'label');put(s,area(c,r+2,end,r+3),*value,kind);bindings.append({'sheet':sn,'cell':f'{c}{off+r+2}','formula':value[0],'value':value[1]})
   if section==1:text(s,area('A',20,'L',20),'Investimento e recebimentos · trimestres 1–10 (M1–M30)','Investment and collections · quarters 1–10 (M1–M30)','sub')
   if section==8:
    # Two native NPV heatmaps from the existing 30 investor cash-flow series.
    # Fixed zero midpoint: negative NPV never receives a green viability colour.
    for panel,rate in enumerate([.10,.12]):
     r0=21+panel*10;text(s,area('A',r0,'L',r0),f'VAL novos investidores · desconto {rate:.0%} · M€',f'New investor NPV · discount {rate:.0%} · EUR m','section')
     for j in range(5):
      _,adr=src.get('V06 Matriz ADR',f'{col(j+2)}6');s.write_formula(off+r0,3+j,f"='V06 Matriz ADR'!{col(j+2)}6",F['thead'],adr)
     for i in range(6):
      _,occ=src.get('V06 Matriz ADR',f'B{7+i}');s.write_formula(off+r0+1+i,2,f"='V06 Matriz ADR'!B{7+i}",F['percent'],occ)
      for j in range(5):
       flow=f'V06 F{26+i*5+j:02}';cash=[src.get(flow,f'B{m+5}')[1] for m in range(181)];v=sum(c/(1+rate)**(m/12) for m,c in enumerate(cash));f=f"'{flow}'!B5+NPV((1+{rate})^(1/12)-1,'{flow}'!B6:B185)";s.write_formula(off+r0+1+i,3+j,'='+f,F['heat'],v);bindings.append({'sheet':sn,'cell':f'{col(3+j)}{off+r0+2+i}','formula':f,'value':v})
     s.conditional_format(off+r0+1,3,off+r0+6,7,{'type':'3_color_scale','min_type':'num','min_value':-25000000,'mid_type':'num','mid_value':0,'max_type':'num','max_value':25000000,'min_color':'F2C8CC','mid_color':'FFFFFF','max_color':'C8E5D7'})

   else:plot(s,f'A{off+21}',[(k,*labels[k]) for k in pagekeys[section]],width=625,height=172)
   text(s,area('A',39 if section==8 else 35,'H',39 if section==8 else 35),'Indicador / base de análise','Indicator / analysis basis','thead');text(s,area('I',39 if section==8 else 35,'L',39 if section==8 else 35),'Valor','Value','thead')
   for n,(lab,value,kind) in enumerate(records[:6]):
    r=(40 if section==8 else 36)+n*2;put(s,area('A',r,'H',r+1),*lab,'text');put(s,area('I',r,'L',r+1),*value,kind);bindings.append({'sheet':sn,'cell':f'I{off+r}','formula':value[0],'value':value[1]})
   npvf,npvv=src.get('DASH 10 V06','A26')
   finding=f'IF(({npvf})<0,{tr("Conclusão: VAL dos novos investidores negativo à taxa exigida.","Finding: new investor NPV is negative at the required return.")},{tr("Conclusão: VAL não negativo no cenário; validar condições e contratos.","Finding: non-negative scenario NPV; validate conditions and contracts.")})'
   put(s,area('A',49,'L',50),finding,'Conclusão: VAL dos novos investidores negativo à taxa exigida.','alert')
   risks=[('Riscos: contratos, captação de capital, financiamento e procura por comprovar.','Risks: contracts, equity funding, debt and market demand require evidence.'),('Riscos: sobrecustos, prazos de construção e recebimentos documentais.','Risks: overruns, construction timing and documented collections.'),('Riscos: funding gaps, reservas protegidas e covenants não aprovados.','Risks: funding gaps, protected reserves and unapproved covenants.'),('Riscos: vendas previstas não são contratos ou recebimentos documentados.','Risks: forecast sales are not documented contracts or collections.'),('Riscos: yield alvo não garantido; rendas limitadas à caixa do pool.','Risks: target yield is not guaranteed; owner income is limited by pool cash.'),('Riscos: volumes MICE não documentados; fees são condições de estudo.','Risks: MICE event volumes lack evidence; operator terms are study assumptions.'),('Riscos: incentivos não aprovados; HOTUSA é reserva jurídica separada.','Risks: incentives are unapproved; HOTUSA is a separate legal reserve.'),('Riscos: dividendos condicionados por dívida, lucros, reservas e cobertura.','Risks: dividends depend on debt, profits, reserves and coverage.'),('Riscos: grelha discreta; sem extrapolar equilíbrio. Sensibilidades são do cenário ativo.','Risks: discrete grid; no extrapolated break-even. Sensitivities use the active scenario.'),('Riscos: não concluir viabilidade incondicional antes de validar contratos e mercado.','Risks: do not conclude unconditional feasibility before validating contracts and market.')]
   if section==9:
    decisions=[('Promotor: confirmar CAPEX, prazos e fontes de fundos antes de compromissos.','Developer: confirm CAPEX, schedule and funding before commitments.'),('Investidores: confrontar VAL/TIR com retorno exigido; negociar condições antes de investir.','Investors: compare NPV/IRR with required return; negotiate terms before investing.'),('Banca: validar garantias e covenants; DSCR Hotel não é covenant aprovado.','Lenders: validate security and covenants; Hotel DSCR is not an approved covenant.'),('Compradores: rendimento projetado depende da caixa do pool; não é renda garantida.','Buyers: projected income depends on pool cash; it is not guaranteed rent.'),('Operadora: confirmar fees, obrigações e reservas no contrato de exploração.','Operator: confirm fees, obligations and reserves in the operating contract.')]
    for j,(a,b) in enumerate(decisions):text(s,area('A',51+j,'L',51+j),a,b,'note')
   else:
    text(s,area('A',51,'L',53),*risks[section],'note')
    text(s,area('A',54,'L',55),'Implicação: validar as condições críticas e negociar estrutura de investimento antes da decisão.','Implication: validate critical conditions and negotiate investment terms before deciding.','note')
   text(s,area('A',57,'L',58),'Cenário simulado · fontes: motor financeiro V06 · EUR · validar pressupostos e documentação.','Simulated scenario · source: V06 financial engine · EUR · validate assumptions and evidence.','note');text(s,area('A',59,'L',59),f'KK-01 Executive Financial Control Center · V06 · página {page+1}/{pages}',f'KK-01 Executive Financial Control Center · V06 · page {page+1}/{pages}','sub')
 return bindings
