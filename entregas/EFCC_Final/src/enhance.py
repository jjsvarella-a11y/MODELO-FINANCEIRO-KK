"""EFCC extension to the V06 generator: live controls, shared 180-month scenarios, UI.
All cases use core.Model.calculate, not approximations or duplicate financial rules.
"""
from core import *
from datetime import date
DASHES=[('EXEC V06','Dashboard Executivo','Executive Dashboard'),('DASH 01 V06','Investimento e Orçamento de Construção','Investment & Construction Budget'),('DASH 02 V06','Fiscalidade, Incentivos e Seguros','Taxation, Incentives & Insurance'),('DASH 03 V06','Desempenho Financeiro Consolidado','Consolidated Financial Performance'),('DASH 04A V06','Promoção e Comercialização Imobiliária','Real Estate Development & Sales'),('DASH 04B V06','Exploração Residencial e Sell & Leaseback','Residential Operations & Sell & Leaseback'),('DASH 05 V06','Exploração Hoteleira e Centro MICE','Hotel & MICE Operations'),('DASH 06 V06','Estrutura Acionista e Blocos de Investimento','Shareholding & Investment Blocks'),('DASH 07 V06','Rentabilidade dos Acionistas e Distribuições','Equity Returns & Distributions'),('DASH 10 V06','Relatório Executivo Consolidado de Viabilidade e Investimento','Consolidated Investment & Feasibility Report')]
STATE={}
def initialise(wb,fm,model,p,sheets):
 st=STATE;st.update(wb=wb,fm=fm,model=model,raw=p.copy(),sheets=sheets)
 controls=model.sheet('V06 Hipoteses',['Key','Parâmetro PT','Parameter EN','Valor / Value','Unidade','Base / Limitação'])
 hp=legacy.config();items=[
 ('hotel_mode','Cenário operacional','Operating case','Base Case','list','Base Case retains existing V06 hotel scenario'),
 ('adr_input','ADR Custom','Custom ADR',legacy.HOTELS[p['hotel_scenario'].v][0],'EUR','120–250; 5 EUR steps; custom case only'),
 ('occ_input','Ocupação Custom','Custom occupancy',legacy.HOTELS[p['hotel_scenario'].v][1],'%','45–78%; 1 percentage point steps; existing ramp preserved'),
 ('mice_util','Índice utilização MICE','MICE utilisation index',.5,'%','20–80%; 5pp steps; 50% normalisation only, no evidenced event capacity'),
 ('ffe_rate','Dotação FF&E','FF&E contribution',hp['ffe'],'%','2–6%; .5pp steps; excludes initial FF&E in CAPEX'),
 ('adr_growth','Crescimento ADR','ADR growth',hp['adr_growth'],'%','0–6%; .5pp steps; operational anniversaries'),
 ('hotel_fees','Perfil fees Hotel','Hotel fee profile','Custom Terms','list','Study terms, not contractual proposals; custom preserves 3/3/6'),
 ('marketing','Marketing Custom','Custom marketing',0,'% revenue','Additive only within aggregate revenue-based fees'),
 ('distribution','Distribuição Custom','Custom distribution',0,'% revenue','No additional duplication of existing systems fee'),
 ('owner_priority_on','Ativar prioridade proprietário','Enable owner priority',0,'0/1','Off; scenario only, no adopted contract'),
 ('owner_priority','Prioridade anual EUR','Annual owner priority',0,'EUR/year','Deducted from positive adjusted GOP before incentive'),
 ('res_profile','Perfil fees residencial','Residential fee profile','Custom Terms','list','Independent from hotel'),
 ('res_base_custom','Gestão residencial Custom','Custom residential management',hp['res_base'],'% revenue','Existing V06 3%'),
 ('res_system_custom','Agregado residencial Custom','Custom residential aggregate',hp['res_system'],'% revenue','Existing V06 5%; no duplicate hotel fees'),
 ('reserve_policy','Política reserva geral','General reserve policy','Flexible','list','Documented V06 baseline 2M; Conservative 3M, Balanced 2.5M'),
 ('debt_reserve','Reserva contratual dívida','Debt contractual reserve',0,'EUR','No contracted amount supplied; explicit exposure'),
 ('other_reserve','Outras reservas contratuais','Other contractual reserves',0,'EUR','No documented contractual schedule'),
 ('protect_known_tax','Proteger impostos vencidos','Protect known tax payable',1,'0/1','Previous closed-year tax payable less current payment; no circularity'),
 ('ffe_spend','Reposição / dotação FF&E','Replacement / FF&E allocation',1,'% contribution','100% legacy forecast; not evidenced actuals; unspent fund excluded from exit'),
 ('matrix_metric','Indicador matriz','Matrix metric','NPV','list','NPV / IRR / EBITDA / RevPAR / Exit Equity Value'),
 ('discount_mode','Taxa desconto','Discount rate mode','12%','list','8% / 10% / 12% / Custom; central effective rate'),
 ('report_mode','Formato relatório','Report format','Complete Report','list','Complete Report 10 pages / Executive Brief 2 pages'),
 ('report_case','Base do relatório','Report case','Active Scenario','list','Reference Case uses frozen baseline parameters and source inventory'),
 ('report_date','Data referência','Reference date','2026-10-22','ISO date','Reference date; not evidence that recalculation completed'),
 ('scenario_id','Identificação cenário','Scenario ID','EFCC-20261022-01','text','Record with parameters for reproducibility'),
 ('status_month','Mês disponibilidade','Availability month',30,'month','Commercial forecast cutoff; documented status separate'),
 ('stabilized_year','Ano operacional estabilizado','Stabilized operating year',5,'year','Measured from hotel opening; 1–12 to fit horizon'),
 ('selected_adr','ADR comparação matriz','Selected matrix ADR',200,'EUR','150 / 175 / 200 / 225 / 250'),
 ('selected_occ','Ocupação comparação matriz','Selected matrix occupancy',.65,'%','50 / 55 / 60 / 65 / 70 / 75%')]
 lists={'hotel_mode':['Conservative','Base Case','Favorable','Upside Stress','Custom'],'hotel_fees':['Operator Proposed Terms','KK Target Terms','KK Optimized Terms','Custom Terms'],'res_profile':['Reference Study','KK Target Terms','KK Optimized Terms','Custom Terms'],'reserve_policy':['Conservative','Balanced','Flexible'],'matrix_metric':['NPV','IRR','EBITDA','RevPAR','Exit Equity Value'],'discount_mode':['8%','10%','12%','Custom'],'report_mode':['Complete Report','Executive Brief'],'report_case':['Active Scenario','Reference Case']}
 fields={};fieldrows={}
 # Inputs are visible in contextual panels. Registry mirrors them, rather than
 # presenting a second competing editable copy.
 panel={'hotel_mode':('DASH 05 V06','H23'),'adr_input':('DASH 05 V06','H25'),'occ_input':('DASH 05 V06','H27'),'mice_util':('DASH 05 V06','N25'),'ffe_rate':('DASH 05 V06','N27'),'adr_growth':('DASH 05 V06','N29'),'hotel_fees':('DASH 05 V06','H32'),'marketing':('DASH 05 V06','H34'),'distribution':('DASH 05 V06','N34'),'owner_priority_on':('DASH 05 V06','H36'),'owner_priority':('DASH 05 V06','N36'),'res_profile':('DASH 04B V06','H23'),'res_base_custom':('DASH 04B V06','H25'),'res_system_custom':('DASH 04B V06','N25'),'reserve_policy':('DASH 07 V06','H23'),'matrix_metric':('V06 Matriz ADR','B3'),'discount_mode':('V06 Matriz ADR','F3'),'stabilized_year':('DASH 05 V06','N23')}
 st.update(items=items,lists=lists,fields=fields,fieldrows=fieldrows,panel=panel)
 for r,(key,pt,en,val,unit,note) in enumerate(items,5):
  fieldrows[key]=r;controls.write_row(r-1,0,[key,pt,en]);controls.write(r-1,4,unit);controls.write(r-1,5,note)
  if key in panel:
   sn,addr=panel[key];x=X(f"'{sn}'!{addr}",val);model.cell(controls,r,3,x,'text' if isinstance(val,str) else 'pct' if unit.startswith('%') else 'num')
  else:
   controls.write(r-1,3,val,fm['input']);x=X(f"'V06 Hipoteses'!$D${r}",val)
   if key in lists:controls.data_validation(r-1,3,r-1,3,{'validate':'list','source':lists[key]})
  fields[key]=x
 controls.set_column('A:C',30);controls.set_column('D:D',27);controls.set_column('E:E',16);controls.set_column('F:F',85)
 p.update({k:fields[k] for k in ['hotel_mode','adr_input','occ_input','mice_util','ffe_rate','adr_growth','owner_priority_on','owner_priority','debt_reserve','other_reserve','protect_known_tax','ffe_spend']})
 profile=fields['hotel_fees']
 p['base_fee']=IF(profile.eq('Operator Proposed Terms'),.04,IF(profile.eq('KK Target Terms'),.025,IF(profile.eq('KK Optimized Terms'),.02,p['base_fee'])))
 p['incentive']=IF(profile.eq('Operator Proposed Terms'),.10,IF(profile.eq('KK Target Terms'),.06,IF(profile.eq('KK Optimized Terms'),.05,p['incentive'])))
 p['systems']=IF(profile.eq('Operator Proposed Terms'),.04,IF(profile.eq('KK Target Terms'),.03,IF(profile.eq('KK Optimized Terms'),.015,p['systems']+fields['marketing']+fields['distribution'])))
 rp=fields['res_profile'];p['res_base']=IF(rp.eq('Reference Study'),.08,IF(rp.eq('KK Target Terms'),.06,IF(rp.eq('KK Optimized Terms'),.05,fields['res_base_custom'])))
 p['res_system']=IF(rp.eq('Reference Study'),.03,IF(rp.eq('KK Target Terms'),.02,IF(rp.eq('KK Optimized Terms'),.015,fields['res_system_custom'])))
 rp=fields['reserve_policy'];p['reserve']=IF(rp.eq('Conservative'),3000000,IF(rp.eq('Balanced'),2500000,2000000))
 dm=fields['discount_mode'];p['discount']=IF(dm.eq('8%'),.08,IF(dm.eq('10%'),.1,IF(dm.eq('12%'),.12,p['discount'])))
 st['effective']=p
 # Actual inputs are kept separate; absence is not a zero receipt.
 ledger=model.sheet('V06 Cobrancas',['Unit ID','Month actual','Received EUR actual','Payment type','Document reference','Date ISO'])
 ledger.autofilter(3,0,503,5);ledger.set_column('A:F',24)
 for r in range(4,504):
  for c in range(6):ledger.write_blank(r,c,None,fm['input'])
 ledger.write('H5','Actual receipts register; empty means no documentary evidence. Forecast engine remains labelled forecast.')
 # Frozen reference inventory is a snapshot for reproducibility, not an editable database.
 refunits=model.sheet('V06 Ref Unidades',['ID','Price EUR','Sale month','Deed month','Snapshot'])
 refdataset=[]
 for r,u in enumerate(model.units,5):
  nu=u.copy()
  for c,key in enumerate(['id','price','sale','deed']):refunits.write(r-1,c,u[key].v);nu[key]=X(f"'V06 Ref Unidades'!${col(c)}${r}",u[key].v)
  refunits.write(r-1,4,'Frozen C0 V06 inventory');refdataset.append(nu)
 st['reference_units']=refdataset

def scenarios(model,scenario_defs,scenario_models):
 st=STATE
 # Full 180-month branches, one core.calculate factory for every matrix cell.
 cases=[]
 for occ in [.50,.55,.60,.65,.70,.75]:
  for adr in [150,175,200,225,250]:
   i=len(cases)+1;label=f'ADR {adr} Occ {occ:.0%}';over={'hotel_mode':'Custom','adr_input':adr,'occ_input':occ}
   # Matrix coordinates are live inputs, not constants used only at build time.
   over['adr_input']=X(f"'V06 Matriz ADR'!{col(2+(i-1)%5)}6",adr);over['occ_input']=X(f"'V06 Matriz ADR'!B{7+(i-1)//5}",occ)
   scenario_defs.append((label,over));scenario_models.append(model.calculate(f'V06 M{i:02}',over));cases.append((label,adr,occ))
 st['matrix_cases']=cases
 extra=[('Hotel Proposed',{'base_fee':.04,'systems':.04,'incentive':.10}),('Hotel KK Target',{'base_fee':.025,'systems':.03,'incentive':.06}),('Hotel KK Optimized',{'base_fee':.02,'systems':.015,'incentive':.05}),('Residential Study',{'res_base':.08,'res_system':.03}),('Residential KK Target',{'res_base':.06,'res_system':.02}),('Residential KK Optimized',{'res_base':.05,'res_system':.015}),('Reserve Conservative',{'reserve':3000000}),('Reserve Balanced',{'reserve':2500000}),('Reserve Flexible',{'reserve':2000000}),('Debt Grace 30',{'grace':30}),('Debt Grace 54',{'grace':54}),('Debt Tenor 180',{'term':180}),('Debt Sweep 50%',{'sweep':.5})]
 for i,(label,ov) in enumerate(extra,1):scenario_defs.append((label,ov));scenario_models.append(model.calculate(f'V06 O{i:02}',ov))
 # Reference case: frozen original controls, inventory and subscription calendar.
 raw=st['raw'];refparams={k:asx(v.v) for k,v in raw.items()};refparams.update({'protect_known_tax':0,'hotel_mode':'Base Case','adr_input':legacy.HOTELS['Favorável'][0],'occ_input':legacy.HOTELS['Favorável'][1],'adr_growth':.03,'mice_util':.5,'ffe_rate':.05,'ffe_spend':1,'owner_priority_on':0,'owner_priority':0,'res_base':.03,'res_system':.05,'debt_reserve':0,'other_reserve':0})
 original_units=model.units;original_tickets=model.tickets;model.units=st['reference_units'];model.tickets=[tuple(asx(v.v) if isinstance(v,X) else v for v in t) for t in original_tickets]
 # Calendar helper signature must include inventory identity for the frozen snapshot.
 ref=model.calculate('V06 Referencia',refparams);model.units=original_units;model.tickets=original_tickets
 scenario_defs.append(('Reference Case',refparams));scenario_models.append(ref);st['reference_model']=ref
 return cases

def complete(wb,fm,model,params,main,metrics,scenario_models,scenario_defs,sheets,card_defs,translations,tr,dictrows):
 st=STATE;st.update(main=main,metrics=metrics,scenario_models=scenario_models,scenario_defs=scenario_defs,tr=tr,translations=translations)
 st['tr_count_start']=translations.dim_rowmax+1
 st['lang']=params['language'].f
 st['titles']=[tr('EFCC.nav.'+str(i),f'{i:02} — '+pt,f'{i:02} — '+en) for i,(_,pt,en) in enumerate(DASHES,1)]
 st['fm2']={}
 styles={'header':{'bg_color':'B31F2B','font_color':'FFFFFF','bold':True,'font_size':18,'valign':'vcenter'},'brand':{'bg_color':'B31F2B','font_color':'FFFFFF','bold':True,'font_size':13,'text_wrap':True,'valign':'vcenter'},'inset':{'bg_color':'FFFFFF'},'language':{'bg_color':'E8EDF3','font_color':'203246','bold':True,'font_size':10},'side':{'bg_color':'142331','font_color':'FFFFFF','text_wrap':True,'font_size':9,'valign':'vcenter'},'active':{'bg_color':'B31F2B','font_color':'FFFFFF','text_wrap':True,'font_size':9,'bold':True,'valign':'vcenter'},'panel':{'font_color':'142B45','bg_color':'EAF0F5','bold':True,'font_size':12},'money':{'font_size':22,'bold':True,'bg_color':'F2F5F8','num_format':'#,##0 "€";[Red](#,##0) "€";–'},'count':{'font_size':22,'bold':True,'bg_color':'F2F5F8','num_format':'#,##0'},'pct':{'font_size':22,'bold':True,'bg_color':'F2F5F8','num_format':'0.00%'},'ratio':{'font_size':22,'bold':True,'bg_color':'F2F5F8','num_format':'0.00"×"'},'label':{'font_color':'43576B','bg_color':'F2F5F8','text_wrap':True,'font_size':10,'valign':'vcenter'},'input':{'font_color':'1565A0','bg_color':'E4EEF8','bold':True,'num_format':'#,##0.00','border':1,'border_color':'C8D6E3'},'inputpct':{'font_color':'1565A0','bg_color':'E4EEF8','bold':True,'num_format':'0.0%','border':1,'border_color':'C8D6E3'},'inputtext':{'font_color':'1565A0','bg_color':'E4EEF8','bold':True,'border':1,'border_color':'C8D6E3','font_size':9},'small':{'font_size':9,'font_color':'596D80','text_wrap':True,'valign':'top'},'text':{'font_size':10,'font_color':'203246','text_wrap':True},'num':{'font_size':10,'num_format':'#,##0.00;[Red](#,##0.00)'},'percent':{'font_size':10,'num_format':'0.00%'},'head':{'bg_color':'344C62','font_color':'FFFFFF','bold':True,'text_wrap':True,'font_size':9},'alert':{'font_color':'B31F2B','font_size':10,'text_wrap':True,'bold':True}}
 for k,v in styles.items():st['fm2'][k]=wb.add_format({'font_name':'Aptos','valign':'vcenter',**v})
 ff=st['fm2'];rows=main['rows'];cols=main['columns'];lang=st['lang'];fields=st['fields']
 def text(ws,addr,pt,en,fmt='text'):
  ref=tr('EFCC.'+ws.name+'.'+addr,pt,en);ws.write_formula(addr,'='+ref,ff[fmt],pt)
 def merge_text(ws,area,pt,en,fmt='text'):
  ws.merge_range(area,'',ff[fmt]);text(ws,area.split(':')[0],pt,en,fmt)
 def value(ws,addr,x,fmt='num'):model.cell(ws,int(''.join(filter(str.isdigit,addr))),xlsxwriter.utility.xl_cell_to_rowcol(addr)[1],x,fmt if fmt in fm else 'num')
 for di,(sn,pt,en) in enumerate(DASHES[:9]):
  s=sheets[sn]
  # Redraw only presentation; KPI addresses used by DASH10 remain identical.
  s.table.clear();s.merge.clear();s.merged_cells.clear();s.charts.clear();s.images.clear();s.hyperlinks.clear()
  s.hide_gridlines(2);s.hide_row_col_headers();s.set_zoom(80);s.freeze_panes(5,4);s.set_default_row(19);s.set_column('A:A',1);s.set_column('B:C',16);s.set_column('D:D',1.5);s.set_column('E:P',11);s.set_column('Q:Q',1)
  for r in range(64):
   for c in range(4):s.write_blank(r,c,None,ff['side'])
  s.merge_range('B2:C4','JS-CV\nCONSTRUÇÕES E\nINVESTIMENTOS, LDA.',ff['brand'])
  s.merge_range('E2:M4','KK-01  |  HOTEL & RESIDENCES',ff['header']);s.merge_range('N2:P4','',ff['inset'])
  import os
  logo=Path(os.environ.get('KK01_LOGO','/workspace/shared/logo jscv(1).jpg'))
  if logo.exists():
   from PIL import Image
   with Image.open(logo) as im:w,h=im.size
   scale=min(205/w,70/h);s.insert_image('N2',str(logo),{'x_scale':scale,'y_scale':scale,'x_offset':3,'y_offset':2,'object_position':1});st['logo_sha256']=__import__('hashlib').sha256(logo.read_bytes()).hexdigest()
  merge_text(s,'E5:G5','Idioma global','Global language','language');s.merge_range('H5:J5','',ff['language']);s.write_formula('H5','=HYPERLINK("#\'V06 Controlos\'!B5",'+lang+'&"  ▾")',ff['language'],'PT  ▾')
  s.merge_range('K5:M5','',ff['language']);s.write_url('K5',"internal:'EXEC V06'!A1",ff['language'],string='⌂ Início / Home')
  s.merge_range('N5:P5','',ff['language']);s.write_formula('N5','=HYPERLINK("#\'V06 Imprimir\'!A1",IF('+lang+'="PT","Imprimir / PDF","Print / PDF"))',ff['language'],'Imprimir / PDF')
  for ni,(target,_,_) in enumerate(DASHES):
   rr=7+ni*5;s.merge_range(rr-1,1,rr+3,2,'',ff['active'] if ni==di else ff['side']);s.write_formula(rr-1,1,'=HYPERLINK("#\''+target+'\'!A1",'+st['titles'][ni]+')',ff['active'] if ni==di else ff['side'],f'{ni+1:02} — '+DASHES[ni][1])
  # Six live KPI cards, monetary units applied by meaning rather than all cells.
  for ci,(lpt,len_,x,dkey,oldfmt) in enumerate(card_defs[sn]):
   rr=7+(ci//3)*6;cc=4+(ci%3)*4;label=tr('EFCC.card.'+sn+'.'+str(ci),lpt,len_)
   s.merge_range(rr-1,cc,rr,cc+2,'',ff['label']);s.write_formula(rr-1,cc,'='+label,ff['label'],lpt);s.write_url(rr-1,cc+3,f"internal:'V06 Dicionario'!A{dictrows[dkey]}",ff['label'],string='ⓘ')
   ft='pct' if oldfmt=='cardpct' else 'ratio' if oldfmt=='cardx' else 'count' if (di==4 and ci==0 or di==7 and ci in [0,1]) else 'money'
   s.merge_range(rr+1,cc,rr+3,cc+3,'',ff[ft]);s.write_formula(rr+1,cc,'='+x.f,ff[ft],x.v)
  merge_text(s,'E20:P20',f'{di+1:02} — '+pt,f'{di+1:02} — '+en,'panel')
  # Editable contextual controls; identical sources feed every downstream view.
  for key,(target,addr) in st['panel'].items():
   if target!=sn:continue
   _,lpt,len_,val,unit,note=next(x for x in st['items'] if x[0]==key)
   r,c=xlsxwriter.utility.xl_cell_to_rowcol(addr);s.merge_range(r,c,r,c+2,'',ff['inputtext'] if isinstance(val,str) else ff['inputpct'] if unit.startswith('%') else ff['input']);s.write(r,c,val,ff['inputtext'] if isinstance(val,str) else ff['inputpct'] if unit.startswith('%') else ff['input'])
   labelc=max(4,c-3);s.merge_range(r,labelc,r,c-1,'',ff['small']);text(s,col(labelc)+str(r+1),lpt,len_,'small');s.write_comment(r,c,note)
   if key in st['lists']:s.data_validation(r,c,r,c,{'validate':'list','source':st['lists'][key]})
   bounds={'adr_input':(120,250,5),'occ_input':(.45,.78,.01),'mice_util':(.2,.8,.05),'ffe_rate':(.02,.06,.005),'adr_growth':(0,.06,.005),'stabilized_year':(1,12,1),'owner_priority_on':(0,1,1)}
   if key in bounds:
    lo,hi,step=bounds[key];formula=f'AND(ISNUMBER({addr}),{addr}>={lo},{addr}<={hi},ABS(({addr}-{lo})/{step}-ROUND(({addr}-{lo})/{step},0))<0.000001)'
    s.data_validation(r,c,r,c,{'validate':'custom','value':'='+formula,'error_type':'stop','error_title':'Study range','error_message':f'Use {lo} to {hi}, step {step}.'})
  if di==6:
   merge_text(s,'E22:P22','Pressupostos Hotel · Base Case preserva V06; Custom ativa ADR/ocupação editáveis','Hotel assumptions · Base Case preserves V06; Custom activates ADR/occupancy inputs','small')
   merge_text(s,'E38:P39','MICE: índice normalizado ao cenário original (50%). Capacidade/dias/eventos/participantes carecem de dados documentais. Não se assume ocupação média de 1.500 pessoas.','MICE: index normalised to original case (50%). Capacity/days/events/attendance need documentary inputs. Average attendance is not assumed to be 1,500.','small')
   merge_text(s,'E40:H40','Fees efetivos / bases','Effective fees / bases','panel')
   for j,(lpt,len_,x) in enumerate([('Base / receita','Base / revenue',params['base_fee']),('Agregado M+D+S / receita','Aggregate M+D+S / revenue',params['systems']),('Incentivo / GOP após fees','Incentive / GOP after fees',params['incentive'])]):
    merge_text(s,f'E{42+j}:H{42+j}',lpt,len_,'small');s.write_formula(f'I{42+j}','='+x.f,ff['percent'],x.v)
   s.write_url('K42',"internal:'V06 Matriz ADR'!A1",ff['inputtext'],string='ADR × Occupancy →');s.write_url('K44',"internal:'V06 Operacao'!A1",ff['inputtext'],string='GOPPAR / FF&E →')
   merge_text(s,'E46:P46','Receita e EBITDA · projeção 180 meses','Revenue and EBITDA · 180-month forecast','panel');chart_anchor='E47'
  else:
   if di==5:
    merge_text(s,'E28:P30','Fees residenciais separados do Hotel. A renda paga é limitada à caixa do pool; yield não garantido.','Residential fees remain separate from Hotel. Owner rent is limited by pool cash; yield is not guaranteed.','small')
    s.write_url('E32',"internal:'V06 Operacao'!A1",ff['inputtext'],string='Residential / FF&E →')
   elif di==8:
    merge_text(s,'E26:P28','Reserva geral separada de HOTUSA, FF&E, impostos e compromissos. Cobertura histórica 12M ≥ 1,20× e restantes restrições continuam ativas.','General reserve is separate from HOTUSA, FF&E, tax and commitments. Prior-12-month coverage ≥ 1.20× and other restrictions remain active.','small')
    s.write_url('E30',"internal:'V06 Otimizacao'!A1",ff['inputtext'],string='Reserve / debt / operator comparisons →')
   else:
    target={0:'V06 Hipoteses',1:'V06 Capitulos',2:'V06 Controlos',3:'V06 Otimizacao',4:'V06 Disponibilidade',7:'V06 Capital'}[di]
    merge_text(s,'E23:P24','Abrir pressupostos e detalhe financeiro','Open assumptions and financial detail','small');s.write_url('E26',f"internal:'{target}'!A1",ff['inputtext'],string='→ '+target)
    for j,(lpt,len_,x) in enumerate([('Taxa juro ativa','Active interest rate',params['rate']),('Desconto equity ativo','Active equity discount',params['discount']),('Reserva geral ativa','Active general reserve',params['reserve'])]):
     merge_text(s,f'E{28+j}:J{28+j}',lpt,len_,'small');s.write_formula(f'L{28+j}','='+x.f,ff['percent'] if j<2 else ff['num'],x.v)
   merge_text(s,'E34:P34','Evolução financeira · motor central','Financial trend · central engine','panel');chart_anchor='E35'
  keys=[('cash','reserve'),('capex','construction_insurance'),('vat_credit','tax_paid'),('debt','cash'),('receipts','deed_sales'),('owner_distribution','owner_target'),('revenue','ebitda'),('equity','investor_outlay'),('dividend','cash')][di]
  ch=wb.add_chart({'type':'line'})
  for j,key in enumerate(keys):
   chartlabel=tr('EFCC.chart.'+sn+'.'+key,*{'cash':('Caixa','Cash'),'reserve':('Reserva protegida','Protected reserve'),'capex':('CAPEX','CAPEX'),'construction_insurance':('Seguro obra','Construction insurance'),'vat_credit':('Crédito IVA','VAT credit'),'tax_paid':('IRPC pago','Tax paid'),'debt':('Dívida','Debt'),'receipts':('Recebimentos','Collections'),'deed_sales':('Escrituras','Deed sales'),'owner_distribution':('Renda suportável','Affordable owner distributions'),'owner_target':('Renda alvo','Owner target'),'revenue':('Receita','Revenue'),'ebitda':('EBITDA','EBITDA'),'equity':('Capital KK','KK capital'),'investor_outlay':('Capital investidores','Investor capital'),'dividend':('Dividendos','Dividends')}[key])
   ch.add_series({'name':'='+chartlabel,'categories':"='V06 Motor'!$A$5:$A$184",'values':f"='V06 Motor'!${col(cols[key])}$5:${col(cols[key])}$184",'values_data':[d[key].v for d in rows],'line':{'color':['183F6B','B31F2B'][j],'width':1.8}})
  ch.set_size({'width':1005,'height':240 if di==6 else 265});ch.set_legend({'position':'bottom','font':{'size':9}});ch.set_chartarea({'border':{'none':True}});ch.set_plotarea({'border':{'none':True}});ch.set_y_axis({'num_format':'0,," M€"','major_gridlines':{'visible':True,'line':{'color':'E5EBF1'}}});ch.set_x_axis({'label_interval':12});s.insert_chart(chart_anchor,ch,{'x_scale':.94,'y_scale':1})
  footer=61 if di==6 else 51
  s.merge_range(f'E{footer}:P{footer+1}','',ff['alert']);s.write_formula(f'E{footer}',f'=IF({metrics["Base"]["gap"].f}>0,IF({lang}="PT","ALERTA: défice de financiamento.","ALERT: funding gap."),IF({metrics["Base"]["npv"].f}<0,IF({lang}="PT","VAL negativo à taxa exigida; não é erro de cálculo.","Negative NPV at required return; this is a financial result."),IF({lang}="PT","Validar contratos, pressupostos e risco.","Validate contracts, assumptions and risk.")))',ff['alert'],'VAL negativo à taxa exigida; não é erro de cálculo.')
  if di!=6:
   merge_text(s,'E55:P57','Simulação de estudo. HOTUSA é reserva jurídica; incentivos sujeitos a aprovação. Pressupostos e documentos externos carecem de validação.','Study simulation. HOTUSA is a legal reserve; incentives require approval. Assumptions and external documents require validation.','small')
   s.write_url('E59',"internal:'V06 Otimizacao'!A1",ff['inputtext'],string='Sensitivity / fees / reserves / debt →')
  s.set_landscape();s.set_paper(9);s.fit_to_pages(1,1);s.print_area('A1:Q64');s.set_margins(.15,.15,.2,.2);s.set_footer('&LKK-01 · EFCC · Candidate · Native Excel acceptance pending&R&P / &N')
 build_analysis(model,params,main,metrics,scenario_models,scenario_defs)
 build_availability(model,params,main)
 build_print_help(model)
 # Effective parameters are recorded once for report reproducibility.
 reg=model.sheet('V06 Registo Ativo',['Key','Effective value','Unit / basis','Source','Reference value'])
 for r,(key,x) in enumerate(params.items(),5):
  reg.write(r-1,0,key);model.cell(reg,r,1,x,'text' if isinstance(x.v,str) else 'num');reg.write(r-1,2,next((d[4] for d in DEFAULTS if d[0]==key),next((d[4] for d in st['items'] if d[0]==key),'see dictionary')));reg.write(r-1,3,x.f);reg.write(r-1,4,st['raw'][key].v if key in st['raw'] else x.v)
 reg.set_column('A:A',28);reg.set_column('B:C',25);reg.set_column('D:D',100);reg.set_column('E:E',25)
 # Additional KPI definitions are bilingual and reference actual appended fields.
 definitions=[('Available Room Nights','Noites disponíveis','Available Room Nights','200 × 365 / 12 when active','nights'),('Occupied Room Nights','Noites vendidas','Occupied Room Nights','Available nights × occupancy','nights'),('TRevPAR','Receita total por quarto disponível','Total revenue per available room','Total revenue / available nights','EUR/night'),('GOPPAR','GOP por quarto disponível','GOP per available room','GOP / available nights','EUR/night'),('EBITDA Margin','Margem EBITDA','EBITDA Margin','EBITDA / total revenue','%'),('FF&E Fund','Fundo restrito FF&E','Restricted FF&E Fund','Prior fund + contribution - replacement spend; excluded from free cash and terminal value','EUR'),('MICE Utilisation','Índice de utilização MICE','MICE utilisation index','Legacy MICE revenue × index / 50%; documentary event volumes unavailable','% index'),('Break-even','Limiar de equilíbrio','Break-even','Lowest tested ADR/occupancy where recalculated investor NPV >= 0, holding the other selected coordinate fixed','grid only')]
 ds=wb.get_worksheet_by_name('V06 Dicionario');start=ds.dim_rowmax+2
 for i,(key,pt,en,formula,unit) in enumerate(definitions,start):ds.write_row(i,0,[key,pt,en,formula,formula,formula,'V06 Operacao / V06 Matriz ADR',unit,'Validar pressupostos','Validate assumptions','Context','Documented only','Forecast / source limitations apply'])
 st['reference']={'fields':{k:v.v for k,v in fields.items()},'effective':{k:v.v for k,v in params.items()},'matrix_cases':st['matrix_cases'],'logo_embedded':'logo_sha256' in st}


def build_analysis(model,p,main,metrics,mods,defs):
 st=STATE;wb=st['wb'];ff=st['fm2'];fm=st['fm'];f=st['fields'];tr=st['tr'];lang=st['lang']
 s=model.sheet('V06 Matriz ADR',['ADR × Occupancy']);s.set_column('A:A',2);s.set_column('B:B',29);s.set_column('C:G',22);s.set_column('H:M',18);s.freeze_panes(6,2);s.hide_row_col_headers()
 s.merge_range('B1:G1','',ff['panel']);s.write_formula('B1','='+tr('EFCC.matrix.title','Matriz de Sensibilidade ADR × Ocupação','ADR × Occupancy Sensitivity Matrix'),ff['panel'],'Matriz de Sensibilidade ADR × Ocupação')
 s.write('A3','KPI');s.write('E3','Discount');
 for key in ['matrix_metric','discount_mode']:
  _,addr=st['panel'][key];r,c=xlsxwriter.utility.xl_cell_to_rowcol(addr);val=f[key].v;s.write(r,c,val,ff['inputtext']);s.data_validation(r,c,r,c,{'validate':'list','source':st['lists'][key]})
 s.write_url('H3',"internal:'V06 Controlos'!B65",ff['inputtext'],string='Custom discount →')
 s.write_row('C6',[150,175,200,225,250],ff['num'])
 for ri,oc in enumerate([.5,.55,.6,.65,.7,.75],7):s.write(ri-1,1,oc,ff['percent']);s.data_validation(ri-1,1,ri-1,1,{'validate':'decimal','criteria':'between','minimum':.45,'maximum':.78})
 s.data_validation('C6:G6',{'validate':'decimal','criteria':'between','minimum':120,'maximum':250})
 cases={label:mod for mod,(label,_) in zip(mods,defs)};st['models']=cases
 # Dedicated raw results keep the heatmap numeric and do not use Excel data tables.
 h=model.sheet('V06 Matriz Valores',['ADR','Occupancy','NPV','IRR','Hotel EBITDA selected year','RevPAR opening','Exit Equity Value','Gap','Peak debt'])
 st['matrix_rows']={}
 for i,(label,adr,oc) in enumerate(st['matrix_cases'],5):
  mi=metrics[label];mod=cases[label];rs=mod['rows'];c=mod['columns'];start=30+p['construction_delay']+(f['stabilized_year']-1)*12+1
  end=start+11;a=max(1,min(169,int(start.v)));eb=X(f"SUM(INDEX('{mod['sheet']}'!${col(c['ebitda'])}$5:${col(c['ebitda'])}$184,{start.f}):INDEX('{mod['sheet']}'!${col(c['ebitda'])}$5:${col(c['ebitda'])}$184,{end.f}))",sum(x['ebitda'].v for x in rs[a-1:a+11]))
  rev=X(f"INDEX('{mod['sheet']}'!${col(c['revpar'])}$5:${col(c['revpar'])}$184,31+{p['construction_delay'].f})",rs[30+int(p['construction_delay'].v)]['revpar'].v)
  idx=i-5;ac=f"'V06 Matriz ADR'!{col(2+idx%5)}6";ocref=f"'V06 Matriz ADR'!B{7+idx//5}"
  vals=[X(ac,adr),X(ocref,oc),mi['npv'],mi['irr'],eb,rev,mi['exit'],mi['gap'],mi['debt']]
  for j,x in enumerate(vals):model.cell(h,i,j,x,'pct' if j in [1,3] else 'num')
  r=6+idx//5;c0=2+idx%5;formula=f'INDEX(\'V06 Matriz Valores\'!$C${i}:$G${i},1,MATCH({f["matrix_metric"].f},{{"NPV","IRR","EBITDA","RevPAR","Exit Equity Value"}},0))';vv=mi['npv'].v
  s.write_formula(r,c0,'='+formula,ff['num'],vv);s.conditional_format(r,c0,r,c0,{'type':'formula','criteria':f'AND({f["matrix_metric"].f}="IRR",ISNUMBER({col(c0)}{r+1}))','format':wb.add_format({'num_format':'0.0%'})});st['matrix_rows'][label]=i
 s.conditional_format('C7:G12',{'type':'3_color_scale','min_color':'E6B9BC','mid_type':'num','mid_value':0,'mid_color':'FFF0C7','max_color':'CEE2D4'})
 s.merge_range('B14:G15','',ff['small']);s.write_formula('B14','='+tr('EFCC.matrix.note','30 cenários completos de 180 meses. Custos, fees, impostos, dívida, dividendos e saída recalculam no mesmo motor V06. Valores de estudo.','30 full 180-month cases. Costs, fees, tax, debt, dividends and exit recalculate through the same V06 engine. Study values.'),ff['small'],'30 cenários completos de 180 meses; mesmo motor V06.')
 # Break-even only within the tested grid and at the selected other coordinate.
 b=model.sheet('V06 Equilibrio',['Test','Result','Scope / limitation'])
 seladr=f['selected_adr'];selocc=f['selected_occ'];adrr="'V06 Matriz Valores'!$A$5:$A$34";occr="'V06 Matriz Valores'!$B$5:$B$34";npvr="'V06 Matriz Valores'!$C$5:$C$34"
 tested=[(adr,oc,metrics[label]['npv'].v) for label,adr,oc in st['matrix_cases']]
 adrpass=[a for a,o,n in tested if abs(o-selocc.v)<1e-9 and n>=0];occpass=[o for a,o,n in tested if a==seladr.v and n>=0]
 # MINIFS is explicit _xlfn for Excel 365 and can be independently recalculated by LO.
 for r,label,formula,cache in [(5,'ADR for NPV >= 0 at selected occupancy',f'IF(COUNTIFS({occr},{selocc.f},{npvr},">=0")>0,_xlfn.MINIFS({adrr},{occr},{selocc.f},{npvr},">=0"),"No break-even in tested grid")',min(adrpass) if adrpass else 'No break-even in tested grid'),(6,'Occupancy for NPV >= 0 at selected ADR',f'IF(COUNTIFS({adrr},{seladr.f},{npvr},">=0")>0,_xlfn.MINIFS({occr},{adrr},{seladr.f},{npvr},">=0"),"No break-even in tested grid")',min(occpass) if occpass else 'No break-even in tested grid')]:
  b.write(r-1,0,label);b.write_formula(r-1,1,'='+formula,fm['pct'] if r==6 else fm['num'],cache);b.write(r-1,2,'Discrete tested grid; other coordinate held fixed; no interpolation or extrapolation')
 b.write('A7','Required stabilized EBITDA');b.write_formula('B7',f'=IF(COUNTIF({npvr},">=0")>0,_xlfn.MINIFS(\'V06 Matriz Valores\'!$E$5:$E$34,{npvr},">=0"),"No break-even in tested grid")',fm['num'],min([hval for label,adr,oc in st['matrix_cases'] if metrics[label]['npv'].v>=0 for hval in [sum(d['ebitda'].v for d in cases[label]['rows'][78:90])]] or [0]));b.write('C7','Minimum selected-year EBITDA among feasible tested cells; not an algebraic or continuous threshold')
 b.write('A8','Required equity return');model.cell(b,8,1,p['discount'],'pct');b.write('A9','Selected ADR');model.cell(b,9,1,seladr);b.write('A10','Selected occupancy');model.cell(b,10,1,selocc,'pct');b.write('A11','Reference NPV');model.cell(b,11,1,metrics['Reference Case']['npv']);b.write('A12','Active NPV');model.cell(b,12,1,metrics['Base']['npv']);b.write('A13','Distance active minus reference NPV');model.cell(b,13,1,metrics['Base']['npv']-metrics['Reference Case']['npv'])
 s.write_url('B17',"internal:'V06 Equilibrio'!A1",ff['inputtext'],string='Break-even / selected coordinates →');s.write_url('B19',"internal:'DASH 05 V06'!A1",ff['inputtext'],string='← Hotel & MICE');s.print_area('A1:M20');s.set_landscape();s.set_paper(9);s.fit_to_pages(1,1)
 o=model.sheet('V06 Otimizacao',['Case','IRR','NPV','MOIC','Peak debt','Dividends','Exit equity','Gap','Reserve gap','Interest','IRPC paid'])
 for r,label in enumerate(['Reference Case','Base']+[x[0] for x in defs if x[0].startswith(('Hotel ','Residential ','Reserve ','Debt '))],5):
  mi=metrics[label];o.write(r-1,0,label)
  for j,key in enumerate(['irr','npv','moic','debt','dividends','exit','gap','reservegap','interest','tax'],1):model.cell(o,r,j,mi[key],'pct' if key=='irr' else 'num')
 o.set_column('A:A',30);o.set_column('B:K',22)
 op=model.sheet('V06 Operacao',['Month','Available nights','Occupied nights','Rooms EUR','F&B EUR','Other EUR','MICE EUR','Total EUR','RevPAR','TRevPAR','GOP','GOPPAR','EBITDA','EBITDA Margin','Operator fees','FF&E allocation','FF&E spend forecast','FF&E restricted fund','CFADS','MICE venue','MICE event F&B','MICE services','MICE direct costs','MICE contribution','Residential revenue','Residential management','Residential aggregate','Owner distribution','Owner shortfall','Residential FF&E','Pool fund','KK residual','General reserve','HOTUSA reserve','Known tax protection','Debt reserve','Other reserve','FF&E fund check'])
 keys=['month','available_room_nights','occupied_room_nights','room_revenue','fb_revenue','other_revenue','mice_revenue','revenue','revpar','trevpar','gop','goppar','ebitda','ebitda_margin','operator_fees','ffe','ffe_replacement_spend','ffe_restricted_fund','cfads','mice_venue','mice_event_fb','mice_services','mice_direct_cost','mice_contribution','res_revenue','res_management_fee','res_distribution_fee','owner_distribution','owner_shortfall','res_ffe','res_fund','res_kk','general_reserve','hotusa_reserve','known_tax_protection','debt_restricted_reserve','other_restricted_reserve','ffe_fund_check']
 for r,d in enumerate(main['rows'],5):
  for j,key in enumerate(keys):model.cell(op,r,j,d[key],'pct' if key=='ebitda_margin' else 'int' if key=='month' else 'num')
 op.autofilter(3,0,183,len(keys)-1)

def build_availability(model,p,main):
 st=STATE;f=st['fields'];s=model.sheet('V06 Disponibilidade',['Unit ID','Type','Floor','Interior m²','Balcony m²','Area m²','Price/m² EUR','Price EUR forecast','Documented status','Forecast status at cutoff','Deposit %','Deposit EUR','Sale / deposit month forecast','Progress instalments #','Instalment EUR','Last instalment month','Deed month forecast','Actual received EUR','Documented receivable EUR','Forecast received to cutoff EUR','Forecast remaining EUR','Schedule total EUR','Schedule check EUR','Actual evidence records','Payment schedule','Contract amount EUR documented','Contract date documented','Forecast vs actual source'])
 uws=model.wb.get_worksheet_by_name('V06 Unidades');uws.write('T4','Contract amount EUR documented',st['fm']['head']);uws.write('U4','Contract date documented',st['fm']['head']);uws.write('V4','Contract reference',st['fm']['head'])
 cutoff=f['status_month'];cal='V06 Calend 01';rs=main['rows'];st['availability_summary']={}
 for r,u in enumerate(model.units,5):
  base=lambda c:X(f"'V06 Unidades'!{c}{r}",None)
  uid=u['id'];price=u['price'];sale=u['sale']+p['sales_delay'];deed=IF(p['construction_delay']>0,mx(u['deed']+p['sales_delay'],30+p['construction_delay']),u['deed']+p['sales_delay']);n=mx(0,INT((deed-sale-1)/6));prog=ROUND(price*.55);inst=ROUND(prog/mx(1,n));area=u['area']
  uws.data_validation(r-1,14,r-1,14,{'validate':'list','source':['Unknown','Available','Reserved','Sold','Deeded']});uws.write(r-1,14,'Unknown',st['fm']['input']);uws.write_blank(r-1,19,None,st['fm']['input']);uws.write_blank(r-1,20,None,st['fm']['input']);uws.write_blank(r-1,21,None,st['fm']['input'])
  actualf=f'SUMIF(\'V06 Cobrancas\'!$A$5:$A$504,\'V06 Unidades\'!A{r},\'V06 Cobrancas\'!$C$5:$C$504)';records=f'COUNTIF(\'V06 Cobrancas\'!$A$5:$A$504,\'V06 Unidades\'!A{r})'
  # Blank documentary inputs remain unavailable, never fabricated contractual debt.
  act=X(f'IF({records}=0,"Not documented",{actualf})','Not documented')
  due=X(f'IF(AND(ISNUMBER(\'V06 Unidades\'!T{r}),{records}>0),MAX(0,\'V06 Unidades\'!T{r}-{actualf}),"Not documented")','Not documented')
  received=X(f"SUM('{cal}'!M{r}:INDEX('{cal}'!$M${r}:$GJ${r},1,MAX(1,{cutoff.f})))",sum(0 for _ in []))
  # Determine cached forecast using the same scalar schedule factory reference values.
  pv=price.v;sv=sale.v;dv=deed.v;nv=n.v;a=ROUND(pv*.05).v;b=ROUND(pv*.20).v;pr=ROUND(pv*.55).v;iv=ROUND(pr/max(1,nv)).v
  cached=sum((a+b if m==sv else 0)+(pv-a-b-pr if m==dv else 0)+(pr-iv*(nv-1) if m==sv+6*nv else iv if nv>0 and m>sv and m<dv and (m-sv)%6==0 and (m-sv)/6<=nv else 0) for m in range(1,int(cutoff.v)+1));received.v=cached
  total=X(f"SUM('{cal}'!M{r}:GJ{r})",pv)
  schedule=X(f'"5% + 20% M"&{sale.f}&"; 55% / "&{n.f}&" half-years; 20% M"&{deed.f}',f'5% + 20% M{sv}; 55% / {nv} half-years; 20% M{dv}')
  values=[uid,u['tipo'],u['piso'],u['interior'],u['varanda'],area,price/area,price,X(f"'V06 Unidades'!O{r}",'Unknown'),IF(cutoff>=deed,'Deeded',IF(cutoff>=sale,'Sold','Available')),.05,ROUND(price*.05),sale,n,inst,sale+6*n,deed,act,due,received,price-received,total,total-price,X(records,0),schedule,X(f'IF(ISNUMBER(\'V06 Unidades\'!T{r}),\'V06 Unidades\'!T{r},"Not documented")','Not documented'),X(f'IF(\'V06 Unidades\'!U{r}="","Not documented",\'V06 Unidades\'!U{r})','Not documented'),asx('Forecast engine; actual documentary status and receipts separate')]
  for j,x in enumerate(values):model.cell(s,r,j,x,'text' if isinstance(asx(x).v,str) else 'pct' if j==10 else 'num')
 s.autofilter(3,0,97,len(values)-1);s.set_column('A:C',15);s.set_column('D:X',22);s.set_column('Y:AB',52)
 # Monthly aggregate reconciliation for all 180 months.
 rec=model.sheet('V06 Conciliacao Vendas',['Month','Individual schedules EUR','Consolidated engine EUR','Difference EUR','Units total','Status count check'])
 for m,d in enumerate(main['rows'],1):
  r=m+4;ind=X(f"SUM('{cal}'!{col(12+m-1)}5:{col(12+m-1)}98)",d['receipts'].v)
  for j,x in enumerate([m,ind,d['receipts'],ind-d['receipts'],94,d['inventory_check']]):model.cell(rec,r,j,x,'int' if j in [0,4,5] else 'num')
 st['availability_summary']={'available':X('COUNTIF(\'V06 Disponibilidade\'!J5:J98,"Available")',sum(cutoff.v<u['sale'].v+p['sales_delay'].v for u in model.units)),'sold':X('COUNTIF(\'V06 Disponibilidade\'!J5:J98,"Sold")',sum(u['sale'].v+p['sales_delay'].v<=cutoff.v<u['deed'].v+p['sales_delay'].v for u in model.units)),'deeded':X('COUNTIF(\'V06 Disponibilidade\'!J5:J98,"Deeded")',sum(cutoff.v>=u['deed'].v+p['sales_delay'].v for u in model.units)),'received':rng([d['receipts'] for d in main['rows']])}

def build_print_help(model):
 st=STATE;ff=st['fm2'];s=model.sheet('V06 Imprimir',['PT','EN']);s.set_column('A:B',75);s.set_row(4,80)
 instructions=[('Imprimir / Exportar PDF — procedimento funcional sem macros','Print / Export PDF — working macro-free procedure'),('1. Volte ao dashboard pretendido. Conclua o recálculo: Ctrl+Alt+F9 no Excel. Verifique se existem erros.','1. Return to the requested dashboard. Complete recalculation: Ctrl+Alt+F9 in Excel. Check for errors.'),('2. Prima Ctrl+P. Escolha Imprimir folhas ativas. As áreas e a paginação A4 estão configuradas.','2. Press Ctrl+P. Select Print Active Sheets. A4 print areas and pagination are configured.'),('3. Escolha Microsoft Print to PDF ou Exportar > Criar PDF/XPS. Confirme a pré-visualização e guarde.','3. Choose Microsoft Print to PDF or Export > Create PDF/XPS. Check print preview and save.'),('Relatório integral: DASH 10 V06, 10 páginas. Resumo executivo: V06 Brief, 2 páginas. O seletor Formato abre a folha correta.','Full report: DASH 10 V06, 10 pages. Executive brief: V06 Brief, 2 pages. Format selector opens the correct sheet.'),('Os botões abrem este fluxo e não executam exportação automática. O ficheiro não contém VBA nem macros ocultas.','Buttons open this workflow and do not execute automatic export. This workbook has no VBA or hidden macros.'),('Certificação Microsoft Excel 365: NOT TESTED neste ambiente Linux. PDFs anexos produzidos por LibreOffice.','Microsoft Excel 365 certification: NOT TESTED in this Linux environment. Attached PDFs were produced by LibreOffice.')]
 for r,row in enumerate(instructions,5):s.write_row(r-1,0,row,ff['small']);s.set_row(r-1,55)
 for r,(sn,pt,en) in enumerate(DASHES,14):s.write_url(r-1,0,f"internal:'{sn}'!A1",ff['text'],string=pt);s.write_url(r-1,1,f"internal:'{sn}'!A1",ff['text'],string=en)
 s.write_url('A25',"internal:'V06 Brief'!A1",ff['text'],string='Executive Brief — 2 pages')
