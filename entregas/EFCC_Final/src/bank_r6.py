"""R6 banking extension. Reuses all 68 existing monthly engines, never an NPV proxy.
R4 Reference is immutable. Official bank profile uses the user-confirmed R4 TAN 6% / grace 42m.
Build from the verified R5 checkpoint; original graphic parts are not rewritten.
"""
from closure_r5 import *
from datetime import datetime,timezone
BOUT=ROOT/'Fecho_R6_Banco';BOUT.mkdir(exist_ok=True)
BM=BOUT/'KK01_Executive_Financial_Control_Center_R6_1_BANK_CHECKPOINT.xlsx'
BW=WORK.parent/'bank_r6';BW.mkdir(exist_ok=True)
PRO='V06 Banco Perfis';DAT='V06 Banco Dados';SCH='V06 Banco Servico';REP='V06 Banco Relatorio';BRF='V06 Banco Resumo'
BN='DASH 03 V06';CENTER='DASH 10 Centro'
# key, PT, EN, R4 reference, public Custom address, kind/options
PARAMS=[
 ('tan','TAN fixa anual','Fixed annual nominal rate',.06,'H23','pct'),('grace','Carência capital · meses','Principal grace · months',42,'N23','whole'),('term','Prazo após carência · meses','Term after grace · months',120,'H25','whole'),('sweep','Cash sweep · fração','Cash sweep fraction',1,'N25','pct'),('draw_end','Última utilização · mês','Last draw month',30,'H27','whole'),('amort','Amortização','Amortisation','Linear','H29',['Linear','Bullet']),('facility','Linha autorizada €','Authorised facility EUR',30000000,'H31','money'),
 ('rate_type','Tipo de taxa','Rate type','Fixed','H454',['Fixed','Variable']),('index','Indexante anual','Annual index',.025,'N454','pct'),('spread','Spread anual','Annual spread',.035,'H456','pct'),('interest_grace','Carência juros · meses','Interest grace · months',0,'N456','whole'),('opening','Comissão abertura / linha','Opening fee / facility',.01,'H458','pct'),('opening_fixed','Comissão abertura fixa €','Fixed opening fee EUR',0,'N458','money'),('broker','Intermediação / linha','Broker fee / facility',.015,'H460','pct'),('availability','Disponibilização / linha','Availability fee / facility',.0125,'N460','pct'),('commitment','Compromisso anual','Annual commitment fee',0,'H462','pct'),('commit_base','Base comissão compromisso','Commitment fee base','Undrawn cumulative','N462',['Undrawn cumulative','Undrawn outstanding']),('management','Comissão gestão €','Management fee EUR',0,'H464','money'),('period','Periodicidade gestão · meses','Management period · months',12,'N464','whole'),('prepay','Comissão amortização antecipada','Early repayment fee',0,'H466','pct'),('closing','Comissão liquidação final €','Final settlement fee EUR',0,'N466','money'),('guarantee','Garantias €','Guarantees EUR',0,'H468','money'),('expenses','Despesas bancárias €','Bank expenses EUR',0,'N468','money'),('insurance','Seguro crédito adicional €','Additional credit insurance EUR',0,'H470','money'),('draw_start','Primeira utilização · mês','First draw month',1,'N470','whole'),('convention','Convenção de juros','Interest convention','30/360','H472',['30/360','ACT/365']),('start_date','Data indicativa M1 (hipótese)','Indicative M1 date (assumption)',45931,'N472','date'),('prepay_source','Origem amortização antecipada','Early repayment source','Existing','H474',['Existing','None','Excess cash','Sales','Custom']),('prepay_custom','Amortização máxima mensal €','Custom maximum monthly repayment EUR',0,'N474','money'),('dscr_min','Meta interna DSCR','Internal DSCR target',1.2,'H476','ratio'),('include_fees','DSCR inclui demais comissões · 0/1','DSCR includes other fees · 0/1',0,'N476','whole'),('status','Natureza das condições','Terms status','Hypothetical','H478',['Hypothetical','Bank proposal','Approved']),('evidence','Referência documental','Documentary reference','Not supplied','N478','text'),('institution','Instituição / destinatário','Institution / recipient','To identify','H480','text')]
P={x[0]:i+5 for i,x in enumerate(PARAMS)}
def d(k):return q(DAT,'$B$'+str(P[k]))
def reference():return q('EXEC V06','H23')+'="Reference"'
def build_modules():
 f=BW/'Bank_Modules.xlsx';b=xlsxwriter.Workbook(f);fmt={}
 for name,opts in {'text':{'font_size':10},'head':{'bold':True,'font_color':'FFFFFF','bg_color':'142337'},'section':{'bold':True,'bg_color':'E8EEF5','font_color':'142337'},'money':{'num_format':'#,##0 "€";[Red](#,##0) "€"'},'pct':{'num_format':'0.00%'},'ratio':{'num_format':'0.00"×"'},'whole':{'num_format':'0'},'date':{'num_format':'yyyy-mm-dd'},'input':{'bg_color':'E5EFFA','font_color':'1765A8'},'note':{'font_size':9,'font_color':'687586','text_wrap':True}}.items():fmt[name]=b.add_format({'font_name':'Aptos','font_size':10,'valign':'vcenter',**opts})
 def wf(s,addr,formula,kind='text',value=0):s.write_formula(addr,'='+formula,fmt[kind],value)
 ss={n:b.add_worksheet(n) for n in [PRO,DAT,SCH,REP,BRF]}
 for sn,s in ss.items():
  s.hide_gridlines(2);s.freeze_panes(4,1);s.set_column('A:A',12);s.set_column('B:B',32);s.set_column('C:AZ',18);wf(s,'A1','HYPERLINK("#\'EXEC V06\'!A1",'+tr('Início','Home')+')','input','Início');wf(s,'D1','HYPERLINK("#\''+BN+'\'!A1",'+tr('Voltar ao Dashboard','Back to Dashboard')+')','input','Voltar');s.write('A2',sn,fmt['section'])
 # Fixed profiles are never replaced by adoption of a test. 8 blank save slots.
 s=ss[PRO];s.protect('KK01_bank_reference',{'select_locked_cells':True,'select_unlocked_cells':True});s.set_column('A:A',28)
 s.write('A4','Official KK-01 bank base: TAN 6%, principal grace 42 months, aligned with R4.',fmt['note'])
 s.write_row(5,0,['Profile']+[x[2] for x in PARAMS],fmt['head'])
 official={k:v for k,_,_,v,_,_ in PARAMS};official.update(institution='KK-01 reference - no bank approval')
 profiles=['Official KK-01 6%','Proposal Bank A','Proposal Bank B','Custom']+[f'Saved {i:02}' for i in range(1,9)]
 unlocked={}
 for i,name in enumerate(profiles,8):
  s.write(i-1,0,name,fmt['section'])
  values=dict(official)
  if i==9:values.update(tan=.055,status='Hypothetical',institution='Bank A - to identify')
  if i==10:values.update(tan=.065,opening=.01,grace=48,status='Bank proposal',institution='Bank B - to identify')
  for j,(key,pt,en,rv,addr,kind) in enumerate(PARAMS,1):
   form='IF('+reference()+','+str(rv)+',0)' if False else None
   kfmt=kind if isinstance(kind,str) and kind in fmt else 'text'
   if i==11:wf(s,col(j)+str(i),q(BN,addr),kfmt,rv)
   elif i==8:s.write(i-1,j,values[key],fmt[kfmt])
   else:
    # Saved slots require explicit identification; no accidental valid emission.
    value=values[key] if i<11 else (('Not saved' if key in ['evidence','institution'] else 'Hypothetical') if key in ['evidence','institution','status'] else rv)
    if kfmt not in unlocked:unlocked[kfmt]=b.add_format({'font_name':'Aptos','font_size':10,'locked':False,'bg_color':'E5EFFA','num_format':fmt[kfmt].num_format or 'General'})
    s.write(i-1,j,value,unlocked[kfmt])
   if i!=8 and i!=11 and isinstance(kind,list):s.data_validation(i-1,j,i-1,j,{'validate':'list','source':kind})
 s.write('A23','Save: copy Custom row B11:AI11 and paste VALUES into a Saved row; rename recipient/evidence. Never paste over Official row.',fmt['note']);s.set_row(22,36)
 # Effective bank inputs and immutable source values.
 s=ss[DAT];s.write_row(3,0,['Parameter','Active effective','Reference R4','Selected bank profile','Comparison profile'],fmt['head'])
 for i,(key,pt,en,rv,addr,kind) in enumerate(PARAMS,5):
  s.write(i-1,0,key);s.write(i-1,2,rv);j=i-4
  lookup=lambda selector:'INDEX('+q(PRO,'$B$8:$AI$19')+',MATCH('+q(BN,selector)+','+q(PRO,'$A$8:$A$19')+',0),'+str(j)+')'
  wf(s,'D'+str(i),lookup('H33'),'text',rv);wf(s,'E'+str(i),lookup('N33'),'text',official[key]);wf(s,'B'+str(i),'IF('+reference()+',C'+str(i)+',D'+str(i)+')','text',rv)
 wf(s,'B42','IF('+d('rate_type')+'="Variable",'+d('index')+'+'+d('spread')+','+d('tan')+')','pct',.06)
 wf(s,'B43','(1+B42/12)^12-1','pct',1.005**12-1)
 # Monthly projections: no duplicated financial calculations.
 s=ss[SCH];headers=['Month','Opening debt','Drawdowns','Scheduled principal','Interest due','Interest paid','Capitalised interest','Regular fees','Early principal','Early fee','Closing fee','Total fees','Contractual debt service','All bank payments','Closing debt','CFADS (engine definition)','DSCR','Cash less reserves','Unused facility','Funding gap','Borrower net financing CF','Debt check','Cash check']
 s.write_row(3,0,headers,fmt['head']);s.set_column('A:W',18)
 for rr in range(5,185):
  refs=['A','BP','CD','CA','EK','CE','EL',None,'CJ','EM','EN','EO','EP','EQ','CK','BX',None,None,None,'DC',None,'DJ','DI']
  for j,c in enumerate(refs):
   if c:wf(s,col(j)+str(rr),q('V06 Motor',c+str(rr)),'whole' if j==0 else 'money')
  wf(s,'H'+str(rr),q('V06 Motor','BW'+str(rr))+'+'+q('V06 Motor','CC'+str(rr)),'money')
  wf(s,'Q'+str(rr),'IF(M'+str(rr)+'>0,P'+str(rr)+'/M'+str(rr)+',"N/A")','ratio','N/A')
  wf(s,'R'+str(rr),q('V06 Motor','DG'+str(rr))+'-'+q('V06 Motor','BV'+str(rr)),'money')
  wf(s,'S'+str(rr),'MAX(0,'+d('facility')+'-'+q('V06 Motor','CL'+str(rr))+')','money')
  wf(s,'U'+str(rr),'C'+str(rr)+'-N'+str(rr),'money')
  wf(s,'X'+str(rr),'IF(ISNUMBER(Q'+str(rr)+'),Q'+str(rr)+',1E99)','ratio')
  # Boolean flags avoid array formulas / implicit intersection.
  wf(s,'Y'+str(rr),'IF(O'+str(rr)+'>0.005,A'+str(rr)+',0)','whole')
 s.write('A189','Year',fmt['head']);s.write_row(188,1,['Draw','Principal','Interest','Fees','Debt service','CFADS','Closing debt','Min headroom'],fmt['head'])
 for y in range(15):
  rr=190+y;start=5+y*12;end=start+11;s.write(rr-1,0,y+1)
  for j,c in enumerate(['C','D','F','L','M','P'],1):wf(s,col(j)+str(rr),'SUM('+c+str(start)+':'+c+str(end)+')','money')
  wf(s,'H'+str(rr),'O'+str(end),'money');wf(s,'I'+str(rr),'MIN(R'+str(start)+':R'+str(end)+')','money')
 metrics=[('Maximum annual debt service','MAX(F190:F204)','money'),('Minimum DSCR','IF(MIN(X5:X184)=1E99,"N/A",MIN(X5:X184))','ratio'),('Peak outstanding debt','MAX(O5:O184)','money'),('Final debt repayment month','IF(O184>0.005,"Beyond horizon",IF(MAX(Y5:Y184)=0,0,MAX(Y5:Y184)+1))','whole'),('Total interest paid','SUM(F5:F184)','money'),('Total financing fees','SUM(L5:L184)','money'),('Total financing cost','SUM(E5:E184)+SUM(L5:L184)','money'),('Effective annual financing cost','IF(AND(MIN(U5:U184)<0,MAX(U5:U184)>0),(1+IRR(U5:U184,0.005))^12-1,"N/A")','pct'),('Funding gap','MAX(T5:T184)','money'),('Minimum liquidity headroom','MIN(R5:R184)','money'),('Remaining non-revolving capacity','S184','money'),('New investor NPV',q('V06 Cenarios','M5'),'money'),('New investor IRR',q('V06 Cenarios','K5'),'pct'),('MOIC',q('V06 Cenarios','L5'),'ratio'),('Exit equity value',q('V06 Cenarios','J5'),'money'),('Cumulative dividends','SUM('+motor('dividend')+')','money'),('Interest accrued / due','SUM(E5:E184)','money')]
 for rr,(label,formula,kind) in enumerate(metrics,210):s.write(rr-1,0,label,fmt['section']);wf(s,'B'+str(rr),formula,kind)
 # Formula-only emission gate. External emitter MUST additionally rebuild, compare,
 # scan all sheet errors and hash snapshots, never trust cached gate values.
 wf(s,'B230','IF(AND(SUM(C5:C184)>0,MAX(V5:V184)<0.005,MIN(V5:V184)>-0.005,MAX(W5:W184)<0.005,MIN(W5:W184)>-0.005,'+d('institution')+'<>"To identify",'+d('institution')+'<>"Not saved",'+d('evidence')+'<>"Not saved",OR('+d('status')+'<>"Approved",'+d('evidence')+'<>"Not supplied")),"CHECKS PASS - recalculate before export","BLOCKED")','text','BLOCKED')
 s.write('A232','DSCR = engine CFADS / scheduled principal + paid interest + commitment fee; other fees included only when selected. Capitalised interest is paid through subsequent principal. Early repayments are disclosed separately.',fmt['note']);s.set_row(231,50)
 # Chart library reused on dashboard and printable report.
 chartdefs=[('Drawdown & Repayment','column',[('Draw','B'),('Principal','C')]),('Debt Service Schedule','column',[('Principal','C'),('Interest','D')]),('CFADS vs Debt Service','line',[('CFADS','G'),('Debt service','F')]),('Liquidity Critical Months','line',[('Minimum headroom','I')]),('Financing Cost Breakdown','column',[('Interest','D'),('Fees','E')])]
 for j,(title,typ,series) in enumerate(chartdefs):
  ch=b.add_chart({'type':typ})
  for name,c in series:ch.add_series({'name':name,'categories':q(SCH,'$A$190:$A$204'),'values':q(SCH,'$'+c+'$190:$'+c+'$204')})
  ch.set_title({'name':title,'name_font':{'size':11,'color':'142337'}});ch.set_size({'width':480,'height':210});ch.set_legend({'position':'bottom'});ch.set_chartarea({'border':{'none':True}});s.insert_chart(235+j*16,0,ch)
 # 10 fixed A4 pages; formulas use active engine. External emitter substitutes a
 # validated immutable comparison snapshot in the same workbook before release.
 titles=[('Synthèse de négociation bancaire','Bank negotiation summary'),('Conditions et statut documentaire','Terms and documentary status'),('Programme et calendrier de construction','Development and construction timeline'),('Sources, emplois et décaissements','Sources, uses and drawdowns'),('Service de la dette et coût global','Debt service and total cost'),('DSCR et capacité de remboursement','DSCR and repayment capacity'),('Liquidité et mois critiques','Liquidity and critical months'),('Rendement des investisseurs','Investor returns'),('Risques et scénarios de stress','Risks and stress scenarios'),('Réconciliation et traçabilité','Reconciliation and traceability')]
 # Portuguese wording, not French.
 pt_titles=['Síntese de negociação bancária','Condições e estado documental','Programa e calendário de construção','Fontes, aplicações e desembolsos','Serviço da dívida e custo global','DSCR e capacidade de reembolso','Liquidez e meses críticos','Retorno dos investidores','Riscos e cenários de stress','Reconciliação e rastreabilidade']
 for sn,count in [(REP,10),(BRF,2)]:
  s=ss[sn];s.set_column('A:L',8.5);s.set_default_row(14);s.set_paper(9);s.set_portrait();s.fit_to_pages(1,0);s.set_margins(.3,.3,.35,.35);s.print_area(0,0,count*52-1,11);s.set_h_pagebreaks([52*i for i in range(1,count)]);s.set_header('&LKK-01 | JS-CV&RBank negotiation');s.set_footer('&LUnconfirmed terms are hypotheses&C&[Page] / &[Pages]&RControlled version R6')
  for page in range(count):
   r=page*52;title=pt_titles[page] if sn==REP else ['Resumo bancário','Riscos e condições de emissão'][page];en=titles[page][1] if sn==REP else ['Bank executive summary','Risks and issue conditions'][page]
   s.merge_range(r+3,0,r+5,11,'',fmt['head']);wf(s,'A'+str(r+4),tr(title,en),'head',title)
   s.merge_range(r+7,0,r+8,11,'',fmt['section']);wf(s,'A'+str(r+8),q(BN,'H33')+'&" | "&'+d('institution')+'&" | "&'+d('status'),'section','Custom | To identify | Hypothetical')
   s.merge_range(r+10,0,r+12,11,'',fmt['note']);wf(s,'A'+str(r+11),q(SCH,'B230'),'note','BLOCKED')
   for k,(label,formula,kind) in enumerate(metrics[:12]):
    rr=r+15+k*2;s.merge_range(rr-1,0,rr-1,7,label,fmt['text']);s.merge_range(rr-1,8,rr-1,11,'',fmt[kind]);wf(s,'I'+str(rr),q(SCH,'B'+str(210+k)),kind)
   s.merge_range(r+40,0,r+45,11,'',fmt['note']);note=tr('A emissão requer recálculo integral, reconciliação dos 180 meses, identificação das hipóteses e versão congelada. Nenhuma condição é aprovada sem documento. Relatório sujeito a validação nativa Excel 365.','Issue requires a full recalculation, reconciliation of all 180 months, identified assumptions and a frozen version. No terms are approved without documentation. Native Excel 365 validation remains required.')
   wf(s,'A'+str(r+41),note,'note','Emission blocked until validated')
 b.close();return f,profiles

def apply():
 module,profiles=build_modules();merge(MASTER,module,BM)
 with ZipFile(BM) as z:
  wb,pp=paths(z);changes={};styles={};texts=[''.join(n.itertext()) for n in E.fromstring(z.read('xl/sharedStrings.xml'))]
  for c in E.fromstring(z.read(pp[ACT])).find(T('sheetData'))[0]:
   if c.get('t')=='s':
    name=texts[int(c.findtext(T('v')))];
    if name in ['label','input','input_money','input_number','input_pct','note','section','head','pct','money','ratio','number']:styles[name]=int(c.get('s'))
  def sh(sn):return Sheet(E.fromstring(changes.get(pp[sn],z.read(pp[sn]))),styles)
  def validation(s,addr,kind):
   dvs=s.r.find(T('dataValidations'))
   if dvs is None:dvs=E.SubElement(s.r,T('dataValidations'),count='0')
   for old in list(dvs):
    if old.get('sqref')==addr:dvs.remove(old)
   attrs={'type':'list' if isinstance(kind,list) else 'decimal','sqref':addr,'showErrorMessage':'1','errorStyle':'stop','allowBlank':'0'}
   if not isinstance(kind,list):attrs['operator']='between'
   dv=E.SubElement(dvs,T('dataValidation'),**attrs)
   if isinstance(kind,list):E.SubElement(dv,T('formula1')).text='"'+','.join(kind)+'"'
   else:E.SubElement(dv,T('formula1')).text='0';E.SubElement(dv,T('formula2')).text='1' if kind=='pct' else '1000000000'
   dvs.set('count',str(len(dvs)))
  s=sh(BN);s.text('E33:G33','Cenário bancário','Bank scenario','label');s.merge('H33:J33','Custom',style='input');validation(s,'H33',profiles);s.text('K33:M33','Comparação bancária','Bank comparison','label');s.merge('N33:P33','Official KK-01 6%',style='input');validation(s,'N33',profiles)
  s.merge('E34:P35','Condições bancárias e cenários →','HYPERLINK("#\''+BN+'\'!E450",'+tr('Condições bancárias e cenários →','Bank terms and saved scenarios →')+')','input')
  s.merge('E36:P37','TAN ativa: 6%','"TAN: "&TEXT('+q(DAT,'B42')+',"0.00%")&" | "&'+q(BN,'H33')+' & " | " & '+d('status'),'note')
  s.text('E450:Y452','Simulador bancário · parâmetros Personalizados','Bank simulator · Custom parameters','section')
  for key,pt,en,rv,addr,kind in PARAMS:
   if int(addr[1:])<450:continue
   rr=int(addr[1:]);left=addr.startswith('H');s.text(('E' if left else 'K')+str(rr)+':'+('G' if left else 'M')+str(rr+1),pt,en,'label');area=addr+':'+('J' if left else 'P')+str(rr+1);s.merge(area,rv,style='input_money' if kind=='money' else 'input_number' if kind in ['whole','ratio'] else 'input_pct' if kind=='pct' else 'input')
   if kind!='text':validation(s,addr,kind)
  s.merge('Q454:Y458','Perfis guardados →','HYPERLINK("#\''+PRO+'\'!A6",'+tr('Perfis guardados →','Saved profiles →')+')','input');s.merge('Q460:Y464','Serviço da dívida →','HYPERLINK("#\''+SCH+'\'!A1",'+tr('Serviço da dívida →','Debt service schedule →')+')','input');s.merge('Q466:Y470','Emitir no DASH 10 →','HYPERLINK("#\''+CENTER+'\'!E90",'+tr('Emitir no DASH 10 →','Issue from DASH 10 →')+')','input');s.text('E484:Y489','Custom utiliza os campos editáveis; outros perfis usam a tabela de cenários. TAN é nominal, não taxa de desconto. Prazo exclui carência. ACT/365 usa dias reais do calendário indicado; indexante variável constante no horizonte é hipótese. Comissões de referência R4: abertura 1%, broker 1,5%, disponibilização 1,25%. A Base Oficial usa TAN 6% e carência 42 meses, conforme R4; sem aprovação bancária documental.','Custom uses the editable fields; other profiles use the scenario table. TAN is nominal, not the equity discount rate. Tenor excludes grace. ACT/365 uses actual calendar days; a constant variable index is a hypothesis. R4 fees: opening 1%, broker 1.5%, availability 1.25%. The Official Base uses TAN 6% and grace 42 months, aligned with R4; no documentary bank approval.','note')
  changes[pp[BN]]=s.finish()
  # Same active bank source feeds the existing public control chain and all sensitivities.
  s=sh(ACT)
  oldmap={'H23':'tan','N23':'grace','H25':'term','N25':'sweep','H27':'draw_end','H29':'amort','H31':'facility'}
  for i,item in enumerate(INPUTS,5):
   if item['sheet']==BN and item['cell'] in oldmap:
    key=oldmap[item['cell']];formula=q(DAT,'B42') if key=='tan' else d(key);s.put('B'+str(i),item['value'],formula)
  changes[pp[ACT]]=s.finish()
  s=sh('V06 Controlos')
  for rr,key in [(46,'draw_start'),(47,'opening'),(48,'broker'),(49,'availability'),(50,'commitment')]:s.put('B'+str(rr),next(x[3] for x in PARAMS if x[0]==key),d(key))
  changes[pp['V06 Controlos']]=s.finish()
  count=0
  for sn,p in pp.items():
   if sn!='V06 Motor' and not re.fullmatch(r'V06 [SMO]\d+',sn):continue
   s=sh(sn)
   for rr in range(5,185):
    def r(c):return q(sn,c+str(rr))
    def old(c):return s.cells[c+str(rr)].findtext(T('f'))
    def put(c,f,value=0):s.put(c+str(rr),value,f,int(s.cells.get(c+str(rr),{}).get('s','0')))
    originals={c:old(c) for c in ['BW','CC','CD','CE','CA','CJ','CK','CM','CU','DG','DI','DJ','DC','CZ','CV']}
    month=r('A');cap=d('facility');win=d('draw_end');start=d('draw_start')
    # Keep sensitivity-specific annual rate expression; replace only convention.
    factor='IF('+d('convention')+'="ACT/365",DAY(EOMONTH('+d('start_date')+','+month+'-1))/365,1/12)'
    due=originals['CE'].replace('/12','*'+factor)
    paid='IF('+month+'<='+d('interest_grace')+',0,'+due+')'
    # Closed-form draw equation: zero paid interest in grace, otherwise exact rate.
    draw=originals['CD'].replace('/12','*'+factor)
    rate_term=re.search(r"\(\('"+re.escape(sn)+r"'!BP\d+\*(.*?)\)/12\)",originals['CD'])
    # The formula contains a scalar annual-rate term in both opening interest and
    # draw denominator. During capitalisation these cash-interest terms are zero.
    if rate_term:
     annual=rate_term[1];draw=draw.replace(annual,'IF('+month+'<='+d('interest_grace')+',0,'+annual+')')
    else:raise AssertionError((sn,rr,originals['CD']))
    openingtrigger='AND('+r('BQ')+'<0.005,'+r('CB')+'>0,('+r('BZ')+'-'+r('CA')+'-'+r('BP')+'*\'V06 Controlos\'!$B$41/12)<'+r('BV')+'-0.005)'
    regular='IF('+openingtrigger+','+cap+'*('+d('opening')+'+'+d('broker')+'+'+d('availability')+')+'+d('opening_fixed')+',0)+IF('+month+'='+start+','+d('guarantee')+'+'+d('expenses')+'+'+d('insurance')+',0)+IF(AND('+month+'>='+start+','+month+'<='+win+',MOD('+month+'-'+start+',MAX(1,'+d('period')+'))=0),'+d('management')+',0)'
    commit='IF(AND('+month+'>='+start+','+month+'<='+win+'),MAX(0,'+cap+'-IF('+d('commit_base')+'="Undrawn outstanding",'+r('BP')+','+r('BQ')+'))*'+d('commitment')+'*'+factor+',0)'
    # Principal early-payment cap includes fee and protected liquidity. New options
    # do not change the original Legacy sweep branch at the R4 default.
    repaycap='MAX(0,('+r('CF')+'-'+r('BV')+'-'+r('CI')+'-'+d('closing')+')/(1+'+d('prepay')+'))'
    sweep='IF('+d('prepay_source')+'="None",0,MIN(MAX(0,'+r('BP')+'+'+r('CD')+'+'+r('EL')+'-'+r('CA')+'),'+repaycap+',IF('+d('prepay_source')+'="Existing",'+originals['CJ']+',IF('+d('prepay_source')+'="Sales",MAX(0,'+r('O')+'-'+r('Q')+'-'+r('R')+'),IF('+d('prepay_source')+'="Custom",'+d('prepay_custom')+','+repaycap+')))))'
    extended={'BW':commit,'CC':regular,'CD':draw,'CE':paid,'CJ':sweep,'CK':'MAX(0,'+r('BP')+'+'+r('CD')+'+'+r('EL')+'-'+r('CA')+'-'+r('CJ')+')','CM':originals['CM']+'-'+r('EL')+'-'+r('EM')+'-'+r('EN'),'CU':r('EP'),'DG':originals['DG']+'-'+r('EM')+'-'+r('EN'),'DI':originals['DI']+'+'+r('EM')+'+'+r('EN'),'DJ':originals['DJ']+'-'+r('EL'),'DC':'MAX(0,-('+r('CF')+'-'+r('EM')+'-'+r('EN')+'))','CZ':'MAX(0,'+r('CF')+'-'+r('CJ')+'-'+r('BV')+'-'+r('EM')+'-'+r('EN')+')','CV':originals['CV'].replace(r('CU'),'('+r('CE')+'+'+r('CA')+'+'+r('BW')+')')+'-'+r('EM')+'-'+r('EN')}
    for c,f in extended.items():put(c,'IF('+reference()+','+originals[c]+','+f+')',sr.get(sn,c+str(rr))[1] or 0)
    put('EK','IF('+reference()+','+r('CE')+','+due+')');put('EL','MAX(0,'+r('EK')+'-'+r('CE')+')');put('EM',r('CJ')+'*IF('+reference()+',0,'+d('prepay')+')');put('EN','IF(AND('+r('BP')+'+'+r('CD')+'>0.005,'+r('CK')+'<=0.005),IF('+reference()+',0,'+d('closing')+'),0)');put('EO',r('BW')+'+'+r('CC')+'+'+r('EM')+'+'+r('EN'));put('EP',r('CE')+'+'+r('CA')+'+'+r('BW')+'+IF('+reference()+',0,IF('+d('include_fees')+'=1,'+r('CC')+'+'+r('EM')+'+'+r('EN')+',0))');put('EQ',r('CE')+'+'+r('CA')+'+'+r('CJ')+'+'+r('EO'))
   changes[p]=s.finish();count+=180
  s=sh(CENTER);s.text('E90:Y92','Emissão de relatórios bancários','Bank report issuance','section')
  settings=[('E94:J95',tr('Cenário guardado →','Saved scenario →'),'HYPERLINK("#\''+BN+'\'!H33",'+q(BN,'H33')+')'),('K94:P95',tr('Comparação →','Comparison →'),'HYPERLINK("#\''+BN+'\'!N33",'+q(BN,'N33')+')'),('E98:J99',tr('Destinatário →','Recipient →'),'HYPERLINK("#\''+BN+'\'!H480",'+d('institution')+')'),('K98:P99',tr('Validação','Validation'),q(SCH,'B230'))]
  for area,label,formula in settings:s.merge(area,'Bank selection',formula,'input')
  s.text('E102:G103','Formato','Format','label');s.merge('H102:P103','Bank 10 pages',style='input');validation(s,'H102',['Consolidated 10 pages','Bank 10 pages','Executive 2 pages']);s.text('E106:G107','Apresentação','Presentation','label');s.merge('H106:P107','Base vs Test',style='input');validation(s,'H106',['Base vs Test','Selected only'])
  s.text('E110:Y115','Emissão controlada: executar bank_emit.py ou o procedimento Excel 365 incluído. O emissor recalcula o motor completo, verifica erros e saldos, congela pressupostos/resultados e só então gera o PDF. As ligações nativas do Excel não executam recálculo/exportação nem guardam cenários automaticamente.','Controlled issuance: run bank_emit.py or the included Excel 365 procedure. The emitter recalculates the full engine, checks errors and balances, freezes assumptions/results and only then creates the PDF. Native Excel hyperlinks do not execute recalculation/export or automatically save scenarios.','note');s.merge('E118:P119','Relatório bancário →','HYPERLINK("#\''+REP+'\'!A1",'+tr('Relatório bancário →','Bank report →')+')','input');changes[pp[CENTER]]=s.finish()
  # Print appendix numeric columns were too narrow in R5, identified by PDF scan.
  s=sh(CON)
  for c in s.r.find(T('cols')):
   if c.get('min')=='4':c.set('width','20')
  changes[pp[CON]]=s.finish()
  for s in wb.find(T('sheets')):
   if s.get('name') in [DAT]:s.set('state','hidden')
  # Open at the executive dashboard after adding banking helper sheets.
  names=[n.get('name') for n in wb.find(T('sheets'))]
  for v in wb.find(T('bookViews')):v.set('activeTab',str(names.index('EXEC V06')));v.set('firstSheet',str(names.index('EXEC V06')))
  changes['xl/workbook.xml']=xml(wb)
  temp=BM.with_suffix('.temp.xlsx')
  with ZipFile(temp,'w',ZIP_DEFLATED,compresslevel=6) as out:
   for item in z.infolist():out.writestr(copy.copy(item),changes.get(item.filename,z.read(item.filename)))
 temp.replace(BM)
 (BOUT/'Bank_Implementation.json').write_text(json.dumps({'version':'R6 CHECKPOINT','source_r5_sha256':hashlib.sha256(MASTER.read_bytes()).hexdigest(),'engine_months_extended':count,'parameters':[{'key':p[0],'pt':p[1],'en':p[2],'r4_value':p[3],'custom_cell':p[4],'type':p[5],'effective_cell':f'{DAT}!B{P[p[0]]}'} for p in PARAMS],'official_bank_benchmark':{'tan':.06,'grace_months':42,'status':'Hypothetical; not approved','aligned_with':'immutable R4 6% / 42m; confirmed by user'},'native_excel365':'NOT TESTED'},ensure_ascii=False,indent=2));print('Built',BM,count,flush=True)
if __name__=='__main__':apply()
