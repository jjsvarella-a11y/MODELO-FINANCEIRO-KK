"""Editable global Base / independent Comparative, using identical R6.1 engine templates.
No openpyxl serialization; original graphic/chart parts are preserved verbatim.
"""
from bank_r6 import *
from zip_copy import rewrite
from datetime import datetime,timezone
OUT7=ROOT/'Entrega_Consolidada';OUT7.mkdir(exist_ok=True)
WORK7=WORK.parent/'editable_v07';WORK7.mkdir(exist_ok=True)
FINAL7=OUT7/'KK01_Executive_Financial_Control_Center.xlsx'
BASE7='V07 Base';CMP7='V07 Comparativo';HIST7='V07 Historico';REPORT7='V07 Relatorios';STATUS7='V07 Estado';MET7='V07 Resultados'
CLONES=['V06 Motor','V06 Controlos','V06 Registo Ativo','V06 Hipoteses',ACT,DAT,'V06 Curva','V06 Unidades','V06 Capital','V06 Calend 01',CON,'V06 F01','V06 Fluxos Invest','V06 Blocos',SCH]
SHORT=['Motor','Controlos','Registo','Hipoteses','Ativos','Banco','Curva','Unidades','Capital','Calendario','Orcamento','Fluxos','BlocoFluxos','Blocos','Divida']
MAPS={mode:dict(zip(CLONES,['V07 '+prefix+' '+x for x in SHORT])) for mode,prefix in [('Base','B'),('Comparative','C')]}
DEFS=[]
for i,item in enumerate(INPUTS,5):
 dct=dict(item);dct.update(key='I'+str(i),initial=sr.get(item['sheet'],item['cell'])[1] if False else item['value']);DEFS.append(dct)
for key,pt,en,rv,addr,kind in PARAMS:DEFS.append(dict(key='bank_'+key,label_pt=pt,label_en=en,initial=rv,kind=kind,sheet=BN,cell=addr))
for key,pt,en,rv,kind in [('adr','ADR inicial','Opening ADR',185,'money'),('occ','Ocupação inicial','Opening occupancy',.58,'pct'),('growth','Crescimento ADR','ADR growth',.03,'pct'),('mice','Utilização MICE','MICE utilisation',.5,'pct'),('basefee','Base management fee','Base management fee',.03,'pct'),('incentive','Incentive fee','Incentive fee',.06,'pct'),('systems','Marketing + distribuição + sistemas','Marketing + distribution + systems',.03,'pct'),('ffe','Reserva FF&E','FF&E reserve',.05,'pct'),('discount','Taxa desconto investidores','Investor discount rate',.12,'pct')]:DEFS.append(dict(key=key,label_pt=pt,label_en=en,initial=rv,kind=kind))
ROWS={d['key']:i+6 for i,d in enumerate(DEFS)}
IRO={(i['sheet'],i['cell']):'I'+str(j) for j,i in enumerate(INPUTS,5)}
BANKOLD={'H23':'tan','N23':'grace','H25':'term','N25':'sweep','H27':'draw_end','H29':'amort','H31':'facility'}
BANKPAN=[('tan','TAN','TAN',['2.0%','2.5%','3.0%','3.5%','4.0%','4.5%','5.0%','5.5%','6.0%','6.5%','7.0%','7.5%','8.0%','Custom']),('facility','Linha autorizada','Authorised facility',['20000000','25000000','30000000','35000000','40000000','Custom']),('term','Prazo após carência (meses)','Tenor after grace (months)',['60','120','144','180','240','Custom']),('grace','Carência de capital (meses)','Capital grace (months)',['0','12','24','30','36','42','Custom']),('opening','Comissão de abertura','Opening fee',['0%','0.5%','1.0%','1.5%','2.0%','Custom']),('commitment','Comissão de compromisso','Commitment fee',['0%','0.25%','0.50%','0.75%','1.00%','Custom']),('prepay','Amortização antecipada','Early repayment fee',['0%','0.5%','1.0%','1.5%','2.0%','Custom']),('prepay_source','Estratégia de reembolso','Repayment strategy',['Existing','None','Excess cash','Sales','Custom'])]
HOTELPAN=[('adr','ADR inicial','Opening ADR',['Profile','120','140','160','180','185','200','225','250','Custom']),('occ','Ocupação inicial','Opening occupancy',['Profile','45%','50%','55%','58%','60%','65%','70%','75%','78%','Custom']),('growth','Crescimento ADR','ADR growth',['0%','1%','2%','3%','4%','5%','6%','Custom']),('mice','Utilização MICE','MICE utilisation',['20%','30%','40%','50%','60%','70%','80%','Custom']),('basefee','Management fee','Management fee',['Profile','2.0%','2.5%','3.0%','3.5%','4.0%','Custom']),('incentive','Incentive fee','Incentive fee',['Profile','5%','6%','7%','8%','9%','10%','Custom']),('systems','Marketing + distribuição + sistemas','Marketing + distribution + systems',['Profile','1.5%','2.0%','2.5%','3.0%','3.5%','4.0%','Custom']),('ffe','Reserva FF&E','FF&E reserve',['2%','3%','4%','5%','6%','Custom'])]
PANELS={BN:BANKPAN,'DASH 05 V06':HOTELPAN}
PANELREF={('bank_'+key if sn==BN else key,mode):(sn,('H' if mode=='Base' else 'N')+str(23+j*2)) for sn,panel in PANELS.items() for j,(key,*_) in enumerate(panel) for mode in ['Base','Comparative']}
def store(mode):return BASE7 if mode=='Base' else CMP7
def inp(mode,key):return q(store(mode),'$D$'+str(ROWS[key]))
def active(key):return 'IF(\'EXEC V06\'!$H$23="Comparative",'+inp('Comparative',key)+','+inp('Base',key)+')'
def rebound(f,mode):
 # Exact reference parser, including absolute/relative variants. Never edit string labels.
 def sub(m):
  sn=m[1];addr=m[2].replace('$','')
  if sn=='DASH 05 V06' and addr=='N23':return q(sn,'$N$41')
  if sn=='EXEC V06' and addr=='H23':return '"Test"'
  if sn=='V06 Cenarios' and addr in ['J5','K5','L5','M5']:return q(MET7,('$B$' if mode=='Base' else '$C$')+str({'J5':13,'K5':11,'L5':12,'M5':10}[addr]))
  key=IRO.get((sn,addr))
  if key:return inp(mode,key)
  return q(MAPS[mode].get(sn,sn),m[2])
 return re.sub(r"'([^']+)'!(\$?[A-Z]+\$?\d+)",sub,f)
ALIASES={}
for addr,key in BANKOLD.items():ALIASES[IRO[(BN,addr)]]='bank_'+key
for addr,key in {'H25':'adr','H27':'occ','N29':'growth','N25':'mice','N27':'ffe','H29':'basefee','H30':'incentive','N32':'systems'}.items():ALIASES[IRO[('DASH 05 V06',addr)]]=key
ALIASES[IRO[('DASH 07 V06','H29')]]='discount'
ALIASES[IRO[('DASH 07 V06','N25')]]='discount'
def initial_choice(d):
 v=d['initial'];key=d['key'];kind=d.get('kind')
 if key in ['adr','occ','basefee','incentive','systems']:return 'Profile'
 if kind=='pct':return f'{v*100:g}%'
 return v

def build():
 source=Source(BM);b=xlsxwriter.Workbook(WORK7/'Modules.xlsx');formats={}
 for name,opts in {'text':{},'head':{'bold':True,'bg_color':'142337','font_color':'FFFFFF'},'section':{'bold':True,'bg_color':'E9EFF5'},'input':{'bg_color':'DCEAF8','font_color':'1765A8'},'pct':{'num_format':'0.00%'},'money':{'num_format':'#,##0 "€";[Red](#,##0) "€"'},'inputpct':{'bg_color':'DCEAF8','num_format':'0.00%'},'number':{'num_format':'0.00'},'note':{'font_size':9,'font_color':'687586'}}.items():formats[name]=b.add_format({'font_name':'Aptos','font_size':10,'text_wrap':True,'valign':'vcenter',**opts})
 sheets={sn:b.add_worksheet(sn) for sn in [BASE7,CMP7,HIST7,REPORT7,STATUS7,MET7]+[n for mp in MAPS.values() for n in mp.values()]}
 def wf(s,a,f,v=0,kind='text'):s.write_formula(a,'='+f,formats[kind],v)
 for sn,s in sheets.items():
  s.hide_gridlines(2);s.set_column('A:A',37);s.set_column('B:D',22);s.set_column('E:H',25);s.freeze_panes(5,1);s.set_landscape();s.set_paper(9);s.fit_to_pages(1,0);s.repeat_rows(0,4);s.set_margins(.3,.3,.3,.3)
  wf(s,'A1','HYPERLINK("#\'EXEC V06\'!A1",'+tr('Início','Home')+')','Início','input');wf(s,'C1','HYPERLINK("#\'DASH 10 Centro\'!A1",'+tr('Voltar ao Dashboard','Back to Dashboard')+')','Voltar','input');s.write('A2',sn,formats['head'])
 for mode in ['Base','Comparative']:
  s=sheets[store(mode)];s.write_row('A5',['Parâmetro / Parameter','Seleção / Selection','Custom manual','Efetivo / Effective','Último guardado / Last saved','Estado / Status','Origem / Source'],formats['head']);s.write('J5','Last saved selection');s.write('K5','Last saved Custom');s.set_column('J:K',18,None,{'hidden':True})
  for definition in DEFS:
   key=definition['key'];rr=ROWS[key];v=definition['initial'];kind=definition.get('kind');wf(s,'A'+str(rr),tr(definition['label_pt'],definition['label_en']),definition['label_pt']);s.write('C'+str(rr),v,formats['inputpct' if kind=='pct' else 'input']);s.write('E'+str(rr),v,formats['pct' if kind=='pct' else 'text'])
   panel=PANELREF.get((key,mode));choice=initial_choice(definition);s.write('J'+str(rr),choice);s.write('K'+str(rr),v)
   if panel:wf(s,'B'+str(rr),q(*panel),choice)
   else:
    s.write('B'+str(rr),choice,formats['input']);options=definition.get('options') or (kind if isinstance(kind,list) else None)
    if not options and definition.get('sheet') in source.pp:
     oldroot=E.fromstring(source.z.read(source.pp[definition['sheet']]))
     for validation in oldroot.findall('.//'+T('dataValidation')):
      if definition.get('cell') in validation.get('sqref','').split():
       vf=validation.findtext(T('formula1')) or ''
       if vf.startswith(chr(34)):options=vf.strip(chr(34)).split(',')
    if key=='discount':options=['8%','10%','12%','Custom']
    if options:s.data_validation('B'+str(rr),{'validate':'list','source':[str(x) for x in options]})
   if isinstance(v,(int,float)):
    f=f'IF(B{rr}="Custom",C{rr},IF(ISNUMBER(B{rr}),B{rr},IF(RIGHT(B{rr},1)="%",_xlfn.NUMBERVALUE(SUBSTITUTE(B{rr},"%",""),".",",")/100,_xlfn.NUMBERVALUE(B{rr},".",","))))'
   else:f=f'B{rr}'  # Categories remain text, including the legitimate Custom mode.
   if key in ['adr','occ']:
    pf=inp(mode,IRO[('DASH 05 V06','H23')]);rates=[160,185,210,250] if key=='adr' else [.52,.58,.64,.75];profile_formula=f'IF({pf}="Conservative",{rates[0]},IF({pf}="Favorable",{rates[2]},IF({pf}="Upside Stress",{rates[3]},{rates[1]})))';f=f'IF(B{rr}="Profile",{profile_formula},{f})'
   if key in ['basefee','systems','incentive']:
    profile=inp(mode,IRO[('DASH 05 V06','H32')]);rates={'basefee':[.04,.025,.02],'systems':[.04,.03,.015],'incentive':[.10,.06,.05]}[key]
    pf=f'IF({profile}="Operator Proposed Terms",{rates[0]},IF({profile}="KK Target Terms",{rates[1]},IF({profile}="KK Optimized Terms",{rates[2]},C{rr})))'
    f=f'IF(B{rr}="Profile",{pf},{f})'
   if key in ALIASES:
    f=inp(mode,ALIASES[key]);wf(s,'B'+str(rr),f,v)
   wf(s,'D'+str(rr),f,v,'pct' if kind=='pct' else 'money' if kind=='money' else 'text');wf(s,'F'+str(rr),f'IF(D{rr}=E{rr},'+tr('Guardado','Saved')+','+tr('Alterado','Changed')+')','Guardado')
   s.write('G'+str(rr),(definition.get('sheet','Effective parameter')+'!'+definition.get('cell',''))+' — hipótese / assumption',formats['note'])
   if isinstance(v,(int,float)):
    s.data_validation('C'+str(rr),{'validate':'decimal','criteria':'between','minimum':-1 if key.startswith('I') else 0,'maximum':1 if kind=='pct' else 1e9})
  s.autofilter(4,0,len(DEFS)+4,6);s.print_area(0,0,len(DEFS)+4,6)
  s.write('I6','Base = global working assumptions, not an executed contract. Tenor follows principal grace. Custom values are in column C. Saving/version restoration uses scenario_v07.py; a hyperlink cannot run a program.',formats['note']);s.set_column('I:I',55)
 # Clone the same complete monthly engine and its scalar dependencies twice.
 for mode,mp in MAPS.items():
  for old,new in mp.items():
   s=sheets[new];root=E.fromstring(source.z.read(source.pp[old]));s.hide()
   for c in root.iter(T('c')):
    addr=c.get('r');f,v=source.get(old,addr)
    if rc(addr)[0]<3:continue
    if f:s.write_formula(addr,'='+rebound(f,mode),None,v if v is not None else 0)
    elif v is not None:s.write(addr,v)
   if old=='V06 Motor':
    for rr in range(5,185):
     wf(s,'T'+str(rr),f'IF(A{rr}>=D{rr},'+q(mp[ACT],'$B$27')+'*(1+'+q(mp[ACT],'$B$31')+f')^E{rr},0)')
     wf(s,'U'+str(rr),f'IF(A{rr}>=D{rr},MIN(MAX(0.74,'+q(mp[ACT],'$B$28')+'),'+q(mp[ACT],'$B$28')+f'+0.012*E{rr}),0)')
   if old==ACT:
    for j,item in enumerate(INPUTS,5):
     key='I'+str(j);formula=inp(mode,key)
     if item['sheet']==BN and item['cell'] in BANKOLD:formula=q(mp[DAT],'B42') if BANKOLD[item['cell']]=='tan' else inp(mode,'bank_'+BANKOLD[item['cell']])
     if item['sheet']=='DASH 05 V06':
      override={'H25':inp(mode,'adr'),'H27':inp(mode,'occ'),'N29':inp(mode,'growth'),'N25':inp(mode,'mice'),'N27':inp(mode,'ffe'),'H32':'"Custom Terms"','H29':inp(mode,'basefee'),'H30':inp(mode,'incentive'),'N32':inp(mode,'systems'),'H34':'0','N34':'0'};formula=override.get(item['cell'],formula)
     if (item['sheet'],item['cell']) in [('DASH 04B V06','N27'),('DASH 04B V06','H29')]:
      p=inp(mode,IRO[('DASH 04B V06','H33')]);formula=f'IF({p}="Conservative",{formula}*0.9,IF({p}="Favorable",{formula}*1.05,{formula}))'
     if (item['sheet'],item['cell'])==('DASH 04B V06','N29'):formula+='+'+inp(mode,IRO[('DASH 04B V06','H31')])
     wf(s,'B'+str(j),formula,source.get(ACT,'B'+str(j))[1])
   if old==DAT:
    for key,*_ in PARAMS:
     for cc in ['B','D','E']:wf(s,cc+str(P[key]),inp(mode,'bank_'+key),source.get(DAT,'B'+str(P[key]))[1])
   if old=='V06 Registo Ativo':wf(s,'B65',inp(mode,'discount'),.12,'pct')
   if old=='V06 Controlos':wf(s,'B65',inp(mode,'discount'),.12,'pct')
 s=sheets[MET7];s.write_row('A5',['KPI','Base','Comparative','Delta','Delta % / pp'],formats['head'])
 metrics=[('Future CAPEX','SUM(\'V06 Motor\'!L5:L184)'),('Peak Bank Debt','MAX(\'V06 Motor\'!CK5:CK184)'),('Minimum Liquidity Headroom','MIN(\'V06 Motor\'!DG5:DG184-\'V06 Motor\'!BV5:BV184)'),('Planned Dividends','SUM(\'V06 Motor\'!DE5:DE184)'),('Investor NPV',source.get('V06 Cenarios','M5')[0]),('Investor IRR',source.get('V06 Cenarios','K5')[0]),('MOIC',source.get('V06 Cenarios','L5')[0]),('Exit Equity Value',source.get('V06 Cenarios','J5')[0]),('M54 Dividends',source.get('V06 Cenarios','E5')[0]),('Financing Cost','SUM(\'V06 Motor\'!EK5:EK184)+SUM(\'V06 Motor\'!EO5:EO184)'),('Room Revenue','SUM(\'V06 Motor\'!V5:V184)'),('EBITDA','SUM(\'V06 Motor\'!AH5:AH184)'),('Operator Fees','SUM(\'V06 Motor\'!AB5:AD184)'),('Funding Gap','MAX(\'V06 Motor\'!DC5:DC184)')]
 # Scalar headroom helper avoids legacy array semantics.
 for mode,mp in MAPS.items():
  s2=sheets[mp[DAT]]
  for rr in range(5,185):wf(s2,'J'+str(rr),q(mp['V06 Motor'],'DG'+str(rr))+'-'+q(mp['V06 Motor'],'BV'+str(rr)))
 for rr,(label,f) in enumerate(metrics,6):
  s.write('A'+str(rr),label,formats['section'])
  if rr==8:f='MIN(\'V06 Banco Dados\'!J5:J184)'
  for cc,mode in [('B','Base'),('C','Comparative')]:wf(s,cc+str(rr),rebound(f,mode),0,'pct' if rr==11 else 'money')
  wf(s,'D'+str(rr),f'C{rr}-B{rr}',0,'pct' if rr==11 else 'money');wf(s,'E'+str(rr),f'IF(ABS(B{rr})>0.0000001,D{rr}/ABS(B{rr}),"N/A")',0,'pct')
 # Parallel bank comparison projects identical existing schedules, not independent formulas.
 for rr in range(210,227):
  target=rr-185;label=source.get(SCH,'A'+str(rr))[1];sheets[MET7].write('A'+str(target),label,formats['section'])
  for cc,mode in [('B','Base'),('C','Comparative')]:wf(sheets[MET7],cc+str(target),q(MAPS[mode][SCH],'B'+str(rr)))
  wf(sheets[MET7],'D'+str(target),'IF(AND(ISNUMBER(B'+str(target)+'),ISNUMBER(C'+str(target)+')),C'+str(target)+'-B'+str(target)+',"N/A")')
 s=sheets[STATUS7];s.write('A5','Base ID');s.write('B5','BASE-0001');s.write('A6','Comparative ID');s.write('B6','CMP-0001');s.write('A7','Model version');s.write('B7','V07');s.write('A8','Last saved UTC');s.write('B8',datetime.now(timezone.utc).isoformat());s.write('A9','Base save status');s.write('A10','Comparative save status')
 for rr,mode in [(9,'Base'),(10,'Comparative')]:
  sn=store(mode);wf(s,'B'+str(rr),'IF(SUMPRODUCT(--('+q(sn,f'$D$6:$D${len(DEFS)+5}')+'<>'+q(sn,f'$E$6:$E${len(DEFS)+5}')+'))+SUMPRODUCT(--('+q(sn,f'$B$6:$B${len(DEFS)+5}')+'<>'+q(sn,f'$J$6:$J${len(DEFS)+5}')+'))+SUMPRODUCT(--('+q(sn,f'$C$6:$C${len(DEFS)+5}')+'<>'+q(sn,f'$K$6:$K${len(DEFS)+5}')+'))=0,"Saved","Unsaved")','Saved')
 s.write('A12','Native Excel 365');s.write('B12','Not tested — no Windows Excel runtime');s.write('A13','Official logo binary');s.write('B13','Not present in source; V06 Identidade supports original image insertion');s.write('A15','Version operations');s.write('B15','scenario_v07.py save-base / save-comparative / restore');s.set_column('B:B',80)
 s=sheets[HIST7];s.write_row('A5',['ID','Class','UTC','Parameter','Previous','Updated','NPV','IRR','Description','Snapshot SHA256'],formats['head']);s.write_row('A6',['BASE-0001','Reference Base',datetime.now(timezone.utc).isoformat(),'Initial import R6.1',.06,.06,-12742848.2445606,.05764643565,'TAN 6%; grace 42; ADR 185; occupancy 58%; R4 preserved as history',''])
 s=sheets[REPORT7];settings=[(6,'Formato / Format','Executive One Page',['Executive One Page','Bank Financing','Equity Investor','Comprehensive Project','All Formats']),(7,'Comparação / Comparison','Base Only',['Base Only','Base vs Selected','Base vs Multiple']),(8,'Idioma / Language','PT',['PT','EN','Both']),(9,'Destinatário / Recipient','Comité de investimento',None),(10,'Cenários guardados / Saved scenarios','',None),(11,'Confidencialidade / Confidentiality','Confidential',['Confidential','Internal']),(12,'Apêndices / Appendices','Yes',['Yes','No'])]
 for rr,label,v,opts in settings:
  s.write('A'+str(rr),label,formats['section']);s.write('B'+str(rr),v,formats['input'])
  if opts:s.data_validation('B'+str(rr),{'validate':'list','source':opts})
 wf(s,'B14','IF(\''+STATUS7+'\'!B9="Saved","READY FOR FULL RECALCULATION","BLOCKED: UNSAVED BASE")','READY FOR FULL RECALCULATION');s.write('A16','Emitir / Issue');s.write('B16','report_v07.py — full recalculation, reconciliation and immutable JSON/PDF snapshot',formats['note']);s.set_column('B:B',80);s.write('A18','Default: current saved Base only; never revert to R4. TAN and investor discount are distinct. Issued snapshots and historical versions are not overwritten.',formats['note'])
 b.close();merge(BM,WORK7/'Modules.xlsx',FINAL7)
 patch(source)
 from finish_editable_v07 import finish_ui
 finish_ui(FINAL7)
 from tidy_controls_v07 import apply as tidy_controls
 tidy_controls(FINAL7)
 (WORK7/'Definitions.json').write_text(json.dumps(DEFS,ensure_ascii=False,indent=2))
 print(FINAL7,flush=True)

def patch(source):
 changes={}
 with ZipFile(FINAL7) as z:
  wb,pp=paths(z);texts=[''.join(n.itertext()) for n in E.fromstring(z.read('xl/sharedStrings.xml'))];sty={}
  for c in E.fromstring(z.read(pp[ACT])).find(T('sheetData'))[0]:
   if c.get('t')=='s':sty[texts[int(c.findtext(T('v')))]]=int(c.get('s','0'))
  sty['nav']=sty['input'];sty['text']=sty.get('label',0)
  def sh(sn):return Sheet(E.fromstring(changes.get(pp[sn],z.read(pp[sn]))),sty)
  # Active public context routes to one complete global set, never mixed Base/Test.
  for original in CLONES:
   if original not in [ACT,DAT,'V06 Registo Ativo','V06 Controlos']:continue
   s=sh(original)
   if original==ACT:
    for j,item in enumerate(INPUTS,5):s.put('B'+str(j),source.get(ACT,'B'+str(j))[1],'IF(\'EXEC V06\'!$H$23="Comparative",'+q(MAPS['Comparative'][ACT],'B'+str(j))+','+q(MAPS['Base'][ACT],'B'+str(j))+')')
   if original==DAT:
    for key,*_ in PARAMS:
     for cc in ['B','D','E']:s.put(cc+str(P[key]),source.get(DAT,'B'+str(P[key]))[1],active('bank_'+key))
   if original in ['V06 Registo Ativo','V06 Controlos']:s.put('B65',.12,active('discount'))
   changes[pp[original]]=s.finish()
  s=sh('V06 Motor')
  for rr in range(5,185):
   s.put('T'+str(rr),0,f'IF(A{rr}>=D{rr},'+q(ACT,'$B$27')+'*(1+'+q(ACT,'$B$31')+f')^E{rr},0)')
   s.put('U'+str(rr),0,f'IF(A{rr}>=D{rr},MIN(MAX(0.74,'+q(ACT,'$B$28')+'),'+q(ACT,'$B$28')+f'+0.012*E{rr}),0)')
  changes[pp['V06 Motor']]=s.finish()
  # Rewrite remaining direct raw input references (e.g. assumptions / discount).
  def direct(m):
   if m[1]=='DASH 05 V06' and m[2].replace('$','')=='N23':return q(m[1],'$N$41')
   key=IRO.get((m[1],m[2].replace('$','')))
   return '('+active(key)+')' if key else m[0]
  for sn,p in pp.items():
   if sn.startswith('V07 ') or sn in [ACT,DAT]:continue
   root=E.fromstring(changes.get(p,z.read(p)));dirty=False
   for f in root.iter(T('f')):
    if f.text and not f.text.startswith('HYPERLINK('):
     nf=re.sub(r"'([^']+)'!(\$?[A-Z]+\$?\d+)",direct,f.text)
     if nf!=f.text:f.text=nf;dirty=True
   if sn=='DASH 05 V06':
    for f in root.iter(T('f')):
     if f.text:
      nf=re.sub(r'(?<!!)(?<![A-Za-z0-9_])\$?N\$?23\b','$N$41',f.text)
      if nf!=f.text:f.text=nf;dirty=True
   if dirty:changes[p]=xml(root)
  def dv(s,addr,options):
   dvs=s.r.find(T('dataValidations'))
   if dvs is None:dvs=E.SubElement(s.r,T('dataValidations'),count='0')
   for old in list(dvs):
    if addr in old.get('sqref','').split():dvs.remove(old)
   n=E.SubElement(dvs,T('dataValidation'),type='list',sqref=addr,showErrorMessage='1',errorStyle='stop');E.SubElement(n,T('formula1')).text='"'+','.join(str(x) for x in options)+'"';dvs.set('count',str(len(dvs)))
  # Existing controls outside mirrored panels now expose active value + a real link
  # to independent Base / Comparative inputs, avoiding ignored editable cells.
  for item in INPUTS:
   sn,addr=item['sheet'],item['cell']
   if sn in PANELS:continue
   s=sh(sn);key=IRO[(sn,addr)];target='IF(\'EXEC V06\'!$H$23="Comparative","#\''+CMP7+'\'!B'+str(ROWS[key])+'","#\''+BASE7+'\'!B'+str(ROWS[key])+'")';s.put(addr,item['value'],'HYPERLINK('+target+','+active(key)+')')
   dvs=s.r.find(T('dataValidations'))
   if dvs is not None:
    for node in list(dvs):
     if addr in node.get('sqref','').split():dvs.remove(node)
   changes[pp[sn]]=s.finish()
  for sn,panel in PANELS.items():
   s=sh(sn)
   s.text('E20:J21','BASE DE REFERÊNCIA — EDITÁVEL','REFERENCE BASE — EDITABLE','head');s.text('K20:P21','CENÁRIO COMPARATIVO — EDITÁVEL','COMPARATIVE SCENARIO — EDITABLE','head')
   # Clear obsolete short-panel validations and cells before installing mirrored controls.
   dvs=s.r.find(T('dataValidations'))
   if dvs is not None:
    for node in list(dvs):
     if any(20<=int(re.sub('[^0-9]','',a.split(':')[0]) or '0')<=39 for a in node.get('sqref','').split()):dvs.remove(node)
   for j,(key,pt,en,options) in enumerate(panel):
    rr=23+j*2;fullkey='bank_'+key if sn==BN else key;definition=next(d for d in DEFS if d['key']==fullkey);choice=initial_choice(definition)
    # V08 uses native numeric preset ranges and explicit decimal conversion.
    if key=='tan':choice='6.0%'
    for mode,lc,vc,end in [('Base','E','H','J'),('Comparative','K','N','P')]:
     s.text(f'{lc}{rr}:{chr(ord(vc)-1)}{rr+1}',pt,en,'label');s.merge(f'{vc}{rr}:{end}{rr+1}',choice,style='input_pct' if definition.get('kind')=='pct' else 'input');dv(s,vc+str(rr),options)
   s.link('E39:J40',BASE7,'A1','Custom, perfis e detalhes Base →','Custom, profiles and Base details →');s.link('K39:P40',CMP7,'A1','Custom e detalhes Comparativo →','Custom and Comparative details →')
   if sn=='DASH 05 V06':
    s.text('K41:M41','Ano de leitura','Display year','label');s.merge('N41:P41',5,style='input_number');dv(s,'N41',list(range(1,16)))
   changes[pp[sn]]=s.finish()
  s=sh('EXEC V06');s.put('H23','Base');dv(s,'H23',['Base','Comparative']);s.text('E23:G24','Resultados ativos','Active results','label');s.link('E38:J39',BASE7,'A1','Base de Referência editável →','Editable Reference Base →');s.link('K38:P39',CMP7,'A1','Comparativo independente →','Independent Comparative →');changes[pp['EXEC V06']]=s.finish()
  s=sh(IMP);s.put('B4','Current Base');s.put('C4','Selected Comparative')
  for rr in range(5,14):
   for cc in ['B','C']:s.put(cc+str(rr),0,q(MET7,cc+str(rr+1)))
  changes[pp[IMP]]=s.finish()
  s=sh(INV)
  for rr in range(220,401):
   original=source.get(INV,'B'+str(rr))[0]
   for cc,mode in [('B','Comparative'),('C','Base')]:s.put(cc+str(rr),0,rebound(original,mode))
  for rr in [26,27,28,31]:
   original=source.get(INV,'C'+str(rr))[0]
   for cc,mode in [('B','Base'),('C','Comparative')]:s.put(cc+str(rr),0,rebound(original,mode))
  for cc,mode,flowcol in [('B','Base','C'),('C','Comparative','B')]:s.put(cc+'30',0,'IF(SUM(N5:N11)>0,"N/A",NPV((1+'+inp(mode,'discount')+')^(1/12)-1,'+flowcol+'221:'+flowcol+'400)+'+flowcol+'220)')
  s.put('B24','Current Base');s.put('C24','Comparative');changes[pp[INV]]=s.finish()
  for j,(sn,*_) in enumerate(DASHES):
   s=sh(sn);s.put('E1','V07 · Editable Base / Independent Comparative · 2026-10-25',style='note');s.link('Q42:Y42',HIST7,'A1','Histórico e versões →','History and versions →');s.text('T24:U24','Base atual','Current Base','head');s.text('V24:W24','Comparativo','Comparative','head')
   if j==3:s.link('Q40:Y41',MET7,'A25','Comparação bancária completa →','Full financing comparison →')
   if j==9:
    s.text('E4:Y4','10 — Central de Relatórios Executivos','10 — Executive Reporting Center','section');s.link('E11:P13',REPORT7,'A1','Quatro formatos — configurar / emitir →','Four formats — configure / issue →');s.link('Q11:Y13',STATUS7,'A1','Estado Base e versão →','Base status and version →')
   changes[pp[sn]]=s.finish()
  # Always open Home; financial templates are technical and remain hidden.
  nodes=wb.find(T('sheets'));idx=next(i for i,n in enumerate(nodes) if n.get('name')=='EXEC V06')
  for view in wb.findall('.//'+T('workbookView')):view.set('activeTab',str(idx));view.set('firstSheet',str(idx))
  cp=wb.find(T('calcPr'))
  if cp is not None:cp.set('calcMode','auto');cp.set('fullCalcOnLoad','1');cp.set('forceFullCalc','1')
  changes['xl/workbook.xml']=xml(wb)
  # XlsxWriter modules and original workbook may each retain a selected tab.
  # Only the active executive sheet may remain selected after assembly.
  for sn,path in pp.items():
   root=E.fromstring(changes.get(path,z.read(path)));dirty=False
   for view in root.findall('.//'+T('sheetView')):
    selected='1' if sn=='EXEC V06' else '0'
    if view.get('tabSelected','0')!=selected:view.set('tabSelected',selected);dirty=True
   if dirty:changes[path]=xml(root)
 tmp=WORK7/"Patched.xlsx";rewrite(FINAL7,tmp,changes);tmp.replace(FINAL7)
if __name__=='__main__':build()
