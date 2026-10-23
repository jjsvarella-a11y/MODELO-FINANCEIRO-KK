"""R5: immutable R4 reference, effective parameter routing and decision modules.
No second engine. Existing financial formulas are retained except documented input
routing, construction curve weights, prices and collection timing extensions.
"""
from pathlib import Path
from zipfile import ZipFile,ZIP_DEFLATED
from lxml import etree as E
import xlsxwriter,json,re,copy,hashlib,sys
from xlsxwriter.utility import xl_col_to_name as col,xl_cell_to_rowcol as rc
from assemble import paths,T,xml,resolve,relpath,merge
from premium_ui import Source,DASHES
from premium_refine import Sheet,tr,cellref,bounds
from core import legacy
ROOT=Path(__file__).resolve().parents[1];WORK=Path('/workspace/kk01_efcc/closure_r5');WORK.mkdir(exist_ok=True)
BASE=ROOT/'Premium_Checkpoint/KK01_Executive_Financial_Control_Center_CHECKPOINT_R4.xlsx'
OUT=ROOT/'Fecho_R5';OUT.mkdir(exist_ok=True);MASTER=OUT/'KK01_Executive_Financial_Control_Center_R5_CHECKPOINT.xlsx'
REF='V06 Base R4';ACT='V06 Ativos R5';IMP='V06 Impact R5';CON='V06 Constru R5';CASH='V06 Cash R5';INV='V06 Investor R5';SALES='V06 Sales R5'
SN=[REF,ACT,IMP,CON,CASH,INV,SALES]
from reference_io import load_reference
CC=load_reference(ROOT)['main_columns'];sr=Source(BASE)
def q(sn,a):return f"'{sn}'!{a}"
def motor(key,a=5,b=184):return q('V06 Motor',f'${col(CC[key])}${a}:${col(CC[key])}${b}')
def val(sn,a):return sr.get(sn,a)[1]
INPUTS=[i for i in json.loads((ROOT/'Premium_Checkpoint/Parametros_Interface.json').read_text()) if i['sheet']!='DASH 10 Centro' and not i.get('presentation_only') and (i['sheet'],i['cell'])!=('DASH 05 V06','N23')]
def add(sn,a,v,pt,en,kind='number',options=None):INPUTS.append({'sheet':sn,'cell':a,'value':v,'label_pt':pt,'label_en':en,'kind':kind,'options':options})
add('DASH 02 V06','H27',1,'IVA recuperável','VAT recoverable','pct');add('DASH 02 V06','N27',1,'Multiplicador seguros','Insurance multiplier','ratio')
add('DASH 04A V06','N23',0,'Atraso recebimentos · meses','Collection delay · months');add('DASH 04A V06','H25','Base','Cenário de preços','Pricing scenario','text',['Conservative','Base','Favorable','Custom']);add('DASH 04A V06','N25',1,'Multiplicador Custom','Custom multiplier','ratio')
add('DASH 04B V06','H31',0,'Manutenção adicional / receita','Additional maintenance / revenue','pct');add('DASH 04B V06','H33','Base','Cenário residencial','Residential scenario','text',['Conservative','Base','Favorable','Custom'])
add('DASH 07 V06','N29',180,'Horizonte de saída · meses','Exit horizon · months','number',[120,180])
for j,ch in enumerate(legacy.CHAPTERS):
 add('DASH 01 V06','N'+str(194+j),0,'Ajuste capítulo '+str(j+1),'Chapter adjustment '+str(j+1),'pct');add('DASH 01 V06','T'+str(194+j),0,'ETC capítulo '+str(j+1),'Chapter ETC '+str(j+1),'money')
add('DASH 03 V06','H27',30,'Último mês de utilização','Last draw month')
add('DASH 03 V06','H29','Linear','Amortização','Amortisation','text',['Linear','Bullet'])
add('DASH 03 V06','H31',30000000,'Limite da linha €','Facility limit €','money')
MAPPING={(i['sheet'],i['cell']):q(ACT,'$B$'+str(n+5)) for n,i in enumerate(INPUTS)}
def a(sn,addr):return MAPPING[(sn,addr)]
def writef(s,row,c,f,fmt=None,v=0):s.write_formula(row-1,c,'='+f,fmt,v)
def create_modules():
 b=xlsxwriter.Workbook(WORK/'Modules.xlsx');fm={}
 defs={'text':{},'label':{'font_color':'687586'},'head':{'bold':True,'font_color':'FFFFFF','bg_color':'142337'},'section':{'bold':True,'bg_color':'E9EFF5','font_color':'142337','font_size':11},'base':{'font_color':'142337','num_format':'0.00,," M€";[Red](0.00,," M€");"–"'},'test':{'font_color':'1765A8','num_format':'0.00,," M€";[Red](0.00,," M€");"–"'},'delta':{'font_color':'1765A8','num_format':'+0.00,," M€";-0.00,," M€";"–"'},'pct':{'num_format':'0.0%'},'ratio':{'num_format':'0.00"×"'},'number':{'num_format':'#,##0'},'money':{'num_format':'#,##0 "€";[Red](#,##0) "€"'},'input':{'bg_color':'E6F0FA','font_color':'1765A8'},'input_pct':{'bg_color':'E6F0FA','font_color':'1765A8','num_format':'0.0%'},'input_money':{'bg_color':'E6F0FA','font_color':'1765A8','num_format':'#,##0 "€"'},'input_number':{'bg_color':'E6F0FA','font_color':'1765A8','num_format':'0.00'},'note':{'font_size':9,'font_color':'687586','bg_color':'F5F7FA'},'pp':{'num_format':'+0.0" pp";-0.0" pp";"–"'}}
 for name,d in defs.items():fm[name]=b.add_format({'font_name':'Aptos','font_size':9.5,'valign':'vcenter','text_wrap':True,**d})
 sheets={n:b.add_worksheet(n) for n in SN}
 for sn,s in sheets.items():
  s.hide_gridlines(2);s.freeze_panes(4,1);s.set_column('A:A',30);s.set_column('B:Z',18);s.write_formula('A1','=HYPERLINK("#\'EXEC V06\'!A1",'+tr('Início','Home')+')',fm['input'],'Início');s.write_formula('C1','=HYPERLINK("#\''+{'V06 Constru R5':'DASH 01 V06','V06 Cash R5':'DASH 03 V06','V06 Investor R5':'DASH 06 V06','V06 Sales R5':'DASH 04A V06'}.get(sn,'EXEC V06')+'\'!A1",'+tr('Voltar ao Dashboard','Back to Dashboard')+')',fm['input'],'Voltar ao Dashboard');s.write('A2',sn,fm['section'])
 # Style markers used after merge; no financial calculations here.
 for k,name in enumerate(fm):sheets[ACT].write(0,10+k,name,fm[name])
 s=sheets[REF];s.protect('KK01_reference',{'select_locked_cells':True});s.write_row(3,0,['Parameter','R4 reference value'])
 for n,i in enumerate(INPUTS,5):s.write(n-1,0,i['sheet']+'!'+i['cell']);s.write(n-1,1,i['value'])
 s.write_row(3,6,['Month','CAPEX','Cash','Reserve','Debt','VAT incurred','VAT refund','Dividends','DSCR','Debt service','Draw','Funding gap'])
 keys=['capex','cash','reserve','debt','vat_incurred','vat_refund','dividend','dscr','debt_service','draw','funding_gap']
 for m in range(1,181):
  s.write(m+3,6,m)
  for j,key in enumerate(keys,7):s.write(m+3,j,val('V06 Motor',col(CC[key])+str(m+4)))
 for j in range(15):
  for m in range(181):s.write(m+4,30+j,val('V06 Fluxos Invest',col(j+1)+str(m+5)))
 for j in range(15):
  for k,c in enumerate(['A','B','C','D','E','F','G','H','I','J','K','L']):
   value=val('V06 Blocos',c+str(j+5));s.write(209+j,k,value if value is not None else '')
 for j in range(94):s.write(239+j,0,val('V06 Unidades','A'+str(j+5)));s.write(239+j,1,val('V06 Unidades','H'+str(j+5)))
 s=sheets[ACT];s.write_row(3,0,['Parameter','Effective value','Test address','Reference value'])
 for n,i in enumerate(INPUTS,5):
  s.write(n-1,0,i['label_en']);s.write(n-1,2,i['sheet']+'!'+i['cell']);writef(s,n,3,q(REF,'B'+str(n)),None,i['value'])
  f='IF(\'EXEC V06\'!H23="Reference",'+q(REF,'B'+str(n))+','+q(i['sheet'],i['cell'])+')'
  if (i['sheet'],i['cell']) in [('DASH 04B V06','N27'),('DASH 04B V06','H29')]:
   inp=q(i['sheet'],i['cell']);mode=q('DASH 04B V06','H33');f='IF(\'EXEC V06\'!H23="Reference",'+q(REF,'B'+str(n))+',IF('+mode+'="Conservative",'+inp+'*0.9,IF('+mode+'="Favorable",'+inp+'*1.05,'+inp+')))'
  if (i['sheet'],i['cell'])==('DASH 04B V06','N29'):f='IF(\'EXEC V06\'!H23="Reference",'+q(REF,'B'+str(n))+',\'DASH 04B V06\'!N29+\'DASH 04B V06\'!H31)'
  writef(s,n,1,f,None,i['value'])
 # Budget chapter testing, native weighted construction curve, quarterly schedule.
 s=sheets[CON];s.write('A4','Adjusted construction');writef(s,5,1,'SUM(G12:G41)',fm['money'],56400000);writef(s,6,1,'SUM(H12:H41)+'+a('DASH 01 V06','H25'),fm['money']);writef(s,7,1,'B5*1.12+B6',fm['money'],63168000)
 s.write_row(10,0,['Code','Description','Reference budget','Share','Test adjustment','Adjustment EUR','Adjusted budget','Chapter ETC'])
 for j,ch in enumerate(legacy.CHAPTERS):
  rr=j+12;cost=56400000*ch['share'];s.write(rr-1,0,ch['id']);s.write(rr-1,1,ch['name']);s.write(rr-1,2,cost,fm['money']);s.write(rr-1,3,ch['share'],fm['pct']);writef(s,rr,4,a('DASH 01 V06','N'+str(194+j)),fm['pct']);writef(s,rr,5,'G'+str(rr)+'-C'+str(rr),fm['money']);writef(s,rr,6,'C'+str(rr)+'*(1+'+a('DASH 01 V06','N23')+')*(1+E'+str(rr)+')',fm['money'],cost);writef(s,rr,7,a('DASH 01 V06','T'+str(194+j)),fm['money'])
 s.write_row(44,0,['Quarter','Reference CAPEX','Test CAPEX','Reference cumulative','Test cumulative'])
 for j in range(60):
  rr=j+46;first=j*3+5;last=first+2;s.write(rr-1,0,j+1);s.write(rr-1,1,sum(val('V06 Motor',col(CC['capex'])+str(r)) or 0 for r in range(first,min(last,184)+1)),fm['money']);writef(s,rr,2,'SUM('+motor('capex',first,last)+')',fm['money']) if last<=184 else s.write(rr-1,2,0)
  writef(s,rr,3,'SUM(B46:B'+str(rr)+')',fm['money']);writef(s,rr,4,'SUM(C46:C'+str(rr)+')',fm['money'])
 s.write_row(108,0,['Bridge','Value']);s.write('A110','Base');s.write('B110',63168000);s.write('A111','Cost adjustments');writef(s,111,1,'B5*1.12-63168000');s.write('A112','ETC');writef(s,112,1,'B6');s.write('A113','Adjusted');writef(s,113,1,'B7')
 # Monthly liquidity diagnostics: data projections only, no independent financing engine.
 s=sheets[CASH];cols=['month','cash','reserve','debt','debt_service','draw','vat_incurred','vat_refund','vat_credit','funding_gap','dscr','liquid_headroom','coverage_headroom','dividend'];s.write_row(3,0,cols+['Reserve headroom','DSCR for min'])
 for m in range(1,181):
  rr=m+4
  for j,key in enumerate(cols):writef(s,rr,j,q('V06 Motor',col(CC[key])+str(rr)))
  writef(s,rr,14,'B'+str(rr)+'-C'+str(rr));writef(s,rr,15,'IF(E'+str(rr)+'>0,K'+str(rr)+',1E99)')
 s.write_row(187,0,['Critical point','Month','Value'])
 cases=[('Minimum cash','B','MIN'),('Minimum reserve headroom','O','MIN'),('Peak debt','D','MAX'),('Maximum funding gap','J','MAX'),('Minimum DSCR with service','P','MIN'),('Maximum VAT credit','I','MAX')]
 for i,(label,c,fn) in enumerate(cases,190):
  s.write(i-1,0,label);f=fn+'('+c+'5:'+c+'184)';writef(s,i,2,'IF('+f+'=1E99,"N/A",'+f+')');writef(s,i,1,'IF(C'+str(i)+'="N/A","N/A",MATCH(C'+str(i)+','+c+'5:'+c+'184,0))')
 # Single source of project comparison metrics; reference is the R4 snapshot.
 s=sheets[IMP];s.write_row(3,0,['KPI','Reference R4','Active / Test','Delta','Relative delta'])
 metrics=[('Future CAPEX','Investimento remanescente','Future CAPEX','SUM('+motor('capex')+')',sum(val('V06 Motor',col(CC['capex'])+str(r)) for r in range(5,185)),'money'),('Peak Debt','Pico de dívida','Peak debt','MAX('+motor('debt')+')',val('V06 Cenarios','F5'),'money'),('Liquidity','Folga mínima sobre reservas','Minimum reserve headroom','MIN(\''+CASH+'\'!O5:O184)',min(val('V06 Motor',col(CC['cash'])+str(r))-val('V06 Motor',col(CC['reserve'])+str(r)) for r in range(5,185)),'money'),('Dividends','Dividendos previstos','Planned dividends','SUM('+motor('dividend')+')',sum(val('V06 Motor',col(CC['dividend'])+str(r)) for r in range(5,185)),'money'),('Investor NPV','VAL novos investidores','New investor NPV',q('V06 Cenarios','M5'),val('V06 Cenarios','M5'),'money'),('Investor IRR','TIR novos investidores','New investor IRR',q('V06 Cenarios','K5'),val('V06 Cenarios','K5'),'pct'),('MOIC','MOIC','MOIC',q('V06 Cenarios','L5'),val('V06 Cenarios','L5'),'ratio'),('Exit Equity','Equity na saída','Exit equity',q('V06 Cenarios','J5'),val('V06 Cenarios','J5'),'money'),('M54 Capacity','Dividendos no mês 54','M54 dividends',q('V06 Cenarios','E5'),val('V06 Cenarios','E5'),'money')]
 for r,(key,pt,en,f,v,kind) in enumerate(metrics,5):
  writef(s,r,0,tr(pt,en),fm['label'],pt);s.write(r-1,1,v,fm['pct' if kind=='pct' else 'ratio' if kind=='ratio' else 'base']);writef(s,r,2,f,fm['pct' if kind=='pct' else 'ratio' if kind=='ratio' else 'test'],v);writef(s,r,3,'C'+str(r)+'-B'+str(r),fm['pct' if kind=='pct' else 'ratio' if kind=='ratio' else 'delta']);writef(s,r,4,'IF(ABS(B'+str(r)+')>0.000001,D'+str(r)+'/ABS(B'+str(r)+'),"N/A")',fm['pct'])
 writef(s,16,1,'IF(MAX('+motor('funding_gap')+')>0.005,'+tr('Défice de financiamento','Funding shortfall')+',IF(MIN(\''+CASH+'\'!O5:O184)<-0.005,'+tr('Reserva não coberta','Reserve shortfall')+',IF(C9<0,'+tr('VAL negativo','Negative investor NPV')+','+tr('Sem défice modelado','No modelled shortfall')+')))',fm['note'],'VAL negativo')
 writef(s,18,1,'IF(D9>0,'+tr('VAL melhorou','NPV improved')+','+tr('VAL não melhorou','NPV did not improve')+')&IF(C9<0,'+tr(' mas permanece negativo.',' but remains negative.')+',".")&" "&IF(MAX('+motor('funding_gap')+')>0.005,'+tr('Liquidez insuficiente.','Insufficient liquidity.')+','+tr('Sem défice de financiamento modelado.','No modelled funding shortfall.')+')',fm['note'],'VAL permanece negativo.');s.set_row(17,42)
 # Individual block and group comparisons: use existing per-block dated flows.
 s=sheets[INV];s.write_row(3,0,['Block','Owner','Class','Share','Equivalent','New ticket','Status','Paid actual','Unpaid actual','Proposed rights / terms'])
 for j in range(15):
  r=j+5
  for c in range(4):writef(s,r,c,q('V06 Blocos',col(c)+str(r)),None,val('V06 Blocos',col(c)+str(r)))
  writef(s,r,4,'D'+str(r)+'*50000000',fm['money']);ticket=0 if j<7 else val('V06 Capital','C'+str(j-7+5));s.write(r-1,5,ticket,fm['money']);s.write(r-1,6,'Acquired rights' if j<7 else 'Scenario subscription; confirmation pending')
  if j<7:s.write(r-1,7,0);s.write(r-1,8,0)
  else:writef(s,r,7,'MIN(F'+str(r)+',N(\'V06 Capital\'!H'+str(j-7+5)+'))',fm['money']);writef(s,r,8,'MAX(0,F'+str(r)+'-H'+str(r)+')',fm['money'])
  s.write(r-1,9,'Economic share; governance/info/access terms require documentary approval')
  for off,sn,cls,block,group in [(13,'DASH 06 V06','H25','N25','H27'),(14,'DASH 07 V06','H31','N31','H33')]:
   groupf=q(sn,group);mask='AND(OR('+q(sn,cls)+'="All",'+q(sn,cls)+'=C'+str(r)+'),OR('+q(sn,block)+'="All",'+q(sn,block)+'=A'+str(r)+'),OR('+groupf+'="All",AND('+groupf+'="New",B'+str(r)+'="New"),AND('+groupf+'="Promoters",B'+str(r)+'<>"New")))';writef(s,r,off,'IF('+mask+',1,0)')
 for case,maskcol in enumerate(['N','O']):
  start=0 if case==0 else 12
  s.write_row(24,start,['Metric','Reference','Active','Delta'])
  for m in range(181):
   rr=m+220;s.write(rr-1,0,m)
   for refcase,cc in [(False,1+case*4),(True,2+case*4)]:
    terms=[f'{maskcol}{j+5}*'+(q(REF,col(30+j)+str(m+5)) if refcase else q('V06 Fluxos Invest',col(j+1)+str(m+5))) for j in range(15)];writef(s,rr,cc,'SUM('+','.join(terms)+')')
   writef(s,rr,25+case*2,('0+' if m==0 else col(25+case*2)+str(rr-1)+'+')+col(2+case*4)+str(rr));writef(s,rr,26+case*2,'IF(AND('+col(25+case*2)+str(rr)+'>=0,MIN('+col(2+case*4)+'220:'+col(2+case*4)+str(rr)+')<0),A'+str(rr)+',1E99)')
   writef(s,rr,3+case*4,('0+' if m==0 else col(3+case*4)+str(rr-1)+'+')+col(1+case*4)+str(rr));writef(s,rr,4+case*4,'IF(AND('+col(3+case*4)+str(rr)+'>=0,SUM('+col(1+case*4)+'220:'+col(1+case*4)+str(rr)+')>=0,MIN('+col(1+case*4)+'220:'+col(1+case*4)+str(rr)+')<0),A'+str(rr)+',1E99)')
  selected='SUM('+maskcol+'5:'+maskcol+'19)';prom='SUM('+maskcol+'5:'+maskcol+'11)';nonnew=prom+'>0';cap='SUMPRODUCT($F$5:$F$19,$'+maskcol+'$5:$'+maskcol+'$19)'
  sums=lambda c,base=False:'SUM('+','.join(maskcol+str(j+5)+'*'+(q(REF,col(c)+str(210+j)) if base else q('V06 Blocos',col(c)+str(j+5))) for j in range(15))+')'
  for k,(label,kind) in enumerate([('Capital invested','money'),('Share','pct'),('Dividends','money'),('IRR','pct'),('NPV','money'),('MOIC','ratio'),('Cash-on-cash annual','pct'),('Payback month','number'),('Exit proceeds','money'),('Return dependent on exit','pct'),('Capital returned via dividends','pct')],26):
   s.write(k-1,start,label,fm['label'])
   for baseflag,c in [(True,start+1),(False,start+2)]:
    invested=sums(4,baseflag);div=sums(5,baseflag);exit=sums(6,baseflag);flows=col((2 if baseflag else 1)+case*4)+'220:'+col((2 if baseflag else 1)+case*4)+'400';dc=q('V06 Base R4','B'+str(next(n+5 for n,i in enumerate(INPUTS) if (i['sheet'],i['cell'])==('DASH 07 V06','N25')))) if baseflag else q('V06 Registo Ativo','B65');horizon='180' if baseflag else q('V06 Controlos','B64')
    fs={'Capital invested':invested,'Share':sums(3,baseflag),'Dividends':div,'IRR':'IF(AND('+prom+'=0,'+selected+'>0,MIN('+flows+')<0,MAX('+flows+')>0),(1+IRR('+flows+',0.005))^12-1,"N/A")','NPV':'IF('+nonnew+',"N/A",NPV((1+'+dc+')^(1/12)-1,'+col((2 if baseflag else 1)+case*4)+'221:'+col((2 if baseflag else 1)+case*4)+'400)+'+col((2 if baseflag else 1)+case*4)+'220)','MOIC':'IF('+invested+'>0,('+div+'+'+exit+')/'+invested+',"N/A")','Cash-on-cash annual':'IF('+invested+'>0,'+div+'/'+invested+'/('+horizon+'/12),"N/A")','Payback month':'IF('+nonnew+',"N/A",'+('IF(MIN('+col(26+case*2)+'220:'+col(26+case*2)+'400)=1E99,"Not reached",MIN('+col(26+case*2)+'220:'+col(26+case*2)+'400))' if baseflag else 'IF(MIN('+col(4+case*4)+'220:'+col(4+case*4)+'400)=1E99,"Not reached",MIN('+col(4+case*4)+'220:'+col(4+case*4)+'400))')+')','Exit proceeds':exit,'Return dependent on exit':'IF(('+div+'+'+exit+')>0,'+exit+'/('+div+'+'+exit+'),0)','Capital returned via dividends':'IF('+invested+'>0,'+div+'/'+invested+',"N/A")'}
    formula=fs[label]
    if label in ['Capital invested','MOIC','Cash-on-cash annual','Capital returned via dividends']:formula='IF('+nonnew+',"N/A",'+formula+')'
    writef(s,k,c,formula,fm[kind])
   writef(s,k,start+3,'IF(AND(ISNUMBER('+col(start+1)+str(k)+'),ISNUMBER('+col(start+2)+str(k)+')),'+col(start+2)+str(k)+'-'+col(start+1)+str(k)+',"N/A")',fm[kind])
 # Unit pricing and inventory filters (analytical filters never alter legal sale status).
 s=sheets[SALES];s.write_row(3,0,['ID','Type','Floor','Area','Price','EUR/m2','Forecast status','Matches filters'])
 for j in range(94):
  rr=j+5
  for c,source in enumerate(['A','B','C','D','H']):writef(s,rr,c,q('V06 Unidades',source+str(rr)),None,val('V06 Unidades',source+str(rr)))
  writef(s,rr,5,'IF(D'+str(rr)+'>0,E'+str(rr)+'/D'+str(rr)+',0)');writef(s,rr,6,'IF(\'DASH 04A V06\'!H35>=\'V06 Unidades\'!L'+str(rr)+',"Deeded",IF(\'DASH 04A V06\'!H35>=\'V06 Unidades\'!I'+str(rr)+',"Reserved","Available"))')
  conditions=['OR(\'DASH 04A V06\'!H27="All",\'DASH 04A V06\'!H27=B'+str(rr)+')','OR(\'DASH 04A V06\'!N27="All",\'DASH 04A V06\'!N27=C'+str(rr)+')','OR(\'DASH 04A V06\'!H29="All",\'DASH 04A V06\'!H29=G'+str(rr)+')','D'+str(rr)+'>=\'DASH 04A V06\'!N29','D'+str(rr)+'<=\'DASH 04A V06\'!H31','F'+str(rr)+'>=\'DASH 04A V06\'!H33','F'+str(rr)+'<=\'DASH 04A V06\'!N33'];writef(s,rr,7,'IF(AND('+','.join(conditions)+'),1,0)')
 types=sorted({str(val('V06 Unidades','B'+str(r))) for r in range(5,99)})
 s.write_row(102,0,['Type','Units','Area','Average area','EUR/m2','Average unit price','GDV','Available','Commercialised'])
 for j,typ in enumerate(types,105):
  s.write(j-1,0,typ);writef(s,j,1,'COUNTIF(B5:B98,A'+str(j)+')');writef(s,j,2,'SUMIF(B5:B98,A'+str(j)+',D5:D98)');writef(s,j,3,'IF(B'+str(j)+'>0,C'+str(j)+'/B'+str(j)+',0)');writef(s,j,6,'SUMIF(B5:B98,A'+str(j)+',E5:E98)');writef(s,j,4,'IF(C'+str(j)+'>0,G'+str(j)+'/C'+str(j)+',0)');writef(s,j,5,'IF(B'+str(j)+'>0,G'+str(j)+'/B'+str(j)+',0)');writef(s,j,7,'COUNTIFS(B5:B98,A'+str(j)+',G5:G98,"Available")');writef(s,j,8,'B'+str(j)+'-H'+str(j))
 create_charts(b,sheets,fm,types)
 b.close();return metrics,types
def create_charts(b,sheets,fm,types):
 def chart(sheet,kind,series,cell,stacked=False,y2=False):
  opts={'type':kind}
  if stacked:opts['subtype']='stacked'
  c=b.add_chart(opts)
  for i,(name,cat,values) in enumerate(series):
   args={'name':name,'categories':'='+cat,'values':'='+values}
   if y2 and i==1:args['y2_axis']=True
   if kind!='doughnut':args['line']={'color':['142337','1765A8','C8101E'][i%3]}
   c.add_series(args)
  c.set_legend({'position':'bottom','font':{'size':9}});c.set_size({'width':400,'height':187});c.set_chartarea({'border':{'none':True},'fill':{'color':'FFFFFF'}})
  if y2:c.set_y2_axis({'num_format':'0%','num_font':{'size':9}})
  sheets[sheet].insert_chart(cell,c)
 chart(CON,'column',[('Reference',q(CON,'$A$12:$A$41'),q(CON,'$C$12:$C$41')),('Test',q(CON,'$A$12:$A$41'),q(CON,'$G$12:$G$41'))],'K5')
 chart(CON,'column',[('Reference',q(CON,'$A$46:$A$65'),q(CON,'$B$46:$B$65')),('Test',q(CON,'$A$46:$A$65'),q(CON,'$C$46:$C$65'))],'K20')
 chart(CON,'column',[('Budget bridge',q(CON,'$A$110:$A$113'),q(CON,'$B$110:$B$113'))],'K35')
 chart(CASH,'line',[('Cash',q(CASH,'$A$5:$A$184'),q(CASH,'$B$5:$B$184')),('Reserves',q(CASH,'$A$5:$A$184'),q(CASH,'$C$5:$C$184'))],'S5')
 chart(CASH,'line',[('Debt',q(CASH,'$A$5:$A$184'),q(CASH,'$D$5:$D$184')),('Debt service',q(CASH,'$A$5:$A$184'),q(CASH,'$E$5:$E$184'))],'S20')
 for rr in range(5,185):
  src='SUM('+','.join(q('V06 Motor',col(CC[key])+str(rr)) for key in ['equity','receipts','draw','vat_refund'])+')+MAX(0,'+q('V06 Motor',col(CC['cfads'])+str(rr))+'+'+q('V06 Motor',col(CC['res_kk'])+str(rr))+')'
  writef(sheets[CASH],rr,17,src);writef(sheets[CASH],rr,18,'R'+str(rr)+'-('+q('V06 Motor',col(CC['cash'])+str(rr))+'-'+q('V06 Motor',col(CC['cash_open'])+str(rr))+')')
 chart(CASH,'column',[('Cash sources',q(CASH,'$A$5:$A$184'),q(CASH,'$R$5:$R$184')),('Cash uses',q(CASH,'$A$5:$A$184'),q(CASH,'$S$5:$S$184'))],'S35')
 chart(CASH,'line',[('Reserve headroom',q(CASH,'$A$5:$A$184'),q(CASH,'$O$5:$O$184'))],'S50')
 s=sheets[INV]
 for i,(name,share) in enumerate([('JSCV',.4),('JALS',.1),('New',.5)],50):s.write(i-1,0,name);s.write(i-1,1,share)
 for i,(name,n) in enumerate([('Platinum',3),('Gold',2),('Silver',10)],55):s.write(i-1,0,name);s.write(i-1,1,n)
 s.write('A60','Planned new equity');s.write('B60',25000000);s.write('A61','Documented paid');writef(s,61,1,'SUM(H12:H19)');s.write('A62','Unpaid');writef(s,62,1,'SUM(I12:I19)')
 chart(INV,'doughnut',[('Ownership',q(INV,'$A$50:$A$52'),q(INV,'$B$50:$B$52'))],'R5')
 chart(INV,'column',[('Blocks',q(INV,'$A$55:$A$57'),q(INV,'$B$55:$B$57'))],'R20')
 chart(INV,'column',[('Capital',q(INV,'$A$60:$A$62'),q(INV,'$B$60:$B$62'))],'R35')
 s=sheets[SALES]
 for year in range(1,16):
  rr=149+year;s.write(rr-1,0,year)
  for c,key in [(1,'deed_sales'),(2,'receipts')]:writef(s,rr,c,'SUM('+motor(key,(year-1)*12+5,year*12+4)+')')
  writef(s,rr,3,'SUM(B150:B'+str(rr)+')');writef(s,rr,4,'SUM(C150:C'+str(rr)+')')
 chart(SALES,'column',[('Deeds',q(SALES,'$A$150:$A$164'),q(SALES,'$B$150:$B$164'))],'K5')
 chart(SALES,'line',[('Deeds cumulative',q(SALES,'$A$150:$A$164'),q(SALES,'$D$150:$D$164')),('Collections cumulative',q(SALES,'$A$150:$A$164'),q(SALES,'$E$150:$E$164'))],'K20')
 chart(SALES,'column',[('Available',q(SALES,'$A$105:$A$'+str(104+len(types))),q(SALES,'$H$105:$H$'+str(104+len(types))))],'K35')

def patch():
 metrics,types=create_modules();combined=WORK/'Combined.xlsx';merge(BASE,WORK/'Modules.xlsx',combined)
 changed_formula_count={};extensions={};styles={};new_inputs=[];display=[]
 with ZipFile(combined) as z:
  wb,pp=paths(z);roots={};changes={};rootnames={p:sn for sn,p in pp.items()}
  def root(sn):
   if sn not in roots:roots[sn]=E.fromstring(z.read(pp[sn]))
   return roots[sn]
  def rewrite(f):
   if 'HYPERLINK(' in f:return f
   return re.sub(r"'([^']+)'!\$?([A-Z]+)\$?([0-9]+)(?![0-9])",lambda m:MAPPING.get((m[1],m[2]+m[3]),m[0]),f)
  oldrate=sr.get('V06 Registo Ativo','B65')[0]
  for sn,p in pp.items():
   if sn in SN:continue
   r=root(sn);n=0
   for f in r.iter(T('f')):
    old=f.text or '';new=old.replace(oldrate,q('V06 Registo Ativo','$B$65')) if 'NPV(' in old else old;new=rewrite(new)
    if old!=new:f.text=new;n+=1
   if n:changed_formula_count[sn]=n
  # New styles are copied from already-merged marker cells, keeping correct indices.
  style_names=list(['text','label','head','section','base','test','delta','pct','ratio','number','money','input','input_pct','input_money','input_number','note','pp'])
  cells={c.get('r'):c for c in root(ACT).iter(T('c'))}
  for j,name in enumerate(style_names):styles[name]=int(cells[col(10+j)+'1'].get('s'))
  def sh(sn):return Sheet(root(sn),styles)
  def vcache(sn,addr):return val(sn,addr) if sn in sr.pp else 0
  def put(s,area,formula,cache=0,kind='text'):s.merge(area,cache,formula,kind);display.append({'sheet':s.name,'cell':area.split(':')[0],'formula':formula})
  def label(s,area,pt,en,kind='label'):s.text(area,pt,en,kind)
  def link(s,area,sn,addr,pt,en):s.merge(area,pt,'HYPERLINK("#\''+sn+'\'!'+addr+'",'+tr(pt,en)+')','input')
  def ui(sn):s=sh(sn);s.name=sn;return s
  def validation(s,area,kind,values=None,minimum=0,maximum=1):
   dvs=s.r.find(T('dataValidations'))
   if dvs is None:dvs=E.SubElement(s.r,T('dataValidations'),count='0')
   for old in list(dvs):
    if old.get('sqref') in [area,area.split(':')[0]]:dvs.remove(old)
   dv=E.SubElement(dvs,T('dataValidation'),type=kind,sqref=area,allowBlank='0',showErrorMessage='1',errorTitle='Invalid input',error='Enter a valid value.')
   if kind=='list':E.SubElement(dv,T('formula1')).text='"'+','.join(map(str,values))+'"'
   else:dv.set('operator','between');E.SubElement(dv,T('formula1')).text=str(minimum);E.SubElement(dv,T('formula2')).text=str(maximum)
   dvs.set('count',str(len(dvs)))
  def editable(s,area,value,kind='input',options=None,lo=0,hi=1e9):
   s.merge(area,value,style=kind)
   validation(s,area.split(':')[0],'list' if options else 'decimal',options,lo,hi)
   new_inputs.append({'sheet':s.name,'cell':area.split(':')[0],'value':value,'options':options})
  # Active VAT recovery, insurance and horizon; no financial-algorithm duplication.
  for addr,sn,source in [('B7','DASH 02 V06','H27'),('B33','DASH 02 V06','N27'),('B64','DASH 07 V06','N29')]:sh('V06 Controlos').put(addr,vcache('V06 Controlos',addr),a(sn,source))
  sh('V06 Controlos').put('B39',0,q(CON,'B5')+'/56400000-1');sh('V06 Controlos').put('B75',0,q(CON,'B6'))
  for j in range(30):
   formula='SUM('+','.join(q(CON,'$G$'+str(k+12))+'*'+str(ch['weights'][j]) for k,ch in enumerate(legacy.CHAPTERS))+')/'+q(CON,'$B$5')
   sh('V06 Curva').put('B'+str(j+5),val('V06 Curva','B'+str(j+5)),formula)
  # Unit prices: immutable reference plus explicit pricing multiplier in test mode.
  price_mode=a('DASH 04A V06','H25');factor='IF('+price_mode+'="Conservative",0.9,IF('+price_mode+'="Favorable",1.1,IF('+price_mode+'="Custom",'+a('DASH 04A V06','N25')+',1)))'
  units=sh('V06 Unidades')
  for j in range(94):
   addr='H'+str(j+5);old,v=sr.get('V06 Unidades',addr);original=old or str(v);f='IF(\'EXEC V06\'!H23="Reference",'+q(REF,'B'+str(j+240))+',('+original+')*'+factor+')';units.put(addr,v,f)
  # Collections use original unit calendar, independently from sale/deed dates.
  lag=a('DASH 04A V06','N23');collection_changes=0
  for sn in pp:
   if sn!='V06 Motor' and not re.fullmatch(r'V06 [SMO]\d+',sn):continue
   s=sh(sn)
   for rr in range(5,185):
    c=s.cells.get('O'+str(rr));old=c.findtext(T('f'));m=re.search(r"SUM\('([^']+)'!([A-Z]+)5:[A-Z]+98\)",old or '')
    if not m:raise ValueError((sn,rr,old))
    calendar=m[1];rg=q(calendar,'$M$5:$'+col(191)+'$98');month=rr-4;f='IF('+str(month)+'>'+lag+',SUM(INDEX('+rg+',1,'+str(month)+'-'+lag+'):INDEX('+rg+',94,'+str(month)+'-'+lag+')),0)';s.put('O'+str(rr),vcache(sn,'O'+str(rr)),f);collection_changes+=1
  # Prevent report selector from falling back to the obsolete historical reference.
  centre=ui('DASH 10 Centro');centre.put('N11','Active Scenario','IF(\'EXEC V06\'!H23="Reference","Reference R4","Active Scenario")')
  dvs=centre.r.find(T('dataValidations'))
  if dvs is not None:
   for d in list(dvs):
    if d.get('sqref')=='N11':dvs.remove(d)
   dvs.set('count',str(len(dvs)))
  # Every dashboard: compact impact beside inputs, full reconciliation below.
  for index,(sn,pt,en) in enumerate(DASHES):
   s=ui(sn);s.merge('Q22:Y39',None,style='note');label(s,'Q22:Y23','Impacto Global do Cenário','Project-Wide Scenario Impact','section')
   for ar,pt2,en2 in [('Q24:S25','Indicador','KPI'),('T24:U25','Referência R4','Reference R4'),('V24:W25','Ativo / Teste','Active / Test'),('X24:Y25','Variação','Delta')]:label(s,ar,pt2,en2,'head')
   for j in range(6):
    rr=26+j;mr=5+j;kind=metrics[j][-1];put(s,f'Q{rr}:S{rr}',q(IMP,'A'+str(mr)),metrics[j][1],'label')
    for ar,c,sty in [(f'T{rr}:U{rr}','B','base'),(f'V{rr}:W{rr}','C','test'),(f'X{rr}:Y{rr}','D','delta')]:put(s,ar,q(IMP,c+str(mr)),metrics[j][4] if c!='D' else 0,'pct' if kind=='pct' else sty)
   put(s,'Q33:Y34',q(IMP,'B16'),'VAL negativo','note');link(s,'Q35:Y36','EXEC V06','H23','Modo Base/Teste →','Reference/Test mode →');put(s,'Q37:Y39',q(IMP,'B18'),'VAL permanece negativo.','note')
   label(s,'E150:Y151','Impacto Global do Cenário no Projeto — Base vs Teste','Project-Wide Scenario Impact — Reference vs Test','section')
   for ar,pt2,en2 in [('E152:L153','Indicador','KPI'),('M152:O153','Referência R4','Reference R4'),('P152:R153','Ativo / Teste','Active / Test'),('S152:U153','Delta absoluto / pp','Absolute delta / pp'),('V152:Y153','Delta relativo','Relative delta')]:label(s,ar,pt2,en2,'head')
   for j,metric in enumerate(metrics):
    rr=154+j*2;mr=5+j;kind=metric[-1];put(s,f'E{rr}:L{rr+1}',q(IMP,'A'+str(mr)),metric[1],'label')
    for ar,c,sty in [(f'M{rr}:O{rr+1}','B','base'),(f'P{rr}:R{rr+1}','C','test'),(f'S{rr}:U{rr+1}','D','delta'),(f'V{rr}:Y{rr+1}','E','pct')]:
     f=q(IMP,c+str(mr));style='pct' if c=='E' or kind=='pct' else 'ratio' if kind=='ratio' else sty
     if kind=='pct' and c=='D':f+='*100';style='pp'
     put(s,ar,f,metric[4] if c in ['B','C'] else 0,style)
   put(s,'E173:Y176',q(IMP,'B18'),'VAL permanece negativo.','note');link(s,'E178:L179','EXEC V06','H23','Escolher Base / Teste','Select Reference / Test');link(s,'M178:Y179',sn,'E22','Voltar aos filtros','Back to filters')
   for rr in range(150,180):s.row(rr).set('ht','17');s.row(rr).set('customHeight','1')
   # Replace unrelated hotel assumptions in generic tables with topic-specific metrics.
   if index!=6:
    theme=[('CAPEX','Future CAPEX',5),('Dívida','Debt',6),('Liquidez','Liquidity',7),('Dividendos','Dividends',8),('VAL','NPV',9)]
    for j,(pt2,en2,mr) in enumerate(theme):
     rr=92+j*2;label(s,f'S{rr}:W{rr+1}',pt2,en2);put(s,f'X{rr}:Y{rr+1}',q(IMP,'C'+str(mr)),metrics[mr-5][4],'money')
   # Release stamp distinguishes the controlled workbook without changing chart geometry.
   s.put('E1','R5 CHECKPOINT · base R4 · 2026-10-23',style='note')
  # Active scenario switch is an input; reference snapshot is never overwritten.
  s=ui('EXEC V06');label(s,'E23:G24','Modo global','Global mode');editable(s,'H23:J24','Test','input',['Reference','Test']);label(s,'K23:P26','Reference restaura os pressupostos R4; Test usa os filtros editáveis. Os valores de teste são preservados ao alternar.','Reference restores R4 assumptions; Test uses editable filters. Stored test inputs survive switching.','note');link(s,'E28:J29','DASH 03 V06','E190','Meses críticos de tesouraria','Critical liquidity months');link(s,'K28:P29','DASH 01 V06','E190','Orçamento por capítulos','Budget by chapter');put(s,'E31:P34',q(IMP,'B16'),'VAL negativo','section');label(s,'E36:P38','Referência R4 fixa. Controlos azuis editam o teste; os resultados usam o modo ativo.','Fixed R4 reference. Blue controls edit the test; results follow the active mode.','note')
  # New financial controls remain on their responsible dashboards.
  for i in INPUTS:
   if 'kind' not in i or int(re.search(r'\d+',i['cell'])[0])>=190:continue
   s=ui(i['sheet']);rr,cc=rc(i['cell']);r=rr+1;c=cc
   label(s,f'{col(c-3)}{r}:{col(c-1)}{r}',i['label_pt'],i['label_en']);kind='input_pct' if i['kind']=='pct' else 'input_money' if i['kind']=='money' else 'input_number' if i['kind'] in ['number','ratio'] else 'input';editable(s,f'{col(c)}{r}:{col(c+2)}{r}',i['value'],kind,i.get('options'),0,1000000000 if i['kind']=='money' else 1000)
  # Clearly explain discount preset versus stored custom value.
  s=ui('DASH 07 V06');label(s,'E29:G29','Desconto: 8%, 10%, 12%, Custom','Discount: 8%, 10%, 12%, Custom');label(s,'K25:M25','Custom: só quando selecionado','Custom: used only when selected');label(s,'K26:M26','Desconto anual efetivo','Effective annual discount');put(s,'N26:P26',q('V06 Registo Ativo','B65'),.12,'pct');label(s,'E36:P38','VAL usa (1+taxa anual)^(1/12)−1. A taxa de desconto não altera a TIR.','NPV uses (1+annual rate)^(1/12)−1. Discount changes do not change IRR.','note')
  # Detailed construction table and quarterly comparison, directly editable.
  s=ui('DASH 01 V06');link(s,'E38:P38',s.name,'E190','Orçamento, capítulos e trimestres →','Budget, chapters and quarters →');label(s,'E190:Y191','Orçamento sintético da construção — testes por capítulo','Construction budget — chapter tests','section')
  headings=[('E192:E193','Código','Code'),('F192:J193','Capítulo','Chapter'),('K192:L193','Base €','Reference €'),('M192:M193','Peso','Share'),('N192:N193','Ajuste %','Adjust %'),('O192:P193','Ajuste €','Adjust €'),('Q192:S193','Ajustado €','Adjusted €'),('T192:V193','ETC capítulo €','Chapter ETC €'),('W192:Y193','Total €','Total €')]
  for ar,pt,en in headings:label(s,ar,pt,en,'head')
  for j,ch in enumerate(legacy.CHAPTERS):
   rr=194+j;mr=12+j;put(s,f'E{rr}:E{rr}',q(CON,'A'+str(mr)),ch['id'],'number');put(s,f'F{rr}:J{rr}',q(CON,'B'+str(mr)),ch['name']);put(s,f'K{rr}:L{rr}',q(CON,'C'+str(mr)),56400000*ch['share'],'money');put(s,f'M{rr}:M{rr}',q(CON,'D'+str(mr)),ch['share'],'pct');editable(s,f'N{rr}:N{rr}',0,'input_pct',None,-.9,3);put(s,f'O{rr}:P{rr}',q(CON,'F'+str(mr)),0,'money');put(s,f'Q{rr}:S{rr}',q(CON,'G'+str(mr)),56400000*ch['share'],'money');editable(s,f'T{rr}:V{rr}',0,'input_money',None,0,1e9);put(s,f'W{rr}:Y{rr}',q(CON,'G'+str(mr))+'+'+q(CON,'H'+str(mr)),56400000*ch['share'],'money');s.row(rr).set('ht','29');s.row(rr).set('customHeight','1')
  label(s,'E225:P227','Os capítulos incluem construção direta. Acrescem 2% de fiscalização e 10% de contingência. ETC é adicional; não representa execução comprovada.','Chapters contain direct construction. Add 2% supervision and 10% contingency. ETC is incremental, not documented actual expenditure.','note');put(s,'Q225:Y227',q(CON,'B7'),63168000,'money');label(s,'E230:Y231','Desembolso trimestral — Base vs Teste e acumulado','Quarterly disbursement — Reference vs Test and cumulative','section')
  for j in range(60):
   rr=234+j;mr=46+j
   for ar,c,kind in [(f'E{rr}:G{rr}','A','number'),(f'H{rr}:K{rr}','B','base'),(f'L{rr}:O{rr}','C','test'),(f'P{rr}:T{rr}','D','base'),(f'U{rr}:Y{rr}','E','test')]:put(s,ar,q(CON,c+str(mr)),0,kind)
  for ar,pt,en in [('E232:G233','Trimestre','Quarter'),('H232:K233','Base','Reference'),('L232:O233','Teste','Test'),('P232:T233','Base acumulado','Reference cumulative'),('U232:Y233','Teste acumulado','Test cumulative')]:label(s,ar,pt,en,'head')
  # Critical liquidity points, full monthly timeline, no unsupported DSCR covenant claim.
  s=ui('DASH 03 V06');link(s,'E38:P38',s.name,'E190','Meses críticos e liquidez →','Critical months and liquidity →');label(s,'E190:Y191','Meses críticos de tesouraria','Critical liquidity months','section')
  for j in range(6):
   rr=194+j*2;mr=190+j;put(s,f'E{rr}:P{rr+1}',q(CASH,'A'+str(mr)),'Critical point');put(s,f'Q{rr}:S{rr+1}',q(CASH,'B'+str(mr)),0,'number');put(s,f'T{rr}:Y{rr+1}',q(CASH,'C'+str(mr)),0,'ratio' if j==4 else 'money')
  label(s,'E208:Y211','DSCR apenas nos meses com serviço de dívida; ausência de serviço não é DSCR zero. Não foi introduzido um covenant bancário sem documentação. IVA: pico do crédito e reembolso mostram a pressão temporal.','DSCR includes only months with debt service; no service is not zero DSCR. No undocumented bank covenant has been introduced. VAT credit and refund peaks show timing pressure.','note')
  for j in range(180):
   rr=216+j;mr=5+j
   for ar,c,kind in [(f'E{rr}:F{rr}','A','number'),(f'G{rr}:I{rr}','B','money'),(f'J{rr}:L{rr}','C','money'),(f'M{rr}:O{rr}','D','money'),(f'P{rr}:R{rr}','E','money'),(f'S{rr}:V{rr}','I','money'),(f'W{rr}:Y{rr}','K','ratio')]:put(s,ar,q(CASH,c+str(mr)),0,kind)
  for ar,pt,en in [('E214:F215','Mês','Month'),('G214:I215','Caixa','Cash'),('J214:L215','Reservas','Reserves'),('M214:O215','Dívida','Debt'),('P214:R215','Serviço','Service'),('S214:V215','Crédito IVA','VAT credit'),('W214:Y215','DSCR','DSCR')]:label(s,ar,pt,en,'head')
  # Analytical real estate filters, separate from operational hotel inputs.
  s=ui('DASH 04A V06');floors=sorted({val('V06 Unidades','C'+str(r)) for r in range(5,99)})
  for addr,v,pt,en,choices in [('H27','All','Tipologia','Typology',['All']+types),('N27','All','Piso','Floor',['All']+floors),('H29','All','Estado previsto','Forecast status',['All','Available','Reserved','Deeded']),('N29',0,'Área mínima m²','Min area m²',None),('H31',99999,'Área máxima m²','Max area m²',None),('H33',0,'Preço mínimo €/m²','Min price €/m²',None),('N33',99999,'Preço máximo €/m²','Max price €/m²',None),('H35',0,'Mês de consulta','As-of month',None)]:
   r,c=rc(addr);label(s,f'{col(c-3)}{r+1}:{col(c-1)}{r+1}',pt,en);editable(s,f'{addr}:{col(c+2)}{r+1}',v,'input',choices)
  label(s,'E190:Y191','Inventário e quadro por tipologia','Inventory and typology analysis','section')
  for j in range(len(types)):
   rr=194+j*2;mr=105+j
   for ar,c,kind in [(f'E{rr}:F{rr+1}','A','text'),(f'G{rr}:H{rr+1}','B','number'),(f'I{rr}:J{rr+1}','C','number'),(f'K{rr}:L{rr+1}','D','number'),(f'M{rr}:N{rr+1}','E','money'),(f'O{rr}:Q{rr+1}','F','money'),(f'R{rr}:U{rr+1}','G','money'),(f'V{rr}:W{rr+1}','H','number'),(f'X{rr}:Y{rr+1}','I','number')]:put(s,ar,q(SALES,c+str(mr)),0,kind)
  for ar,pt,en in [('E192:F193','Tipo','Type'),('G192:H193','Nº','Units'),('I192:J193','Área m²','Area m²'),('K192:L193','Média m²','Avg m²'),('M192:N193','€/m²','€/m²'),('O192:Q193','€/unidade','€/unit'),('R192:U193','VGV','GDV'),('V192:W193','Disponíveis','Available'),('X192:Y193','Vendidas*','Sold*')]:label(s,ar,pt,en,'head')
  label(s,'E210:Y213','*Estados comerciais são projeções do calendário, não prova documental. Filtros de consulta não alteram o investimento; preços e atrasos afetam o motor.','*Commercial states are calendar projections, not documentary evidence. View filters do not change investment; pricing and delays affect the engine.','note')
  label(s,'E216:P217','Unidades que correspondem aos filtros','Units matching filters','section');put(s,'Q216:Y217','SUM(\''+SALES+'\'!H5:H98)',94,'number')
  for j in range(94):
   rr=220+j;mr=5+j
   for ar,c,kind in [(f'E{rr}:G{rr}','A','text'),(f'H{rr}:J{rr}','B','text'),(f'K{rr}:L{rr}','C','text'),(f'M{rr}:N{rr}','D','number'),(f'O{rr}:R{rr}','E','money'),(f'S{rr}:U{rr}','F','money'),(f'V{rr}:X{rr}','G','text'),(f'Y{rr}:Y{rr}','H','number')]:put(s,ar,q(SALES,c+str(mr)),0,kind)
  # Investor class, individual block and group: both investor dashboards are selectable.
  blockids=[val('V06 Blocos','A'+str(r)) for r in range(5,20)]
  for case,sn,selection in [(0,'DASH 06 V06',[('H25','Class','Classe',['All','Platinum','Gold','Silver']),('N25','Block','Bloco',['All']+blockids),('H27','Group','Grupo',['New','Promoters','All'])]),(1,'DASH 07 V06',[('H31','Class','Classe',['All','Platinum','Gold','Silver']),('N31','Block','Bloco',['All']+blockids),('H33','Group','Grupo',['New','Promoters','All'])])]:
   s=ui(sn)
   for addr,en,pt,choices in selection:
    rr,c=rc(addr);r=rr+1;label(s,f'{col(c-3)}{r}:{col(c-1)}{r}',pt,en);editable(s,f'{addr}:{col(c+2)}{r}','New' if en=='Group' else 'All','input',choices)
   label(s,'E190:Y191','Impacto do Cenário no Investidor','Investor Impact Summary','section');start=0 if case==0 else 12
   for j in range(11):
    rr=194+j*2;mr=26+j
    for ar,c,kind in [(f'E{rr}:L{rr+1}',start,'label'),(f'M{rr}:P{rr+1}',start+1,'money'),(f'Q{rr}:U{rr+1}',start+2,'money'),(f'V{rr}:Y{rr+1}',start+3,'money')]:put(s,ar,q(INV,col(c)+str(mr)),0,'pct' if j in [1,3,6,9,10] and c!=start else 'ratio' if j==5 and c!=start else 'number' if j==7 and c!=start else kind)
   for ar,pt,en in [('E192:L193','Indicador','Metric'),('M192:P193','Base','Reference'),('Q192:U193','Ativo / Teste','Active / Test'),('V192:Y193','Variação','Delta')]:label(s,ar,pt,en,'head')
   label(s,'E218:Y221','Promotores: direitos adquiridos, sem ticket de entrada conhecido. Não se calcula TIR/VAL/MOIC fictício para estes direitos; seleções que incluem promotores mostram N/A. Capital realizado é documental, distinto do capital previsto de cenário.','Promoters hold acquired rights, without known entry cost. No fictitious IRR/NPV/MOIC is calculated for those rights; selections including promoters show N/A. Paid capital is documentary and distinct from scenario receipts.','note')
  s=ui('DASH 06 V06');label(s,'E224:Y225','Estrutura acionista e blocos — direitos vs novos tickets','Ownership and blocks — acquired rights vs new tickets','section')
  for j in range(15):
   rr=228+j*2;mr=5+j
   for ar,c,kind in [(f'E{rr}:F{rr+1}','A','text'),(f'G{rr}:H{rr+1}','B','text'),(f'I{rr}:J{rr+1}','C','text'),(f'K{rr}:K{rr+1}','D','pct'),(f'L{rr}:N{rr+1}','E','money'),(f'O{rr}:Q{rr+1}','F','money'),(f'R{rr}:T{rr+1}','H','money'),(f'U{rr}:W{rr+1}','I','money'),(f'X{rr}:Y{rr+1}','G','text')]:put(s,ar,q(INV,c+str(mr)),0,kind)
  label(s,'E260:Y264','Platinum 20% / ticket novo 10 M€; Gold 10% / 5 M€; Silver 2% / 1 M€. JSCV: 2 Platinum = 40%; JALS: 5 Silver = 10%. Equivalentes económicos não são entradas novas. Direitos societários, informação e facilidades devem ser aprovados separadamente dos direitos económicos.','Platinum 20% / new ticket €10m; Gold 10% / €5m; Silver 2% / €1m. JSCV: 2 Platinum = 40%; JALS: 5 Silver = 10%. Economic equivalents are not new subscriptions. Governance, information and access terms require separate approval.','note')
  # Per-dashboard print pages: show six global metrics and both reference and active.
  for index in range(9):
   s=ui(f'V06 PDF {index+1:02}');s.merge('A47:L57',None,style='note');label(s,'A47:L48','Impacto Global — Referência R4 vs Ativo','Project-Wide Impact — Reference R4 vs Active','section')
   for ar,pt,en in [('A49:D49','Indicador','KPI'),('E49:F49','Base','Reference'),('G49:H49','Ativo','Active'),('I49:J49','Delta','Delta'),('K49:L49','Delta %','Delta %')]:label(s,ar,pt,en,'head')
   for j in range(6):
    rr=50+j;mr=5+j;put(s,f'A{rr}:D{rr}',q(IMP,'A'+str(mr)),metrics[j][1],'label')
    for ar,c,sty in [(f'E{rr}:F{rr}','B','base'),(f'G{rr}:H{rr}','C','test'),(f'I{rr}:J{rr}','D','delta'),(f'K{rr}:L{rr}','E','pct')]:put(s,ar,q(IMP,c+str(mr)),metrics[j][4] if c in ['B','C'] else 0,'pct' if c=='E' or j==5 else sty)
   label(s,'A56:L57','Motor de 180 meses; hipóteses ativas. VAL e folga de reservas não equivalem a aprovação de financiamento. A matriz ADR/ocupação usa os fluxos reais dos investidores.','180-month model; active assumptions. NPV and reserve headroom are not financing approval. The ADR/occupancy matrix uses actual investor cash flows.','note')
  # A4 consolidated report: replace page-1 and page-10 redundant tables with comparison.
  s=ui('DASH 10 V06')
  for first in [36,576]:
   s.merge(f'A{first}:L{first+12}',None,style='note');label(s,f'A{first}:L{first+1}','Impacto Global — Referência R4 vs Ativo','Project-Wide Impact — Reference R4 vs Active','section')
   for j in range(6):
    rr=first+3+j;mr=5+j;put(s,f'A{rr}:F{rr}',q(IMP,'A'+str(mr)),metrics[j][1],'label');put(s,f'G{rr}:H{rr}',q(IMP,'B'+str(mr)),metrics[j][4],'pct' if j==5 else 'base');put(s,f'I{rr}:J{rr}',q(IMP,'C'+str(mr)),metrics[j][4],'pct' if j==5 else 'test');put(s,f'K{rr}:L{rr}',q(IMP,'D'+str(mr)),0,'pct' if j==5 else 'delta')
   put(s,f'A{first+10}:L{first+12}',q(IMP,'B18'),'VAL permanece negativo.','note')
  # Every module has Home/Back; technical reference and routing are hidden, not public filters.
  sheets=wb.find(T('sheets'))
  for n in sheets:
   if n.get('name') in [REF,ACT,IMP]:n.set('state','hidden')
  cp=wb.find(T('calcPr'))
  if cp is None:cp=E.SubElement(wb,T('calcPr'))
  cp.set('calcMode','auto');cp.set('fullCalcOnLoad','1');cp.set('forceFullCalc','1');changes['xl/workbook.xml']=xml(wb)
  # Normalize worksheet order and optional child cardinality (preserve R1 repair).
  for sn,r in roots.items():
   s=Sheet(r,styles);changes[pp[sn]]=s.finish()
  with ZipFile(MASTER,'w',ZIP_DEFLATED,compresslevel=6) as target:
   for item in z.infolist():target.writestr(copy.copy(item),changes.get(item.filename,z.read(item.filename)))
 spec={'release':'R5 CHECKPOINT','source_sha256':hashlib.sha256(BASE.read_bytes()).hexdigest(),'result_sha256':hashlib.sha256(MASTER.read_bytes()).hexdigest(),'inputs':INPUTS,'new_ui_inputs':new_inputs,'parameter_routing':{i['sheet']+'!'+i['cell']:a(i['sheet'],i['cell']) for i in INPUTS},'routed_formulas':changed_formula_count,'collection_formulas_extended':collection_changes,'display_bindings':display,'reference':'Immutable R4 parameter, KPI, monthly and block-flow snapshot; not a second engine','discount':'V06 Registo Ativo!B65 annual; NPV uses (1+annual)^(1/12)-1 on months 1..180 plus month-zero flow','native_excel365':'NOT TESTED','source_R4_native_acceptance':'NOT CONFIRMED','official_logo':'Original image bytes unavailable; insertion area retained','known_limits':['R4 chart geometry retained unless explicit topic restructuring follows','Promoter return measures remain N/A without historical costs and dates','Scenario view filters do not change economic outcomes by design']}
 (OUT/'Implementation.json').write_text(json.dumps(spec,indent=2,ensure_ascii=False));print('Built',MASTER,flush=True)
if __name__=='__main__':patch()
