"""Premium presentation, canonical front-panel inputs and report control centre.
Reads the stable V06 workbook. No independent financial engine is introduced.
"""
from pathlib import Path
from zipfile import ZipFile
import sys,re,json,copy
import xlsxwriter
from xlsxwriter.utility import xl_col_to_name as col, xl_cell_to_rowcol
from assemble import paths,T,E
from reference_io import load_reference
from core import legacy
from report_final import SECTIONS
ROOT=Path(__file__).resolve().parents[1];WORK=Path('/workspace/kk01_efcc');BASE=WORK/'checkpoints/C1_Stable_4111b2c.xlsx'
CENTER='DASH 10 Centro';LANG="'V06 Controlos'!$B$5"
DASHES=[('EXEC V06','Dashboard Executivo','Executive Dashboard'),('DASH 01 V06','Investimento e Orçamento de Construção','Investment & Construction Budget'),('DASH 02 V06','Fiscalidade, Incentivos e Seguros','Taxation, Incentives & Insurance'),('DASH 03 V06','Desempenho Financeiro Consolidado','Consolidated Financial Performance'),('DASH 04A V06','Promoção e Comercialização Imobiliária','Real Estate Development & Sales'),('DASH 04B V06','Exploração Residencial e Sell & Leaseback','Residential Operations & Sell & Leaseback'),('DASH 05 V06','Exploração Hoteleira e Centro MICE','Hotel & MICE Operations'),('DASH 06 V06','Estrutura Acionista e Blocos de Investimento','Shareholding & Investment Blocks'),('DASH 07 V06','Rentabilidade dos Acionistas e Distribuições','Equity Returns & Distributions'),(CENTER,'Relatório Executivo Consolidado de Viabilidade e Investimento','Consolidated Investment & Feasibility Report')]
NAV=[];ALIASES=[];ADDED_INPUTS=[];INPUTS=[]
class Source:
 def __init__(self,path):
  self.z=ZipFile(path);self.w,self.pp=paths(self.z);self.ss=[''.join(n.itertext()) for n in E.fromstring(self.z.read('xl/sharedStrings.xml'))];self.cache={};self.styles=E.fromstring(self.z.read('xl/styles.xml'));self.formats={int(n.get('numFmtId')):n.get('formatCode') for n in self.styles.find(T('numFmts'))}
 def cells(self,sn):
  if sn not in self.cache:self.cache[sn]={c.get('r'):c for c in E.fromstring(self.z.read(self.pp[sn])).iter(T('c'))}
  return self.cache[sn]
 def get(self,sn,addr):
  c=self.cells(sn).get(addr)
  if c is None:return None,None
  f=c.findtext(T('f'));v=c.findtext(T('v'))
  if c.get('t')=='s':v=self.ss[int(v)]
  elif c.get('t')=='inlineStr':v=''.join(c.find(T('is')).itertext())
  elif c.get('t') not in ['str','e'] and v is not None:
   try:v=float(v);v=int(v) if v.is_integer() else v
   except ValueError:pass
  return f,v
 def format_kind(self,sn,addr):
  c=self.cells(sn).get(addr);xf=self.styles.find(T('cellXfs'))[int(c.get('s','0'))];n=int(xf.get('numFmtId','0'));code=self.formats.get(n,'')
  label=str(self.get(sn,'A'+str(xl_cell_to_rowcol(addr)[0]+1))[1] or '').lower() if sn in ['DASH 10 V06','V06 Brief'] else ''
  if 'cobertura' in label or 'coverage' in label or 'moic' in label:return 'ratio_small'
  if any(k in label for k in ['meses','months','carência','tenor']):return 'number'
  return 'percent' if n in [9,10] or '%' in code else 'ratio' if '×' in code else 'num'
 def card_kind(self,sn,addr):
  c=self.cells(sn)[addr];xf=self.styles.find(T('cellXfs'))[int(c.get('s','0'))];n=int(xf.get('numFmtId','0'));code=self.formats.get(n,'');kind=self.format_kind(sn,addr)
  return 'pct' if kind=='percent' else 'ratio' if '×' in code else 'money' if '€' in code else 'count'
 def label(self,sn,addr):
  f,v=self.get(sn,addr)
  if f:return f,v
  return '"'+str(v or '').replace('"','""')+'"',v or ''

def build():
 src=Source(BASE);ref=load_reference(ROOT);rows=ref['main_rows'];cc=ref['main_columns'];hp=legacy.config();out=WORK/'Premium_UI.xlsx';wb=xlsxwriter.Workbook(out);wb.set_calc_mode('auto')
 P={'navy':'142337','graphite':'252932','red':'C8101E','white':'FFFFFF','bg':'F5F7FA','blue':'1765A8','grey':'687586'}
 def fmt(**kw):return wb.add_format({'font_name':'Aptos','font_size':10,'font_color':P['navy'],'valign':'vcenter',**kw})
 F={'bg':fmt(bg_color=P['bg']),'brand':fmt(bg_color=P['red'],font_color='FFFFFF',bold=True,font_size=12,text_wrap=True),'header':fmt(bg_color=P['navy'],font_color='FFFFFF',bold=True,font_size=21,text_wrap=True),'sub':fmt(font_color=P['grey'],font_size=10,text_wrap=True),'side':fmt(bg_color=P['navy'],font_color='FFFFFF',font_size=10,text_wrap=True,indent=1),'active':fmt(bg_color=P['red'],font_color='FFFFFF',bold=True,font_size=10,text_wrap=True,indent=1),'nav':fmt(bg_color='E9EFF5',font_color=P['blue'],bold=True,font_size=10,text_wrap=True),'section':fmt(bg_color='E9EFF5',bold=True,font_size=12,text_wrap=True),'label':fmt(bg_color='FFFFFF',font_color=P['grey'],font_size=10,text_wrap=True),'money':fmt(bg_color='FFFFFF',bold=True,font_size=22,num_format='#,##0 "€";[Red](#,##0) "€";–'),'pct':fmt(bg_color='FFFFFF',bold=True,font_size=22,num_format='0.0%'),'count':fmt(bg_color='FFFFFF',bold=True,font_size=22,num_format='#,##0'),'ratio':fmt(bg_color='FFFFFF',bold=True,font_size=22,num_format='0.00"×"'),'input':fmt(bg_color='E6F0FA',font_color=P['blue'],bold=True,num_format='#,##0.00',bottom=1,bottom_color='B8CCE0'),'inputpct':fmt(bg_color='E6F0FA',font_color=P['blue'],bold=True,num_format='0.00%',bottom=1,bottom_color='B8CCE0'),'inputtext':fmt(bg_color='E6F0FA',font_color=P['blue'],bold=True,text_wrap=True,bottom=1,bottom_color='B8CCE0'),'text':fmt(text_wrap=True),'num':fmt(num_format='#,##0 "€";[Red](#,##0) "€";–'),'ratio_small':fmt(num_format='0.00"×"'),'percent':fmt(num_format='0.0%'),'number':fmt(num_format='#,##0.0'),'thead':fmt(bg_color=P['navy'],font_color='FFFFFF',bold=True,text_wrap=True),'note':fmt(bg_color='E9EFF5',font_color=P['grey'],font_size=9,text_wrap=True),'alert':fmt(font_color=P['red'],bold=True,font_size=10,text_wrap=True)}
 F['activepct']=fmt(bg_color='FFFFFF',font_color=P['navy'],bold=True,num_format='0.00%')
 F['activenum']=fmt(bg_color='FFFFFF',font_color=P['navy'],bold=True,num_format='#,##0.00')
 inactive=fmt(bg_color='EEF0F2',font_color='9299A2')
 F['heat']=fmt(font_size=9,num_format='0.0,," M€";[Red](0.0,," M€")')
 def tr(pt,en):return f'IF({LANG}="PT","{pt.replace(chr(34),chr(34)*2)}","{en.replace(chr(34),chr(34)*2)}")'
 def merge(s,area,expression,value,style='text'):
  s.merge_range(area,'',F[style]);s.write_formula(area.split(':')[0],'='+expression,F[style],value)
 def text(s,area,pt,en,style='text'):merge(s,area,tr(pt,en),pt,style)
 def link(s,area,target,pt,en):
  sh,addr=target;merge(s,area,f'HYPERLINK("#\'{sh}\'!{addr}",{tr(pt,en)})',pt,'nav');NAV.append({'origin_sheet':s.name,'origin_cell':area.split(':')[0],'destination_sheet':sh,'destination_cell':addr,'type':'Excel internal HYPERLINK','PT':pt,'EN':en})
 def clone(s,addr,sn,cell,style='num'):
  f,v=src.get(sn,cell)
  if f:s.write_formula(addr,'='+f,F[style],v)
  else:s.write(addr,v,F[style])
 def control(s,addr,sn,cell,pt,en,kind='input',options=None,bounds=None):
  r,c=xl_cell_to_rowcol(addr);s.merge_range(r,c,r,c+2,'',F[kind]);f,v=src.get(sn,cell);assert f is None,(sn,cell,f)
  s.write(addr,v,F[kind]);text(s,f'{col(c-3)}{r+1}:{col(c-1)}{r+1}',pt,en,'sub')
  if options:s.data_validation(addr,{'validate':'list','source':options,'error_type':'stop','error_title':'Option','error_message':'Select a listed option.'})
  elif bounds:s.data_validation(addr,{'validate':'decimal','criteria':'between','minimum':bounds[0],'maximum':bounds[1],'error_type':'stop','error_message':f'{bounds[0]} … {bounds[1]}'})
  else:
   root=E.fromstring(src.z.read(src.pp[sn]));dv=root.find(T('dataValidations'))
   if dv is not None:
    for d in dv:
     if d.get('sqref')==cell and d.get('type')=='list':
      expr=d.findtext(T('formula1'))
      if expr and expr.startswith('"'):s.data_validation(addr,{'validate':'list','source':expr.strip('"').split(',')})
     elif d.get('sqref')==cell and d.get('type')=='custom':s.data_validation(addr,{'validate':'custom','value':'='+d.findtext(T('formula1')),'error_type':'stop'})
  if isinstance(v,(int,float)) and not options and bounds is None:
   original=E.fromstring(src.z.read(src.pp[sn]));dv=original.find(T('dataValidations'));has_rule=dv is not None and any(d.get('sqref')==cell and d.get('type') in ['list','custom'] for d in dv)
   if not has_rule:
    whole=sn=='V06 Controlos' and cell in ['B8','B10','B35','B37','B38','B42','B43']
    s.data_validation(addr,{'validate':'integer' if whole else 'decimal','criteria':'between','minimum':0,'maximum':1 if kind=='inputpct' else 360 if whole else 1e12,'error_type':'stop','error_title':'Numeric input','error_message':'Use a valid numeric value within the permitted range.'})
  INPUTS.append({'source_sheet':sn,'source_cell':cell,'sheet':s.name,'cell':addr,'value':v,'label_pt':pt,'label_en':en})
  if (sn,cell)!=(s.name,addr):ALIASES.append({'source_sheet':sn,'source_cell':cell,'destination_sheet':s.name,'destination_cell':addr,'value':v})
 def effective(s,addr,pt,en,source,percent=True):
  r,c=xl_cell_to_rowcol(addr);text(s,f'{col(c-3)}{r+1}:{col(c-1)}{r+1}',pt,en,'sub')
  sn,cell=source;_,v=src.get(sn,cell);merge(s,f'{addr}:{col(c+2)}{r+1}',f"'{sn}'!{cell}",v,'activepct' if percent else 'activenum')
 def custom_style(s,addresses,selector,choice):
  for addr in addresses:
   r,c=xl_cell_to_rowcol(addr);s.conditional_format(r,c,r,c+2,{'type':'formula','criteria':f'=${selector[0]}${selector[1:]}<>"{choice}"','format':inactive})
 def newcontrol(s,addr,key,value,pt,en,bounds,kind='inputpct'):
  r,c=xl_cell_to_rowcol(addr);s.merge_range(r,c,r,c+2,value,F[kind]);text(s,f'{col(c-3)}{r+1}:{col(c-1)}{r+1}',pt,en,'sub');s.data_validation(addr,{'validate':'decimal','criteria':'between','minimum':bounds[0],'maximum':bounds[1],'error_type':'stop'});ADDED_INPUTS.append({'key':key,'sheet':s.name,'cell':addr,'value':value});INPUTS.append({'sheet':s.name,'cell':addr,'value':value,'label_pt':pt,'label_en':en,'new_exposure_of_existing_constant':True})
 def frame(sn,i):
  s=wb.add_worksheet(sn);s.hide_gridlines(2);s.hide_row_col_headers();s.set_zoom(85);s.freeze_panes(6,4);s.set_default_row(16);s.set_column('A:A',1);s.set_column('B:C',18);s.set_column('D:D',2);s.set_column('E:P',12);s.set_column('Q:Q',1)
  for r in range(96):
   s.set_row(r,16)
   for c in range(17):s.write_blank(r,c,None,F['side'] if c<4 else F['bg'])
  s.merge_range('B2:C4','JS-CV\nCONSTRUÇÕES E\nINVESTIMENTOS, LDA.',F['brand']);text(s,'E2:M3','KK-01 | HOTEL & RESIDENCES','KK-01 | HOTEL & RESIDENCES','header');link(s,'N2:P3',('V06 Identidade','B8'),'Logótipo original','Original logo');text(s,'E4:P4',f'{i+1:02} — {DASHES[i][1]}',f'{i+1:02} — {DASHES[i][2]}','sub');merge(s,'E1:P1',"'V06 Hipoteses'!D29&\" · \"&'V06 Hipoteses'!D28",'EFCC-20261022-01 · 2026-10-22','note')
  link(s,'B5:C5',(CENTER,'H7'),'PT | EN ▾','PT | EN ▾');link(s,'B6:C6',(sn,'E6'),'Menu de navegação','Navigation menu')
  link(s,'E5:G5',('EXEC V06','A1'),'⌂ Início','⌂ Home');link(s,'H5:J5',(sn,'H23' if i<9 else 'H7'),'Parâmetros ▾','Assumptions ▾');link(s,'K5:M5',(CENTER,'A1'),'Relatório 10','Report 10');link(s,'N5:P5',(f'V06 PDF {i+1:02}','A1') if i<9 else ('V06 Imprimir','A23'),'Imprimir / PDF','Print / PDF')
  for n,(target,pt,en) in enumerate(DASHES):
   rr=8+n*5;merge(s,f'B{rr}:C{rr+3}',f'HYPERLINK("#\'{target}\'!A1",{tr(f"{n+1:02}  {pt}",f"{n+1:02}  {en}")})',f'{n+1:02}  {pt}','active' if i==n else 'side');NAV.append({'origin_sheet':sn,'origin_cell':f'B{rr}','destination_sheet':target,'destination_cell':'A1','type':'sidebar','PT':pt,'EN':en})
  for n,(target,pt,en) in enumerate(DASHES):
   addr=f'{col(4+n)}6';s.write_formula(addr,f'=HYPERLINK("#\'{target}\'!A1","{n+1:02}")',F['active'] if n==i else F['nav'],f'{n+1:02}');s.write_comment(addr,pt+' / '+en);NAV.append({'origin_sheet':sn,'origin_cell':addr,'destination_sheet':target,'destination_cell':'A1','type':'frozen navigation menu','PT':pt,'EN':en})
  link(s,'B60:C62',('V06 Identidade','B8'),'Inserir logótipo original','Insert original logo')
  s.set_landscape();s.set_paper(8);s.fit_to_pages(1,1);s.set_margins(.25,.25,.3,.3);s.print_area('E7:P88');s.repeat_rows(1,3);s.set_footer('&LKK-01 | Executive Financial Control Center&R&P / &N');s.set_tab_color(P['red'] if i<9 else P['navy']);return s
 # Native chart data: projection of existing monthly engine outputs only.
 data=wb.add_worksheet('V06 Visual');data.hide();keys=['cash','reserve','capex','receipts','vat_refund','tax_paid','debt','deed_sales','res_revenue','owner_distribution','owner_target','equity','investor_outlay','dividend','room_revenue','fb_revenue','mice_revenue','other_revenue','ebitda','revenue','ebitda_margin'];data.write_row('A4',['Year']+keys);stocks={'cash','reserve','debt','vat_credit','ffe_restricted_fund','res_fund','working_capital','res_wc'}
 chartcols={k:1+j for j,k in enumerate(keys)}
 for y in range(1,16):
  data.write(y+3,0,y)
  for k,c in chartcols.items():
   a=col(cc[k]);f=f"'V06 Motor'!{a}{y*12+4}" if k in stocks else f"SUM('V06 Motor'!{a}{(y-1)*12+5}:{a}{y*12+4})";v=rows[y*12-1][k] if k in stocks else sum(r[k] for r in rows[(y-1)*12:y*12]);
   if k=='ebitda_margin':
    e=col(cc['ebitda']);rev=col(cc['revenue']);f=f'IF(SUM(\'V06 Motor\'!{rev}{(y-1)*12+5}:{rev}{y*12+4})>0,SUM(\'V06 Motor\'!{e}{(y-1)*12+5}:{e}{y*12+4})/SUM(\'V06 Motor\'!{rev}{(y-1)*12+5}:{rev}{y*12+4}),0)';den=sum(r['revenue'] for r in rows[(y-1)*12:y*12]);v=sum(r['ebitda'] for r in rows[(y-1)*12:y*12])/den if den else 0
   data.write_formula(y+3,c,'='+f,F['num'],v)
 def chart(s,anchor,series,type='line',width=495,height=230,stack=False):
  ch=wb.add_chart({'type':type,**({'subtype':'stacked'} if stack else {})})
  for j,(key,pt,en) in enumerate(series):
   c=chartcols[key];name_row=24+len(NAV)+j;data.write_formula(name_row,c,'='+tr(pt,en),F['text'],pt)
   ch.add_series({'name':['V06 Visual',name_row,c],'categories':['V06 Visual',4,0,18,0],'values':['V06 Visual',4,c,18,c],'line':{'color':['142337','C8101E','1765A8','687586'][j%4],'width':2},'fill':{'color':['142337','1765A8','C8101E','687586'][j%4]},'border':{'none':True}})
  if s.name=='DASH 05 V06' and len(series)==1 and series[0][0]=='ebitda':
   second=wb.add_chart({'type':'line'});second.add_series({'name':'EBITDA %','categories':['V06 Visual',4,0,18,0],'values':['V06 Visual',4,chartcols['ebitda_margin'],18,chartcols['ebitda_margin']],'y2_axis':True,'line':{'color':'C8101E','width':2}});ch.combine(second);ch.set_y2_axis({'num_format':'0%','num_font':{'size':9}})
  ch.set_size({'width':width,'height':height});ch.set_chartarea({'border':{'none':True},'fill':{'color':'FFFFFF'}});ch.set_plotarea({'border':{'none':True}});ch.set_legend({'position':'bottom','font':{'name':'Aptos','size':10}});ch.set_x_axis({'num_font':{'name':'Aptos','size':9}});ch.set_y_axis({'num_format':'0,," M€"','num_font':{'name':'Aptos','size':9},'major_gridlines':{'visible':True,'line':{'color':'E6EBF1'}}});s.insert_chart(anchor,ch)
 def total(k):return f"SUM('V06 Motor'!{col(cc[k])}5:{col(cc[k])}184)",sum(r[k] for r in rows)
 def table(s,start,records):
  text(s,f'E{start}:L{start}','Indicador / base','Indicator / basis','thead');text(s,f'M{start}:P{start}','Valor','Value','thead')
  for i,(pt,en,f,v,kind) in enumerate(records):
   r=start+1+i*2;text(s,f'E{r}:L{r+1}',pt,en,'text');merge(s,f'M{r}:P{r+1}',f,v,kind)
 def metricrec(key,pt,en,kind='num'):
  f,v=total(key);return pt,en,f,v,kind
 print_records={};print_series={}
 # Copy the six existing KPI formulas and translated labels, without new financial rules.
 for i,(sn,pt,en) in enumerate(DASHES[:9]):
  s=frame(sn,i)
  for n in range(6):
   r=7+(n//3)*6;c=4+(n%3)*4;label=src.label(sn,f'{col(c)}{r}');merge(s,f'{col(c)}{r}:{col(c+2)}{r+1}',*label,'label');f,v=src.get(sn,f'{col(c)}{r+2}');kind=src.card_kind(sn,f'{col(c)}{r+2}');merge(s,f'{col(c)}{r+2}:{col(c+3)}{r+4}',f,v,kind);link(s,f'{col(c+3)}{r}:{col(c+3)}{r+1}',('V06 Dicionario','A4'),'ⓘ','ⓘ')
  text(s,'E20:P21',f'{i+1:02} — {pt}',f'{i+1:02} — {en}','section')
  if sn=='DASH 05 V06':
   inputs=[('H23','H23','Cenário operacional','Operating scenario','inputtext',['Conservative','Base Case','Favorable','Upside Stress','Custom']),('N23','N23','Ano operacional','Operating year','input',None),('H25','H25','ADR Custom','Custom ADR','input',None),('H27','H27','Ocupação Custom','Custom occupancy','inputpct',None),('N25','N25','Utilização MICE','MICE utilisation','inputpct',None),('N27','N27','Reserva FF&E','FF&E reserve','inputpct',None),('N29','N29','Crescimento ADR','ADR growth','inputpct',None),('H32','H32','Condições operadora','Operator terms','inputtext',['Operator Proposed Terms','KK Target Terms','KK Optimized Terms','Custom Terms']),('H34','H34','Marketing Custom','Custom marketing','inputpct',None),('N34','N34','Distribuição Custom','Custom distribution','inputpct',None),('H36','H36','Prioridade ativa 0/1','Priority enabled 0/1','input',None),('N36','N36','Prioridade anual €','Annual priority €','input',None)]
   for addr,old,a,b,kind,opts in inputs:control(s,addr,sn,old,a,b,kind,opts)
   for addr,old,a,b in [('H29','B68','Base fee Custom','Custom base fee'),('H30','B70','Incentivo Custom','Custom incentive'),('N32','B69','Sistemas Custom','Custom systems')]:control(s,addr,'V06 Controlos',old,a,b,'inputpct',bounds=(0,1))
   effective(s,'H24','ADR ativo · abertura','Active opening ADR',(sn,'E9'),False)
   effective(s,'H26','Ocupação ativa · abertura','Active opening occupancy',(sn,'I9'))
   for addr,r,a,b in [('H28',68,'Base fee ativo','Active base fee'),('N28',69,'Sistemas ativos · total','Active systems · total'),('N30',70,'Incentivo ativo','Active incentive')]:effective(s,addr,a,b,('V06 Registo Ativo','B'+str(r)))
   custom_style(s,['H25','H27'],'H23','Custom');custom_style(s,['H29','H30','N32','H34','N34'],'H32','Custom Terms')
   text(s,'E39:P39','Branco: valor ativo. Custom cinzento: guardado, não aplicado. Base Case conserva o cenário Hotel de referência.','White: active value. Grey Custom: stored, not applied. Base Case retains the reference Hotel scenario.','note')
   link(s,'E38:J38',('V06 Matriz ADR','B3'),'Sensibilidade ADR × Ocupação →','ADR × Occupancy sensitivity →');link(s,'K38:P38',('V06 Operacao','A4'),'Operação mensal e reservas →','Monthly operations and reserves →')
   text(s,'E40:P40','Indicadores operacionais · ano selecionado','Operating performance · selected year','section')
   selected=[r for r in rows if r['op_age']==4 and r['available_room_nights']>0]
   def annual(k):return f'SUMIFS(\'V06 Motor\'!{col(cc[k])}5:{col(cc[k])}184,\'V06 Motor\'!E5:E184,$N$23-1,\'V06 Motor\'!DL5:DL184,">0")',sum(r[k] for r in selected)
   # DL is available_room_nights (column 115 zero-based).
   assert col(cc['available_room_nights'])=='DL'
   specs=[('Quartos disponíveis','Available rooms','available_room_nights','count',365),('Dormidas vendidas','Occupied room nights','occupied_room_nights','count',1),('Receita quartos','Room revenue','room_revenue','money',1),('Receita F&B','F&B revenue','fb_revenue','money',1),('Receita MICE','MICE revenue','mice_revenue','money',1),('Receita total','Total revenue','revenue','money',1),('RevPAR','RevPAR','revpar','money',1),('EBITDA','EBITDA','ebitda','money',1)]
   for n,(a,b,k,kind,d) in enumerate(specs):
    rr=42+(n//4)*6;c=4+(n%4)*3
    if k=='revpar':
     rf,rv=annual('room_revenue');af,av=annual('available_room_nights');f=f'IF(({af})>0,({rf})/({af}),0)';v=rv/av if av else 0
    else:f,v=annual(k)
    text(s,f'{col(c)}{rr}:{col(c+2)}{rr+1}',a,b,'label');merge(s,f'{col(c)}{rr+2}:{col(c+2)}{rr+4}',f'({f})/{d}',v/d,kind)
   chart(s,'E55',[('room_revenue','Quartos','Rooms'),('fb_revenue','F&B','F&B'),('mice_revenue','MICE','MICE'),('other_revenue','Outras receitas','Other revenue')],type='column',stack=True,width=335,height=235);chart(s,'M55',[('ebitda','EBITDA','EBITDA')],type='column',width=335,height=235)
   for n,(key,a,b) in enumerate([('room_revenue','Quartos','Rooms'),('fb_revenue','F&B','F&B'),('mice_revenue','MICE','MICE'),('other_revenue','Outras receitas','Other revenue')]):
    data.write_formula(700+n,0,'='+tr(a,b),F['text'],a);f,v=annual(key);data.write_formula(700+n,1,'='+f.replace('$N$23',"'DASH 05 V06'!$N$23"),F['num'],v)
   mix=wb.add_chart({'type':'doughnut'});mix.add_series({'categories':['V06 Visual',700,0,703,0],'values':['V06 Visual',700,1,703,1],'points':[{'fill':{'color':v}} for v in ['142337','1765A8','C8101E','687586']],'data_labels':{'percentage':True,'font':{'size':9}}});mix.set_hole_size(65);mix.set_legend({'position':'bottom','font':{'size':9}});mix.set_chartarea({'border':{'none':True}});mix.set_size({'width':335,'height':235});s.insert_chart('I55',mix)
   records=[]
   for k,a,b in [('gop','GOP — ano selecionado','GOP — selected year'),('operator_fees','Fees — ano selecionado','Fees — selected year'),('ffe','Dotação FF&E — ano','FF&E contribution — year'),('ffe_replacement_spend','Reposição FF&E — ano','FF&E replacement — year')]:f,v=annual(k);records.append((a,b,f,v,'num'))
   f,v=annual('ebitda');den,dv=annual('revenue');records.append(('Margem EBITDA','EBITDA margin',f'IF(({den})>0,({f})/({den}),0)',v/dv if dv else 0,'percent'))
   fundcol=col(cc['ffe_restricted_fund']);fundexpr=f'INDEX(\'V06 Motor\'!{fundcol}5:{fundcol}184,MIN(180,\'V06 Motor\'!D5+$N$23*12-1))';records.append(('Saldo fundo FF&E · fim do ano','FF&E fund · year end',fundexpr,rows[89]['ffe_restricted_fund'],'num'))
   otherf,otherv=annual('other_revenue');records.append(('Outras receitas · ano','Other revenue · year',otherf,otherv,'num'))
   print_records[sn]=records;print_series[sn]=[('revenue','Receita','Revenue'),('ebitda','EBITDA','EBITDA')];table(s,71,records);text(s,'E86:P88','Custom ativa ADR e ocupação editáveis. MICE: índice relativo ao caso original; não comprova volume de eventos. FF&E é fundo restrito, separado do custo operacional.','Custom enables editable ADR and occupancy. MICE is an index against the original case, not evidenced event volumes. FF&E is restricted and separate from operating expense.','note')
  else:
   config={0:[('H23','B41','Juro anual','Annual interest','inputpct'),('N23','B65','Desconto Custom','Custom discount','inputpct')],1:[('H23','B38','Atraso obra · meses','Build delay · months','input'),('N23','B39','Sobrecusto obra','Construction overrun','inputpct'),('H25','B75','ETC adicional €','Additional ETC €','input')],2:[('H23','B9','IRPC','Corporate tax','inputpct'),('N23','B6','IVA','VAT','inputpct'),('H25','B8','Reembolso IVA · meses','VAT refund · months','input'),('N25','B10','Isenção simulada · anos','Simulated exemption · years','input')],3:[('H23','B41','Taxa de juro','Interest rate','inputpct'),('N23','B42','Carência · meses','Grace · months','input'),('H25','B43','Prazo · meses','Tenor · months','input'),('N25','B44','Cash sweep','Cash sweep','inputpct')],4:[('H23','B37','Atraso vendas · meses','Sales delay · months','input')],5:[('H27','B72','Yield alvo inicial','Initial target yield','inputpct'),('N23','B73','Incremento yield anual','Annual yield increment','inputpct')],7:[('H23','B35','Atraso capital · meses','Capital delay · months','input'),('N23','B36','Capital recebido de cenário','Scenario capital received','inputpct')],8:[('N23','B57','Política dividendos','Dividend policy','inputtext'),('H25','B62','Múltiplo saída','Exit multiple','input'),('N25','B65','Desconto Custom','Custom discount','inputpct'),('H27','B61','Distribuição parcial','Partial payout','inputpct'),('N27','B58','Cobertura 12M','12M coverage','input')]}.get(i,[])
   for addr,old,a,b,kind in config:
    # A single source for discount and interest lives on the responsible dashboard.
    if i==0:link(s,f'{col(xl_cell_to_rowcol(addr)[1]-3)}23:{col(xl_cell_to_rowcol(addr)[1]+2)}24',('DASH 03 V06','H23') if old=='B41' else ('DASH 07 V06','N25'),a+' →',b+' →')
    else:control(s,addr,'V06 Controlos',old,a,b,kind)
   if i==3:text(s,'E28:P31','Carência e prazo só alteram prestações enquanto existe dívida. O cash sweep pode liquidar a dívida antes da amortização contratual.','Grace and tenor affect instalments only while debt is outstanding. Cash sweep can repay debt before contractual amortisation.','note')
   if i==5:
    control(s,'H23',sn,'H23','Condições residencial','Residential terms','inputtext',['Reference Study','KK Target Terms','KK Optimized Terms','Custom Terms']);control(s,'H25',sn,'H25','Gestão Custom','Custom management','inputpct');control(s,'N25',sn,'N25','Distribuição Custom','Custom distribution','inputpct')
    for addr,key,val,a,b,bounds,kind in [('N27','res_adr',hp['res_adr'],'ADR residencial inicial','Initial residential ADR',(0,1000),'input'),('H29','res_occ',hp['res_occ'],'Ocupação inicial','Initial occupancy',(0,hp['res_occ_max']),'inputpct'),('N29','res_cost',hp['res_cost'],'OPEX / receita','OPEX / revenue',(0,1),'inputpct'),('N31','res_ffe',hp['res_ffe'],'Dotação FF&E','FF&E contribution',(0,1),'inputpct')]:newcontrol(s,addr,key,val,a,b,bounds,kind)
    effective(s,'H24','Gestão ativa','Active management',('V06 Registo Ativo','B94'));effective(s,'N24','Distribuição ativa','Active distribution',('V06 Registo Ativo','B95'));custom_style(s,['H25','N25'],'H23','Custom Terms')
    text(s,'E30:J31','Custom cinzento: guardado, não aplicado.','Grey Custom: stored, not applied.','note')
    link(s,'E33:J34',('V06 Operacao','A4'),'Rendas, custos e reservas →','Owner income, costs and reserves →');text(s,'K33:P35','Yield alvo não é renda garantida.','Target yield is not guaranteed income.','note')
   if i==8:
    control(s,'H23',sn,'H23','Política de reservas','Reserve policy','inputtext',['Conservative','Balanced','Flexible']);control(s,'H29','V06 Matriz ADR','F3','Taxa desconto ativa','Active discount rate','inputtext',['8%','10%','12%','Custom']);link(s,'K29:P30',('V06 Otimizacao','A4'),'Simulação dividendos e reservas →','Dividend and reserve simulation →')
   if i==8:
    effective(s,'N26','Desconto efetivo','Effective discount',('V06 Registo Ativo','B65'));custom_style(s,['N25'],'H29','Custom')
    text(s,'E32:P34','Custom guardado só se aplica no modo Custom. Distribuição parcial só se aplica na política Partial.','Stored Custom applies only in Custom mode. Partial payout applies only with the Partial policy.','note');custom_style(s,['H27'],'N23','Partial')
   owner={0:'V06 Hipoteses',1:'V06 Capitulos',2:'V06 Controlos',3:'V06 Anual',4:'V06 Disponibilidade',5:'V06 Operacao',7:'V06 Blocos',8:'V06 Fluxos Invest'}[i]
   link(s,'E37:J38',(owner,'A4'),'Consultar detalhe →','Open detail →');link(s,'K37:P38',('V06 Dicionario','A4'),'Definições e bases KPI →','KPI definitions and bases →')
   series=[('cash','Caixa','Cash'),('reserve','Reservas','Reserves')]
   if i==1:series=[('capex','CAPEX','CAPEX'),('receipts','Recebimentos','Collections')]
   elif i==2:series=[('vat_refund','IVA recuperado','VAT refunds'),('tax_paid','IRPC','Corporate tax')]
   elif i==3:series=[('cash','Caixa','Cash'),('debt','Dívida','Debt'),('reserve','Reservas','Reserves')]
   elif i==4:series=[('receipts','Recebimentos','Collections'),('deed_sales','Escrituras','Deed sales')]
   elif i==5:series=[('res_revenue','Receita residencial','Residential revenue'),('owner_distribution','Rendas suportáveis','Supported owner income'),('owner_target','Rendas alvo','Target owner income')]
   elif i==7:series=[('equity','Capital KK','KK equity'),('investor_outlay','Capital investidores','Investor equity')]
   elif i==8:series=[('dividend','Dividendos','Dividends'),('cash','Caixa','Cash')]
   text(s,'E40:P41','Evolução anual · motor financeiro central','Annual trend · central financial engine','section');chart(s,'E43',series,width=1020,height=310)
   recordkeys={0:[('capex','CAPEX futuro','Future CAPEX'),('equity','Capital KK previsto','Planned KK equity'),('debt_service','Serviço da dívida','Debt service')],1:[('construction','Construção e FF&E inicial','Construction and initial FF&E'),('supervision','Fiscalização','Supervision'),('contingency_paid','Contingência utilizada','Contingency used'),('extra_etc','ETC adicional','Additional ETC')],2:[('tax_paid','IRPC pago','Corporate tax paid'),('vat_refund','Reembolsos IVA','VAT refunds'),('construction_insurance','Seguro obra','Construction insurance'),('operating_insurance','Seguros exploração','Operating insurance')],3:[('draw','Desembolsos','Drawdowns'),('interest','Juros','Interest'),('principal','Amortização contratual','Contractual principal'),('sweep','Cash sweep','Cash sweep')],4:[('receipts','Recebimentos previstos','Forecast collections'),('deed_sales','Escrituras','Deed sales'),('sales_cost','Custos comerciais','Selling costs'),('sales_profit','Margem promoção · proxy','Development margin · proxy')],5:[('res_revenue','Receitas residenciais','Residential revenue'),('res_opex','Custos operacionais','Operating costs'),('res_fees','Fees residenciais','Residential fees'),('owner_distribution','Rendas suportáveis','Supported owner income'),('owner_shortfall','Yield alvo não coberto','Uncovered target income'),('res_ffe','Dotação FF&E','FF&E contributions')],7:[('equity','Capital KK previsto','Planned KK equity'),('investor_outlay','Capital novos investidores','New investor equity'),('dividend','Dividendos KK','KK dividends')],8:[('dividend','Dividendos KK','KK dividends'),('interest','Juros da dívida','Debt interest'),('tax_paid','Impostos pagos','Taxes paid')]}.get(i,[])
   print_records[sn]=[metricrec(k,a,b) for k,a,b in recordkeys];print_series[sn]=series;table(s,65,print_records[sn]);text(s,'E82:P86','Cenário de estudo. Recebimentos, financiamento, fiscalidade e condições contratuais carecem de documentação. O VAL negativo é apresentado sem ocultação.','Study scenario. Collections, funding, tax and contractual terms require documentary validation. Negative NPV remains visible.','note')
  link(s,'E89:J90',(CENTER,'A1'),'Relatório consolidado →','Consolidated report →');link(s,'K89:P90',('EXEC V06','A1'),'← Início','← Home')
 # Dedicated native print compositions keep screen controls and sidebars off paper.
 for index,(owner,pt,en) in enumerate(DASHES[:9],1):
  sn=f'V06 PDF {index:02}';ps=wb.add_worksheet(sn);ps.hide_gridlines(2);ps.hide_row_col_headers();ps.set_default_row(12);ps.set_column('A:L',14);ps.set_column('M:P',18);ps.set_landscape();ps.set_paper(8);ps.fit_to_pages(1,1);ps.set_margins(.3,.3,.3,.3);ps.print_area('A1:L60');ps.freeze_panes(7,0);ps.set_zoom(85);ps.set_footer('&LKK-01 | Executive Financial Control Center&R&P / &N')
  text(ps,'A1:I3','KK-01 | HOTEL & RESIDENCES','KK-01 | HOTEL & RESIDENCES','header');link(ps,'J1:L3',('V06 Identidade','B8'),'Logótipo original','Original logo');text(ps,'A5:L7',f'{index:02} — {pt}',f'{index:02} — {en}','section')
  for area,target,a,b in [('M1:M3',('EXEC V06','A1'),'Início','Home'),('N1:N3',(owner,'A1'),'Voltar','Back'),('O1:O3',(owner,'H23'),'Parâmetros','Assumptions'),('P1:P3',('V06 Dicionario','A4'),'Ajuda','Help')]:link(ps,area,target,a,b)
  if owner=='DASH 05 V06':
   merge(ps,'A8:L8',tr('Ano operacional selecionado: ','Selected operating year: ')+f'&\'{owner}\'!N23','Ano operacional selecionado: 5','sub')
   selected=[r for r in rows if r['op_age']==4 and r['available_room_nights']>0];av=sum(r['available_room_nights'] for r in selected);rv=sum(r['room_revenue'] for r in selected)
   items=[('Quartos disponíveis','Available rooms','E44',200,'count'),('Dormidas vendidas','Occupied room nights','H44',sum(r['occupied_room_nights'] for r in selected),'count'),('RevPAR','RevPAR',None,rv/av,'money'),('Receita quartos','Room revenue','K44',rv,'money'),('Receita F&B','F&B revenue','N44',sum(r['fb_revenue'] for r in selected),'money'),('Receita MICE','MICE revenue','E50',sum(r['mice_revenue'] for r in selected),'money'),('Receita total','Total revenue','H50',sum(r['revenue'] for r in selected),'money'),('EBITDA','EBITDA','N50',sum(r['ebitda'] for r in selected),'money')]
   for n,(a,b,cell,v,kind) in enumerate(items):
    r=10+(n//4)*5;c=(n%4)*3;text(ps,f'{col(c)}{r}:{col(c+2)}{r+1}',a,b,'label');f=f"'{owner}'!{cell}" if cell else f"IF('{owner}'!E44>0,'{owner}'!K44/('{owner}'!E44*365),0)";merge(ps,f'{col(c)}{r+2}:{col(c+2)}{r+3}',f,v,kind)
  else:
   for n in range(6):
    r=10+(n//3)*5;c=(n%3)*4;orr=7+(n//3)*6;oc=4+(n%3)*4;f,v=src.label(owner,f'{col(oc)}{orr}');merge(ps,f'{col(c)}{r}:{col(c+3)}{r+1}',f,v,'label');_,value=src.get(owner,f'{col(oc)}{orr+2}');kind=src.card_kind(owner,f'{col(oc)}{orr+2}');merge(ps,f'{col(c)}{r+2}:{col(c+3)}{r+3}',f"'{owner}'!{col(oc)}{orr+2}",value,kind)
  chart(ps,'A22',print_series[owner],width=1230,height=250)
  text(ps,'A39:H39','Indicador / base de análise','Indicator / analysis basis','thead');text(ps,'I39:L39','Valor','Value','thead')
  for n,(a,b,f,v,kind) in enumerate(print_records[owner][:7]):
   r=40+n*2;text(ps,f'A{r}:H{r+1}',a,b,'text')
   # Qualify only the hotel year selector; all data ranges are already qualified.
   f=f.replace('$N$23',"'DASH 05 V06'!$N$23");merge(ps,f'I{r}:L{r+1}',f,v,kind)
  text(ps,'A54:L57','Cenário financeiro simulado. Fonte: motor V06. Contratos, dados comerciais, fiscalidade e financiamento exigem validação documental.','Simulated financial scenario. Source: V06 engine. Contracts, commercial data, tax and funding require documentary validation.','note')
  merge(ps,'A59:L59',"'V06 Hipoteses'!D29&\" · \"&'V06 Hipoteses'!D28",'EFCC-20261022-01 · 2026-10-22','sub')
 # DASH 10 control centre; report pages retain their established technical name.
 s=frame(CENTER,9);text(s,'E18:P19','10 — Centro de relatório e impressão','10 — Report and print centre','section')
 for addr,oldsn,old,pt,en,kind,opts in [('H7','V06 Controlos','B5','Idioma global','Global language','inputtext',['PT','EN']),('H11','V06 Hipoteses','D26','Formato','Format','inputtext',['Complete Report','Executive Brief']),('N11','V06 Hipoteses','D27','Cenário do relatório','Report scenario','inputtext',['Active Scenario','Reference Case']),('H15','V06 Hipoteses','D28','Data referência','Reference date','inputtext',None),('N15','V06 Hipoteses','D29','Identificador cenário','Scenario identifier','inputtext',None)]:control(s,addr,oldsn,old,pt,en,kind,opts)
 merge(s,'E12:P13',f'HYPERLINK(IF($H$11="Executive Brief","#\'V06 Brief\'!A1","#\'DASH 10 V06\'!A1"),IF($H$11="Executive Brief",{tr("Abrir resumo selecionado →","Open selected brief →")},{tr("Abrir relatório integral selecionado →","Open selected complete report →")}))','Abrir relatório integral selecionado →','nav')
 text(s,'K7:M7','Estado','Status','sub');merge(s,'N7:P9',f'IF(ISERROR(SUM(\'V06 Motor\'!G5:EJ184)),{tr("Erro de cálculo","Calculation Error")},{tr("Validação pendente","Validation Pending")})','Validação pendente','note')
 for i,(a,b) in enumerate(SECTIONS):
  r=22+i*4;link(s,f'E{r}:P{r+2}',('DASH 10 V06',f'A{1+i*60}'),f'{i+1:02} — {a}',f'{i+1:02} — {b}')
 link(s,'E64:J66',('V06 Brief','A1'),'Resumo executivo · 2 páginas','Executive brief · 2 pages');link(s,'K64:P66',('DASH 10 V06','A1'),'Relatório integral · 10 páginas','Complete report · 10 pages');link(s,'E69:J71',('V06 Imprimir','A4'),'Procedimento PDF nativo','Native PDF procedure');link(s,'K69:P71',('V06 Registo Ativo','A4'),'Registo de parâmetros ativos','Active parameter register')
 text(s,'E75:P78','Estado: validação nativa pendente. Antes de exportar: Fórmulas → Calcular Agora; confirmar o cenário e guardar. Não se declara recálculo concluído por data ou valores em cache.','Status: native validation pending. Before exporting: Formulas → Calculate Now; confirm scenario and save. Dates or cached values do not prove calculation completion.','note');text(s,'E80:P84','O relatório integral tem paginação própria A4 vertical e não imprime a sidebar. Os painéis 01–09 têm área documental A3 horizontal, sem o menu lateral.','The complete report has its own A4 portrait pagination and excludes the sidebar. Panels 01–09 have A3 landscape document areas without the sidebar.','note')
 from premium_report import build_reports
 report_bindings=build_reports(wb,F,src,ref,data,chartcols)
 # Save specifications used by package integration and native validation.
 wb.get_worksheet_by_name('EXEC V06').activate();wb.close();spec={'source':str(BASE),'modules':str(out),'aliases':ALIASES,'inputs':INPUTS,'new_constant_inputs':ADDED_INPUTS,'navigation':NAV,'dashboard_names':[x[0] for x in DASHES],'palette':P,'report_bindings':report_bindings}
 (WORK/'Premium_Spec.json').write_text(json.dumps(spec,indent=2,ensure_ascii=False));print(out,len(ALIASES),'canonical input transfers',len(NAV),'links')
 return src,ref,spec
if __name__=='__main__':build()
