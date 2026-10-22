"""Refine ten executive views on frozen R1; preserve the financial engine.
Existing input addresses and hidden public KPI/table cells remain canonical.
Only existing sensitivity grid input coordinates acquire exact-current axes.
"""
from pathlib import Path
from zipfile import ZipFile,ZIP_DEFLATED
from lxml import etree as E
from collections import defaultdict
import copy,re,json,hashlib,posixpath
import xlsxwriter
from xlsxwriter.utility import xl_col_to_name as col,xl_cell_to_rowcol
from assemble import paths,T,xml,resolve,relpath
from premium_ui import Source,ROOT,DASHES,LANG,CENTER
from repair_empty_containers import normalize
BASE=Path('/workspace/kk01_efcc/premium_r2/BASE_R1.xlsx');WORK=BASE.parent
OUT=ROOT/'Premium_Checkpoint';OUT.mkdir(exist_ok=True)
MASTER=OUT/'KK01_Executive_Financial_Control_Center_CHECKPOINT.xlsx'
NSR='http://schemas.openxmlformats.org/officeDocument/2006/relationships';NSP='http://schemas.openxmlformats.org/package/2006/relationships';NSC='http://schemas.openxmlformats.org/package/2006/content-types'
CHART_KEYS=['cash','debt','reserve','capex','receipts','revenue','ebitda','construction','supervision','contingency_paid','equity','investor_outlay','draw','tax_paid','vat_refund','construction_insurance','operating_insurance','principal','sweep','interest','debt_service','deed_sales','sales_profit','units_deeded','units_not_deeded','res_revenue','res_opex','res_fees','owner_distribution','owner_target','res_ffe','res_fund','dividend','room_revenue','fb_revenue','mice_revenue','other_revenue','ebitda_margin']
H='DASH 05 V06';Q='DASH 07 V06';ADRS=[150,175,200,225];OCCS=[.50,.60,.65,.70,.75];OROW=[0,2,3,4,5]
def tr(pt,en):return f'IF({LANG}="PT","{pt.replace(chr(34),chr(34)*2)}","{en.replace(chr(34),chr(34)*2)}")'
def cellref(sn,addr):return f"'{sn}'!{addr}"
def bounds(area):
 a,b=(area.split(':')+[area])[:2] if ':' not in area else area.split(':');r,c=xl_cell_to_rowcol(a);r2,c2=xl_cell_to_rowcol(b);return r+1,c,r2+1,c2
class Sheet:
 def __init__(self,root,styles):self.r=root;self.s=styles;self.sd=root.find(T('sheetData'));self.rows={int(x.get('r')):x for x in self.sd};self.cells={c.get('r'):c for c in self.sd.iter(T('c'))}
 def row(self,n):
  if n not in self.rows:self.rows[n]=E.SubElement(self.sd,T('row'),r=str(n))
  return self.rows[n]
 def put(self,addr,value=None,formula=None,style=None):
  rownum,c=xl_cell_to_rowcol(addr);r=self.row(rownum+1);node=self.cells.get(addr)
  if node is None:node=E.SubElement(r,T('c'),r=addr);self.cells[addr]=node
  for x in list(node):node.remove(x)
  node.attrib.pop('t',None)
  if style is not None:node.set('s',str(self.s.get(style,style)))
  if formula is not None:
   E.SubElement(node,T('f')).text=formula.lstrip('=')
   if isinstance(value,str):node.set('t','str')
   E.SubElement(node,T('v')).text='' if value is None else str(value)
  elif isinstance(value,str):node.set('t','inlineStr');E.SubElement(E.SubElement(node,T('is')),T('t')).text=value
  elif value is not None:E.SubElement(node,T('v')).text=str(value)
  return node
 def merge(self,area,value=None,formula=None,style='text'):
  r,c,r2,c2=bounds(area);mc=self.r.find(T('mergeCells'))
  if mc is None:mc=E.SubElement(self.r,T('mergeCells'),count='0')
  for old in list(mc):
   a,b,d,e=bounds(old.get('ref'))
   if not (d<r or a>r2 or e<c or b>c2):mc.remove(old)
  for y in range(r,r2+1):
   for x in range(c,c2+1):self.put(col(x)+str(y),style=style)
  if r!=r2 or c!=c2:E.SubElement(mc,T('mergeCell'),ref=area)
  mc.set('count',str(len(mc)));return self.put(col(c)+str(r),value,formula,style)
 def text(self,area,pt,en,style='text'):return self.merge(area,pt,tr(pt,en),style)
 def link(self,area,sn,addr,pt,en):return self.merge(area,pt,f'HYPERLINK("#\'{sn}\'!{addr}",{tr(pt,en)})','nav')
 def finish(self):
  self.sd[:]=sorted(self.sd,key=lambda r:int(r.get('r')))
  for r in self.sd:r[:]=sorted(r,key=lambda c:xl_cell_to_rowcol(c.get('r'))[1])
  # ECMA-376 worksheet sequence; normalization prevents the R1 defect.
  order=['sheetPr','dimension','sheetViews','sheetFormatPr','cols','sheetData','sheetCalcPr','sheetProtection','protectedRanges','scenarios','autoFilter','sortState','dataConsolidate','customSheetViews','mergeCells','phoneticPr','conditionalFormatting','dataValidations','hyperlinks','printOptions','pageMargins','pageSetup','headerFooter','rowBreaks','colBreaks','customProperties','cellWatches','ignoredErrors','smartTags','drawing','legacyDrawing','legacyDrawingHF','picture','oleObjects','controls','webPublishItems','tableParts','extLst']
  priority=1
  for cf in self.r.findall(T('conditionalFormatting')):
   for rule in cf.findall(T('cfRule')):rule.set('priority',str(priority));priority+=1
  normalize(self.r);self.r[:]=sorted(self.r,key=lambda n:order.index(E.QName(n).localname) if E.QName(n).localname in order else 999)
  return xml(self.r)
def add_styles(root):
 fonts=root.find(T('fonts'));fills=root.find(T('fills'));xfs=root.find(T('cellXfs'));fmts=root.find(T('numFmts'));dxfs=root.find(T('dxfs'))
 if dxfs is None:dxfs=E.SubElement(root,T('dxfs'),count='0')
 styles={};nextfmt=max([163]+[int(n.get('numFmtId')) for n in fmts])+1
 def style(name,size=10,bg='FFFFFF',fg='202833',bold=False,align='left',nf=None):
  nonlocal nextfmt
  fid=len(fonts);f=E.SubElement(fonts,T('font'));E.SubElement(f,T('sz'),val=str(size));E.SubElement(f,T('color'),rgb='FF'+fg);E.SubElement(f,T('name'),val='Aptos');E.SubElement(f,T('family'),val='2')
  if bold:E.SubElement(f,T('b'))
  fillid=len(fills);p=E.SubElement(E.SubElement(fills,T('fill')),T('patternFill'),patternType='solid');E.SubElement(p,T('fgColor'),rgb='FF'+bg);E.SubElement(p,T('bgColor'),indexed='64')
  num=0
  if nf:num=nextfmt;nextfmt+=1;E.SubElement(fmts,T('numFmt'),numFmtId=str(num),formatCode=nf)
  i=len(xfs);xf=E.SubElement(xfs,T('xf'),numFmtId=str(num),fontId=str(fid),fillId=str(fillid),borderId='0',xfId='0',applyAlignment='1');E.SubElement(xf,T('alignment'),horizontal=align,vertical='center',wrapText='1');styles[name]=i
 style('bg',bg='F5F7FA');style('side',size=9.5,bg='252932',fg='FFFFFF');style('side_active',size=9.5,bg='C8101E',fg='FFFFFF',bold=True);style('side_num',size=10,bg='252932',fg='FFFFFF',bold=True,align='center');style('side_num_active',size=10,bg='C8101E',fg='FFFFFF',bold=True,align='center');style('brand',size=11,bg='C8101E',fg='FFFFFF',bold=True)
 style('header',size=21,bg='142337',fg='FFFFFF',bold=True);style('sub',size=10,fg='687586',bg='F5F7FA');style('section',size=13,bg='E9EFF5',fg='142337',bold=True);style('nav',size=9.5,bg='E9EFF5',fg='1765A8',bold=True);style('text',size=9.5);style('note',size=9,bg='F5F7FA',fg='687586');style('label',size=9.5,fg='687586');style('calculated',size=10,bg='F5F7FA',align='right');style('table_head',size=9.5,bg='142337',fg='FFFFFF',bold=True)
 for name,nf in [('money','#,##0 "€";[Red](#,##0) "€";–'),('pct','0.0%'),('count','#,##0'),('ratio','0.00"×"'),('adr','0 "€"'),('area','#,##0 "m²"')]:style(name,size=9.5,align='right',nf=nf)
 for name,nf in [('card_money','0.00,," M€";[Red](0.00,," M€");–'),('card_pct','0.0%'),('card_count','#,##0'),('card_ratio','0.00"×"'),('card_adr','0 "€"')]:style(name,size=20,fg='142337',bold=True,align='right',nf=nf)
 for name,nf in [('input_money','#,##0 "€"'),('input_pct','0.0%'),('input_count','0'),('input_ratio','0.00"×"'),('input_text',None)]:style(name,size=10,bg='E6F0FA',fg='1765A8',bold=True,align='right',nf=nf)
 style('calc_money',size=10,bg='F5F7FA',align='right',nf='0 "€"');style('calc_pct',size=10,bg='F5F7FA',align='right',nf='0.0%')
 style('card_text',size=11,bold=True,fg='142337',align='right');style('heat',size=11,bg='F5F7FA',align='right',nf='0.00,," M€";(0.00,," M€");–');style('input',size=10,bg='E6F0FA',fg='1765A8',bold=True,align='right',nf='#,##0 "€"')
 styles['dxf']={}
 for name,color in [('negative','FBE5E7'),('near','FFF0C2'),('positive','E0F1E5')]:
  styles['dxf'][name]=len(dxfs);d=E.SubElement(dxfs,T('dxf'));f=E.SubElement(d,T('font'));E.SubElement(f,T('color'),rgb='FF202833');p=E.SubElement(E.SubElement(d,T('fill')),T('patternFill'),patternType='solid');E.SubElement(p,T('fgColor'),rgb='FF'+color);E.SubElement(p,T('bgColor'),rgb='FF'+color)
 styles['dxf']['active']=len(dxfs);d=E.SubElement(dxfs,T('dxf'));b=E.SubElement(d,T('border'))
 for edge in ['left','right','top','bottom']:E.SubElement(E.SubElement(b,T(edge),style='medium'),T('color'),rgb='FFC8101E')
 for name,c in [('legend_negative','FBE5E7'),('legend_near','FFF0C2'),('legend_positive','E0F1E5')]:style(name,size=9,bg=c)
 for n in [fonts,fills,xfs,fmts,dxfs]:n.set('count',str(len(n)))
 return styles

def chart_assets(src,cc,rows):
 book=xlsxwriter.Workbook(WORK/'Design_Assets.xlsx');visual=book.add_worksheet('V06 Visual');keys=CHART_KEYS
 kc={key:240+i for i,key in enumerate(keys)};stocks={'cash','debt','reserve','units_not_deeded','res_fund'}
 for year in range(1,16):
  visual.write(year+3,0,year)
  for key,c in kc.items():
   if key=='ebitda_margin':
    numerator=sum(r['ebitda'] for r in rows[(year-1)*12:year*12]);den=sum(r['revenue'] for r in rows[(year-1)*12:year*12]);f=f'IF(SUM(\'V06 Motor\'!{col(cc["revenue"])}{(year-1)*12+5}:{col(cc["revenue"])}{year*12+4})>0,SUM(\'V06 Motor\'!{col(cc["ebitda"])}{(year-1)*12+5}:{col(cc["ebitda"])}{year*12+4})/SUM(\'V06 Motor\'!{col(cc["revenue"])}{(year-1)*12+5}:{col(cc["revenue"])}{year*12+4}),0)';v=numerator/den if den else 0
   else:
    a=col(cc[key]);f=f"'V06 Motor'!{a}{year*12+4}" if key in stocks else f"SUM('V06 Motor'!{a}{(year-1)*12+5}:{a}{year*12+4})";v=rows[year*12-1][key] if key in stocks else sum(r[key] for r in rows[(year-1)*12:year*12])
   visual.write_formula(year+3,c,'='+f,None,v)
 definitions=[
 [[('cash','Caixa','Cash'),('debt','Dívida','Debt'),('reserve','Reservas','Reserves')],[('capex','CAPEX','CAPEX'),('receipts','Recebimentos','Collections')],[('revenue','Receitas','Revenue'),('ebitda','EBITDA','EBITDA')]],
 [[('construction','Construção','Construction'),('supervision','Fiscalização','Supervision'),('contingency_paid','Contingência','Contingency')],[('capex','CAPEX','CAPEX'),('equity','Capital KK','KK equity'),('draw','Crédito utilizado','Drawdowns')],[('cash','Caixa','Cash'),('reserve','Reservas','Reserves')]],
 [[('tax_paid','IRPC','Corporate tax'),('vat_refund','IVA recuperado','VAT recovered')],[('construction_insurance','Seguro obra','Construction cover'),('operating_insurance','Seguro exploração','Operating cover')],[('revenue','Receita','Revenue'),('tax_paid','IRPC pago','Tax paid')]],
 [[('cash','Caixa','Cash'),('debt','Dívida','Debt')],[('draw','Desembolsos','Drawdowns'),('principal','Principal','Principal'),('sweep','Cash sweep','Cash sweep')],[('interest','Juros','Interest'),('debt_service','Serviço da dívida','Debt service')]],
 [[('deed_sales','Escrituras','Deed sales'),('receipts','Recebimentos','Collections')],[('units_deeded','Frações escrituradas','Units deeded')],[('sales_profit','Margem promoção · proxy','Development margin · proxy')]],
 [[('res_revenue','Receitas','Revenue'),('res_opex','OPEX','OPEX'),('res_fees','Fees','Fees')],[('owner_target','Yield alvo','Target income'),('owner_distribution','Rendas suportáveis','Supported income')],[('res_ffe','FF&E','FF&E'),('res_fund','Fundo residencial','Residential fund')]],
 [[('room_revenue','Quartos','Rooms'),('fb_revenue','F&B','F&B'),('mice_revenue','MICE','MICE'),('other_revenue','Outras','Other')],[('room_revenue','Quartos','Rooms'),('fb_revenue','F&B','F&B'),('mice_revenue','MICE','MICE'),('other_revenue','Outras','Other')],[('ebitda','EBITDA','EBITDA')]],
 [[('equity','Capital primário KK','Primary KK equity'),('investor_outlay','Capital investidores','Investor equity')],[('dividend','Dividendos KK','KK dividends')],[('cash','Caixa','Cash'),('reserve','Reservas','Reserves')]],
 [[('dividend','Dividendos','Dividends')],[('cash','Caixa','Cash'),('reserve','Reservas','Reserves')],[('investor_outlay','Capital investidores','Investor equity')]]]
 titles=[['Liquidez e dívida','Investimento e recebimentos','Receita e EBITDA'],['Orçamento de desenvolvimento','Fontes e aplicações','Liquidez protegida'],['Fiscalidade e recuperação IVA','Seguros','Receita e imposto'],['Caixa e dívida','Movimentos da dívida','Custo e serviço da dívida'],['Vendas e recebimentos','Escrituras por ano','Margem de promoção'],['Resultados residenciais','Yield dos proprietários','FF&E e reservas'],['Evolução das receitas','Composição da receita','EBITDA e margem'],['Subscrição de capital','Distribuições anuais','Liquidez e proteção'],['Dividendos anuais','Liquidez e reservas','Cash flow dos investidores']]
 counter=0
 for i,(sn,pt,en) in enumerate(DASHES[:9]):
  sheet=book.add_worksheet(sn)
  for j,series in enumerate(definitions[i]):
   counter+=1;is_mix=i==6 and j==1;is_waterfall=i==8 and j==2;typ='doughnut' if is_mix else 'column' if j==0 or is_waterfall else 'line';opts={'type':typ}
   if (i==6 and j==0) or is_waterfall:opts['subtype']='stacked'
   ch=book.add_chart(opts)
   if is_waterfall:
    previous=0
    for year in range(1,16):
     first,last=(year-1)*12+6,year*12+5;delta=sum(src.get('V06 F01','B'+str(r))[1] or 0 for r in range(first,last+1));end=previous+delta
     for k,(f,v) in enumerate([(f'MIN(SUM(\'V06 F01\'!B6:B{first-1}),SUM(\'V06 F01\'!B6:B{last}))' if year>1 else f'MIN(0,SUM(\'V06 F01\'!B6:B{last}))',min(previous,end)),(f'MAX(0,SUM(\'V06 F01\'!B{first}:B{last}))',max(0,delta)),(f'MAX(0,-SUM(\'V06 F01\'!B{first}:B{last}))',max(0,-delta))]):visual.write_formula(year+3,350+k,'='+f,None,v)
     previous=end
    for k,name in enumerate(['Base','Entradas / Inflows','Saídas / Outflows']):ch.add_series({'name':name,'categories':['V06 Visual',4,0,18,0],'values':['V06 Visual',4,350+k,18,350+k],'fill':{'none':True} if k==0 else {'color':'1765A8' if k==1 else 'C8101E'},'border':{'none':True}})
   else:
    for k,(key,a,b) in enumerate(series):
     c=kc[key];labelrow=30+counter*4+k;visual.write_formula(labelrow,c,'='+tr(a,b),None,a)
     if is_mix:
      selected=[r for r in rows if r['op_age']==src.get(H,'N23')[1]-1 and r['available_room_nights']>0];v=sum(r[key] for r in selected);a_col=col(cc[key]);f=f'SUMIFS(\'V06 Motor\'!{a_col}5:{a_col}184,\'V06 Motor\'!E5:E184,\'{H}\'!N23-1,\'V06 Motor\'!DL5:DL184,">0")';visual.write_formula(200+k,340,'='+tr(a,b),None,a);visual.write_formula(200+k,341,'='+f,None,v)
      if k==len(series)-1:ch.add_series({'categories':['V06 Visual',200,340,203,340],'values':['V06 Visual',200,341,203,341],'points':[{'fill':{'color':x}} for x in ['142337','1765A8','C8101E','687586']],'data_labels':{'percentage':True,'value':False,'category':False,'series_name':False,'legend_key':False,'position':'center','num_format':'0%','font':{'size':9,'color':'FFFFFF','bold':True}}})
     else:ch.add_series({'name':['V06 Visual',labelrow,c],'categories':['V06 Visual',4,0,18,0],'values':['V06 Visual',4,c,18,c],'line':{'color':['142337','1765A8','C8101E','687586'][k%4],'width':1.5},'fill':{'color':['142337','1765A8','C8101E','687586'][k%4]},'border':{'none':True}})
    if i==6 and j==2:
     second=book.add_chart({'type':'line'});second.add_series({'name':'EBITDA %','categories':['V06 Visual',4,0,18,0],'values':['V06 Visual',4,kc['ebitda_margin'],18,kc['ebitda_margin']],'y2_axis':True,'line':{'color':'C8101E','width':1.5}});second.set_y2_axis({'num_format':'0%','num_font':{'size':9}});ch.combine(second);ch.set_y2_axis({'num_format':'0%','num_font':{'size':9}})
   ch.set_size({'width':422,'height':187});ch.set_chartarea({'fill':{'color':'FFFFFF'},'border':{'none':True}});ch.set_plotarea({'border':{'none':True}});ch.set_legend({'position':'bottom','font':{'name':'Aptos','size':9}})
   if not is_mix:ch.set_x_axis({'num_font':{'name':'Aptos','size':9},'major_tick_mark':'none'});ch.set_y_axis({'num_format':'0' if i==4 and j==1 else '0,," M€"','num_font':{'size':9},'major_gridlines':{'visible':True,'line':{'color':'EDF1F5'}}})
   else:ch.set_hole_size(62);ch.set_plotarea({'layout':{'x':.12,'y':.05,'width':.75,'height':.75}})
   sheet.insert_chart('E57' if j==0 else 'L57' if j==1 else 'S57',ch,{'object_position':1})
 book.close();return titles

def run():
 src=Source(BASE)
 import gzip
 with gzip.open(ROOT/'Referencia_Python_V06_SYNC.json.gz','rt') as f:ref=json.load(f)
 cc=ref['main_columns'];rows=[{key:src.get('V06 Motor',col(c)+str(r))[1] for key,c in cc.items()} for r in range(5,185)];titles=chart_assets(src,cc,rows)
 with ZipFile(BASE) as z,ZipFile(WORK/'Design_Assets.xlsx') as assets:
  w,pp=paths(z);_,ap=paths(assets);changes={};added={};stylesroot=E.fromstring(z.read('xl/styles.xml'));styles=add_styles(stylesroot);routes=[];newcards=[];sensitivity=[]
  def route(sheet,area,dest,cell,a,b):sheet.link(area,dest,cell,a,b);routes.append({'origin':sheet.name+'!'+area.split(':')[0],'destination':dest+'!'+cell,'type':'Native HYPERLINK','PT':a,'EN':b})
  def projection(sheet,area,sn,addr,style='money'):
   _,v=src.get(sn,addr);sheet.merge(area,v,cellref(sn,addr),style);newcards.append({'sheet':sheet.name,'cell':area.split(':')[0],'source_sheet':sn,'source_cell':addr,'value':v})
  for i,(sn,pt,en) in enumerate(DASHES):
   sheet=Sheet(E.fromstring(z.read(pp[sn])),styles);sheet.name=sn;r=sheet.r;center=sn==CENTER;hidden=set() if center else set(range(7,22))|set(range(42,54))|set(range(65,85))
   # All canonical financial output/parameter cells stay at their original addresses.
   preserved={a:copy.deepcopy(c) for a,c in sheet.cells.items() if (a[0] not in ['A','B','C','D'] and int(re.search(r'\d+',a)[0]) in hidden) or (a[0] not in ['A','B','C','D'] and 22<=int(re.search(r'\d+',a)[0])<=39)}
   if not center:
    for row in list(sheet.sd):
     for c in list(row):
      a=c.get('r');rn,cn=xl_cell_to_rowcol(a)
      if cn<4 or (rn+1 not in hidden and not 22<=rn+1<=39):row.remove(c);sheet.cells.pop(a,None)
   else:
    for row in sheet.sd:
     for c in list(row):
      if xl_cell_to_rowcol(c.get('r'))[1]<4:row.remove(c);sheet.cells.pop(c.get('r'),None)
   mc=r.find(T('mergeCells'))
   if mc is not None:
    for m in list(mc):
     a,c,b,d=bounds(m.get('ref'))
     if c<4 or (not center and not(a in hidden and b in hidden) and not(22<=a<=b<=39)):mc.remove(m)
   for n in range(1,148 if i==6 else 118):
    row=sheet.row(n);row.set('ht','12');row.set('customHeight','1');row.attrib.pop('hidden',None)
    if n in hidden:row.set('hidden','1')
   for n in [2,3,4]:sheet.row(n).set('ht','18')
   # Uniform columns: 245px sidebar, 21-column executive grid.
   cols=r.find(T('cols'));cols[:]=[]
   for first,last,width in [(1,1,1),(2,2,4),(3,3,29),(4,4,2),(5,25,8.3),(26,26,1)]:E.SubElement(cols,T('col'),min=str(first),max=str(last),width=str(width),customWidth='1')
   for v in r.findall(T('sheetViews')+'/'+T('sheetView')):
    v.set('zoomScale','85');v.set('showGridLines','0');v.set('showRowColHeaders','0');v.set('topLeftCell','A1')
    for old in list(v):
     if old.tag in [T('pane'),T('selection')]:v.remove(old)
    E.SubElement(v,T('pane'),xSplit='4',ySplit='6',topLeftCell='E7',activePane='bottomRight',state='frozen');E.SubElement(v,T('selection'),pane='bottomRight',activeCell='E22' if not center else 'H7',sqref='E22' if not center else 'H7')
   # Identity is text until the original image is available; do not redraw a logo.
   sheet.merge('B2:C4','JS-CV\nCONSTRUÇÕES E INVESTIMENTOS, LDA.',style='brand');sheet.merge('E2:V3','KK-01 | HOTEL & RESIDENCES',style='header');route(sheet,'W2:Y3','V06 Identidade','B8','Logótipo original','Original logo');sheet.text('E4:Y4',f'{i+1:02} — {pt}',f'{i+1:02} — {en}','sub');sheet.merge('E1:Y1','PREMIUM R2 · CHECKPOINT',f'"PREMIUM R2 · "&\'V06 Hipoteses\'!D29&" · "&\'V06 Hipoteses\'!D28','note');route(sheet,'B5:C5',CENTER,'H7','PT / EN ▾','PT / EN ▾');route(sheet,'B6:C6',sn,'E6','Navegação','Navigation')
   for area,target,addr,a,b in [('E5:H5','EXEC V06','A1','Início','Home'),('I5:L5',sn,'H23' if not center and i else 'E26' if not center else 'H7','Parâmetros','Assumptions'),('M5:P5','V06 Dicionario','A4','Dicionário KPI','KPI dictionary'),('Q5:T5',f'V06 PDF {i+1:02}' if i<9 else 'V06 Imprimir','A1' if i<9 else 'A23','Imprimir / Exportar PDF','Print / Export PDF'),('U5:Y5',CENTER,'A1','Relatório 10','Report 10')]:route(sheet,area,target,addr,a,b)
   visible=[n for n in range(7,118) if n not in hidden]
   for j,(target,a,b) in enumerate(DASHES):
    first,last=visible[j*4],visible[j*4+3];sheet.merge(f'B{first}:B{last}',f'{j+1:02}',f'HYPERLINK("#\'{target}\'!A1","{j+1:02}")','side_num_active' if i==j else 'side_num');sheet.merge(f'C{first}:C{last}',a,f'HYPERLINK("#\'{target}\'!A1",{tr(a,b)})','side_active' if i==j else 'side');routes += [{'origin':sn+f'!B{first}','destination':target+'!A1','type':'Sidebar number'},{'origin':sn+f'!C{first}','destination':target+'!A1','type':'Sidebar title'}];route(sheet,f'{col(4+j)}6',target,'A1',f'{j+1:02}',f'{j+1:02}')
   for n in visible[:40]:
    for c in [0,3]:sheet.put(col(c)+str(n),style='side')
   if not center:
    # Controls retain their cells and their native validations/conditional formats.
    sheet.text('E22:P22','Pressupostos e controlos ativos','Active assumptions and controls','section')
    if i==6:
     sheet.merge('H32:J33',src.get(H,'H32')[1],style='input_text');sheet.text('E32:G33','Condições operadora','Operator terms','sub')
     sheet.text('Q22:Y22','Remuneração da operadora','International operator terms','section');projection(sheet,'Q23:Y24',H,'H32','card_text')
     for rr,lab,eng,source in [(26,'Base fee · receita','Base fee · revenue','H28'),(28,'Incentivo · GOP ajustado','Incentive · adjusted GOP','N30'),(30,'Marketing + distribuição + sistemas','Marketing + distribution + systems','N28')]:sheet.text(f'Q{rr}:V{rr+1}',lab,eng,'label');projection(sheet,f'W{rr}:Y{rr+1}',H,source,'pct')
     sheet.text('Q32:Y32','Comparação de condições · estudo','Terms comparison · study assumptions','section')
     for a,v in [('Q33','Perfil'),('T33','Base'),('V33','Inc.'),('X33','M+D+S')]:sheet.put(a,v,style='table_head')
     for rr,a,b,vals in [(34,'Proposta','Proposed',(.04,.10,.04)),(35,'KK Target','KK Target',(.025,.06,.03)),(36,'KK Optimized','KK Optimized',(.02,.05,.015))]:
      sheet.text(f'Q{rr}:S{rr}',a,b,'text')
      for c,value in zip(['T','V','X'],vals):sheet.merge(f'{c}{rr}:{col(xl_cell_to_rowcol(c+str(rr))[1]+1)}{rr}',value,style='pct')
     sheet.text('E39:P39','Azul: inputs. Cinzento: calculado. Custom só se aplica no modo Custom.','Blue: inputs. Grey: calculated. Stored Custom applies only in Custom mode.','note')
     sheet.text('Q38:Y39','Valores de estudo; validar contratos. Custom individualiza Marketing, Distribution e Systems nos controlos.','Study assumptions; validate contracts. Custom separates Marketing, Distribution and Systems in the controls.','note')
    else:
     sheet.text('Q22:Y22','Cenário ativo','Active scenario','section')
     for rr,a,b,sn2,addr,kind in [(24,'Hotel','Hotel','V06 Controlos','B67','card_text'),(27,'ADR · abertura','Opening ADR',H,'H24','adr'),(30,'Ocupação · abertura','Opening occupancy',H,'H26','pct'),(33,'Taxa de desconto','Discount rate',Q,'N26','pct'),(36,'Política de reservas','Reserve policy',Q,'H23','card_text')]:sheet.text(f'Q{rr}:V{rr+1}',a,b,'label');projection(sheet,f'W{rr}:Y{rr+1}',sn2,addr,kind)
    if i==6:cards=[('Quartos disponíveis','Available rooms',H,'E44','count'),('Dormidas vendidas','Occupied room nights',H,'H44','count'),('RevPAR · €/quarto disponível','RevPAR · €/available room',H,'K50','adr'),('Receita de quartos','Room revenue',H,'K44','money'),('Receita F&B','F&B revenue',H,'N44','money'),('Receita MICE','MICE revenue',H,'E50','money'),('Receita total','Total revenue',H,'H50','money')]
    else:
     cards=[]
     for a in ['E9','I9','M9','E15','I15','M15']:
      rr,cn=xl_cell_to_rowcol(a);labeladdr=col(cn)+str(rr-1);f,label=src.label(sn,labeladdr);cards.append((label,label,sn,a,src.card_kind(sn,a)))
     cards.append(('Perfil operacional','Operating profile','V06 Controlos','B67','text'))
    for j,(a,b,sn2,addr,kind) in enumerate(cards):
     c=4+j*3;area=f'{col(c)}40:{col(c+2)}41'
     if i==6 or j==6:sheet.text(area,a,b,'label')
     else:f,label=src.label(sn,col(xl_cell_to_rowcol(addr)[1])+str(xl_cell_to_rowcol(addr)[0]-1));sheet.merge(area,label,f,'label')
     projection(sheet,f'{col(c)}54:{col(c+2)}56',sn2,addr,'card_'+('text' if kind=='text' else kind))
    for j,title in enumerate(titles[i]):sheet.text(f'{col(4+j*7)}57:{col(10+j*7)}57',title,[['Liquidity and debt','Investment and collections','Revenue and EBITDA'],['Development budget','Sources and uses','Protected liquidity'],['Taxes and VAT recovery','Insurance','Revenue and tax'],['Cash and debt','Debt movements','Debt cost and service'],['Sales and collections','Annual unit deeds','Development margin'],['Residential results','Owner yield','FF&E and reserves'],['Revenue evolution','Revenue composition','EBITDA and margin'],['Equity subscriptions','Annual distributions','Liquidity and protection'],['Annual dividends','Liquidity and reserves','Investor cash flow']][i][j],'section')
    # Three aligned financial tables, all linked to existing canonical outputs.
    if i==6:
     tables=[('Resultados operacionais','Operating results',[('Receita de quartos','Room revenue',H,'K44','money'),('Receita F&B','F&B revenue',H,'N44','money'),('Receita MICE','MICE revenue',H,'E50','money'),('Outras receitas','Other revenue',H,'M84','money'),('Receita total','Total revenue',H,'H50','money'),('GOP','GOP',H,'M72','money'),('Encargos da operadora','Operator fees',H,'M74','money'),('EBITDA','EBITDA',H,'N50','money')]),('Métricas de rentabilidade','Profitability metrics',[('RevPAR','RevPAR',H,'K50','adr'),('Margem EBITDA','EBITDA margin',H,'M80','pct'),('ADR · abertura','Opening ADR',H,'H24','adr'),('Ocupação · abertura','Opening occupancy',H,'H26','pct'),('Utilização MICE','MICE utilisation',H,'N25','pct'),('Crescimento ADR','ADR growth',H,'N29','pct')]),('Reservas e proteção','Reserves and protection',[('Dotação FF&E · ano','Annual FF&E contribution',H,'M76','money'),('Utilização FF&E · ano','Annual FF&E replacement',H,'M78','money'),('Saldo FF&E · fim ano','FF&E fund · year end',H,'M82','money'),('Taxa FF&E','FF&E rate',H,'N27','pct'),('Reservas gerais','General reserve','V06 Motor','EB184','money')])]
    else:
     oldkeys={0:['CAPEX futuro','Capital KK previsto','Serviço da dívida'],1:['Construção e FF&E','Fiscalização','Contingência','ETC adicional'],2:['IRPC pago','Reembolsos IVA','Seguro obra','Seguro exploração'],3:['Desembolsos','Juros','Amortização contratual','Cash sweep'],4:['Recebimentos','Escrituras','Custos comerciais','Margem promoção · proxy'],5:['Receitas residenciais','OPEX','Fees residenciais','Rendas suportáveis','Yield alvo não coberto','FF&E'],7:['Capital KK previsto','Capital novos investidores','Dividendos KK'],8:['Dividendos KK','Juros da dívida','Impostos pagos']}[i]
     tables=[('Resultados financeiros','Financial results',[(x,x,sn,'M'+str(66+j*2),'money') for j,x in enumerate(oldkeys)]),('Indicadores executivos','Executive indicators',[(str(src.get(sn,col(xl_cell_to_rowcol(a)[1])+str(xl_cell_to_rowcol(a)[0]-1))[1]),str(src.get(sn,col(xl_cell_to_rowcol(a)[1])+str(xl_cell_to_rowcol(a)[0]-1))[1]),sn,a,src.card_kind(sn,a)) for a in ['E9','I9','M9','E15','I15','M15']]),('Condições ativas','Active conditions',[('ADR · abertura','Opening ADR',H,'H24','adr'),('Ocupação · abertura','Opening occupancy',H,'H26','pct'),('MICE','MICE',H,'N25','pct'),('FF&E','FF&E',H,'N27','pct'),('Desconto','Discount',Q,'N26','pct')])]
    for j,(a,b,records) in enumerate(tables):
     c=4+j*7;sheet.text(f'{col(c)}90:{col(c+6)}91',a,b,'section')
     for k,(a,b,sn2,addr,kind) in enumerate(records):
      rr=92+k*2
      if i!=6 and j==0:
       source_label='E'+str(66+k*2);sheet.merge(f'{col(c)}{rr}:{col(c+3)}{rr+1}',src.get(sn,source_label)[1],cellref(sn,source_label),'text')
      elif i!=6 and j==1:
       f,v=src.label(sn2,col(xl_cell_to_rowcol(addr)[1])+str(xl_cell_to_rowcol(addr)[0]-1));sheet.merge(f'{col(c)}{rr}:{col(c+3)}{rr+1}',v,f,'text')
      else:sheet.text(f'{col(c)}{rr}:{col(c+3)}{rr+1}',a,b,'text')
      projection(sheet,f'{col(c+4)}{rr}:{col(c+6)}{rr+1}',sn2,addr,kind)
    sheet.text('E110:Y111','Pressupostos de estudo. Valores provenientes do motor central. Consultar dicionário para bases e unidades.','Study assumptions. Values come from the central engine. See the dictionary for definitions and units.','note');route(sheet,'E112:K112','V06 Dicionario','A4','Dicionário de KPIs','KPI dictionary');route(sheet,'L112:R112','EXEC V06','A1','Início','Home');route(sheet,'S112:Y112',CENTER,'A1','Relatório consolidado','Consolidated report')
   else:
    sheet.text('Q8:Y9','Relatório de investimento','Investment report','section');sheet.text('Q11:Y16','Selecione idioma, formato e caso; abra a versão selecionada. Antes de imprimir, recalcule e confirme os pressupostos.','Select language, format and case, then open the selected report. Recalculate and confirm assumptions before printing.','text');route(sheet,'Q18:Y20','DASH 10 V06','A1','Pré-visualizar relatório · 10 páginas','Preview report · 10 pages');route(sheet,'Q22:Y24','V06 Brief','A1','Executive Brief · 2 páginas','Executive Brief · 2 pages');route(sheet,'Q26:Y28','V06 Imprimir','A4','Imprimir / Exportar PDF','Print / Export PDF');sheet.text('Q31:Y36','Checkpoint: validação Excel 365 e ficheiro de logótipo original pendentes.','Checkpoint: native Excel 365 validation and original logo file pending.','note')
   # Units and data-entry appearance, without changing any input or formula.
   for item in json.loads((ROOT/'Parametros_Interface.json').read_text()):
    if item['sheet']!=sn:continue
    address=item['cell'];value=item['value'];kind=src.format_kind(sn,address)
    style='input_text' if isinstance(value,str) else 'input_pct' if kind=='percent' else 'input_ratio' if any(word in item.get('label_pt','').lower() for word in ['múltiplo','cobertura']) else 'input_count' if (sn==H and address in ['N23','H36']) or any(word in item.get('label_pt','').lower() for word in ['meses','prazo','carência','ano operacional','anos','0/1']) else 'input_money'
    sheet.cells[address].set('s',str(styles[style]))
   if sn==H:
    sheet.cells['H24'].set('s',str(styles['calc_money']))
    for address in ['H26','H28','N28','N30']:sheet.cells[address].set('s',str(styles['calc_pct']))
    sheet.text('E25:G25','ADR Custom · €/noite','Custom ADR · €/night','label')
   # Remove obsolete blank native link containers; retain formula hyperlinks.
   hs=r.find(T('hyperlinks'))
   if hs is not None:r.remove(hs)
   dim=r.find(T('dimension'));dim.set('ref','A1:Y147' if i==6 else 'A1:Y117');changes[pp[sn]]=sheet.finish()
  # Reuse the existing 30 full scenario engines: 20 displayed combinations,
  # plus exact-current ADR and occupancy axes, without another financial engine.
  technical=Sheet(E.fromstring(z.read(pp['V06 Matriz ADR'])),styles);technical.name='V06 Matriz ADR';technical.put('B8',src.get(H,'I9')[1],cellref(H,'I9'));technical.put('G6',src.get(H,'E9')[1],cellref(H,'E9'));technical.text('A15:G17','A linha adicional usa a ocupação ativa; a coluna adicional usa o ADR ativo. Os 20 pontos do dashboard permanecem fixos.','The extra row uses active occupancy; the extra column uses active ADR. The 20 dashboard grid points stay fixed.','note');changes[pp['V06 Matriz ADR']]=technical.finish()
  # Dynamic labels for the nine reused technical variants, keeping scenario IDs.
  scenarios=Sheet(E.fromstring(z.read(pp['V06 Cenarios'])),styles)
  for ri in range(6):
   for ci in range(5):
    if ri==1 or ci==4:
     rr=30+ri*5+ci;scenarios.put('A'+str(rr),f'ADR current / occupancy grid {ri+1}',f'"ADR "&\'V06 Matriz ADR\'!{col(ci+2)}6&" | Occ "&TEXT(\'V06 Matriz ADR\'!B{ri+7},"0.0%")')
  changes[pp['V06 Cenarios']]=scenarios.finish()
  def npv(ri,ci):return "'V06 Matriz Valores'!C"+str(5+ri*5+ci)
  h=Sheet(E.fromstring(changes[pp[H]]),styles);h.name=H
  h.text('E114:Y115','Sensibilidade do Valor do Investimento','Investment Value Sensitivity','section');h.text('E116:Q116','ADR × Ocupação → VAL dos Investidores','ADR × Occupancy → Investor NPV','sub');h.text('R116:V116','Limiar perto de zero','Near-zero threshold','label');h.merge('W116:Y116',500000,style='input')
  dv=h.r.find(T('dataValidations'))
  if dv is None:dv=E.SubElement(h.r,T('dataValidations'))
  rule=E.SubElement(dv,T('dataValidation'),type='decimal',operator='between',allowBlank='0',showErrorMessage='1',errorStyle='stop',sqref='W116');E.SubElement(rule,T('formula1')).text='0';E.SubElement(rule,T('formula2')).text='10000000';dv.set('count',str(len(dv)))
  h.text('E118:H119','Ocup. / ADR','Occup. / ADR','table_head')
  for ci,adr in enumerate(ADRS):h.merge(f'{col(8+ci*4)}118:{col(11+ci*4)}119',adr,style='adr')
  def add_heat(sheet,area,formula,value,active=None,threshold="KK_NPV_NEAR_ZERO"):
   sheet.merge(area,value,formula,'heat');cf=E.SubElement(sheet.r,T('conditionalFormatting'),sqref=area)
   if active:
    r=E.SubElement(cf,T('cfRule'),type='expression',dxfId=str(styles['dxf']['active']),priority=str(1+len(sensitivity)*4));E.SubElement(r,T('formula')).text=active
   addr=area.split(':')[0];a='$'+re.match('[A-Z]+',addr)[0]+'$'+re.search('[0-9]+',addr)[0]
   for k,(name,expr) in enumerate([('negative',a+'<-'+threshold),('near','ABS('+a+')<='+threshold),('positive',a+'>'+threshold)],2):
    r=E.SubElement(cf,T('cfRule'),type='expression',dxfId=str(styles['dxf'][name]),priority=str(k+len(sensitivity)*4),stopIfTrue='1');E.SubElement(r,T('formula')).text=expr
   sensitivity.append({'sheet':sheet.name,'cell':addr,'source':formula,'initial_value':value})
  nearest_adr=f'IF(\'{H}\'!$E$9<162.5,150,IF(\'{H}\'!$E$9<187.5,175,IF(\'{H}\'!$E$9<212.5,200,225)))';nearest_occ=f'IF(\'{H}\'!$I$9<0.55,0.5,IF(\'{H}\'!$I$9<0.625,0.6,IF(\'{H}\'!$I$9<0.675,0.65,IF(\'{H}\'!$I$9<0.725,0.7,0.75))))'
  for ri,occ in enumerate(OCCS):
   rr=120+ri*2;h.merge(f'E{rr}:H{rr+1}',occ,style='pct')
   for ci,adr in enumerate(ADRS):
    formula=npv(OROW[ri],ci);v=src.get('V06 Matriz Valores','C'+str(5+OROW[ri]*5+ci))[1];area=f'{col(8+ci*4)}{rr}:{col(11+ci*4)}{rr+1}';active=f'AND({adr}={nearest_adr},{occ}={nearest_occ})'.replace("'"+H+"'!",'');add_heat(h,area,formula,v,active)
  h.text('E131:K132','Negativo: VAL < −limiar','Negative: NPV < −threshold','legend_negative');h.text('L131:R132','Perto de 0: |VAL| ≤ limiar','Near 0: |NPV| ≤ threshold','legend_near');h.text('S131:Y132','Positivo: VAL > limiar','Positive: NPV > threshold','legend_positive')
  outside=tr('Não atingido na grelha','Not reached within tested grid')
  minadr=outside
  for ci in reversed(range(4)):minadr=f'IF({npv(1,ci)}>=0,{ADRS[ci]},{minadr})'
  minocc=outside
  for ri in reversed(range(5)):minocc=f'IF({npv(OROW[ri],4)}>=0,{OCCS[ri]},{minocc})'
  be=[('E133:I134','ADR mínimo · ocupação ativa','Minimum ADR · active occupancy','E135:I137',minadr,'Não atingido na grelha','adr'),('J133:N134','Ocupação mínima · ADR ativo','Minimum occupancy · active ADR','J135:N137',minocc,'Não atingido na grelha','pct'),('O133:S134','VAL do cenário atual','Current scenario investor NPV','O135:S137',cellref(Q,'M15'),src.get(Q,'M15')[1],'money'),('T133:Y134','Défice de VAL para zero','NPV gap to break-even','T135:Y137',f'MAX(0,-\'{Q}\'!M15)',max(0,-src.get(Q,'M15')[1]),'money')]
  for area,a,b,out,f,v,kind in be:h.text(area,a,b,'label');h.merge(out,v,f,kind)
  h.text('E139:Y141','Grelha discreta; equilíbrio testado nos pontos disponíveis, sem extrapolação. Contorno vermelho: ponto mais próximo; o VAL atual exato está acima. Mudar apenas ADR/ocupação atuais move o marcador; combinações fixas substituem esses dois inputs.','Discrete grid; break-even is tested at available points without extrapolation. Red outline: nearest point; exact current NPV is shown above. Changing only current ADR/occupancy moves the marker; fixed combinations override those two inputs.','note');h.text('E139:Y143','Estrutura conceptual aprovada. A versão Excel deve mostrar VAL reais por combinação ADR/ocupação, com os pressupostos ativos do modelo. Grelha discreta, sem extrapolação; o contorno marca o ponto mais próximo.','Conceptual structure approved. Final Excel version must display actual NPV values by ADR and occupancy combination based on the active model assumptions. Discrete grid without extrapolation; the outline marks the nearest point.','note');changes[pp[H]]=h.finish()
  # Replace the two old fixed-rate print grids with the same active-rate 20 cells.
  report=Sheet(E.fromstring(z.read(pp['DASH 10 V06'])),styles);report.name='DASH 10 V06';mc=report.r.find(T('mergeCells'))
  if mc is not None:
   for m in list(mc):
    a,c,b,d=bounds(m.get('ref'))
    if 501<=a<=b<=528:mc.remove(m)
  for row in report.sd:
   if 501<=int(row.get('r'))<=528:
    for c in list(row):row.remove(c);report.cells.pop(c.get('r'),None)
  for cf in list(report.r.findall(T('conditionalFormatting'))):
   if any(501<=int(n)<=528 for n in re.findall('[0-9]+',cf.get('sqref',''))):report.r.remove(cf)
  report.text('A501:L502','Sensibilidade do Valor do Investimento · ADR × Ocupação','Investment Value Sensitivity · ADR × Occupancy','section')
  params=f'"ADR "&\'{H}\'!H24&" € | Occ "&TEXT(\'{H}\'!H26,"0.0%")&" | MICE "&TEXT(\'{H}\'!N25,"0%")&" | FF&E "&TEXT(\'{H}\'!N27,"0%")&" | Discount "&TEXT(\'{Q}\'!N26,"0%")'
  report.merge('A503:L504','ADR 185 € | Occ 58% | MICE 50% | FF&E 5% | Discount 12%',params,'note');report.merge('A505:L505','Custom Terms',f'\'{H}\'!H32&" | Base "&TEXT(\'{H}\'!H28,"0.0%")&" | M+D+S "&TEXT(\'{H}\'!N28,"0.0%")&" | Incentive "&TEXT(\'{H}\'!N30,"0.0%")','note');report.text('A506:B507','Ocup./ADR','Occup./ADR','table_head')
  for ci,adr in enumerate(ADRS):report.merge(f'{col(2+ci*2)}506:{col(3+ci*2)}507',adr,style='adr')
  for ri,occ in enumerate(OCCS):
   rr=508+ri*2;report.merge(f'A{rr}:B{rr+1}',occ,style='pct')
   for ci,adr in enumerate(ADRS):
    f=cellref(H,col(8+ci*4)+str(120+ri*2));v=src.get('V06 Matriz Valores','C'+str(5+OROW[ri]*5+ci))[1];add_heat(report,f'{col(2+ci*2)}{rr}:{col(3+ci*2)}{rr+1}',f,v)
  for area,a,b,kind in [('K506:L507','VAL / NPV','NPV','table_head'),('K508:L510','Negativo','Negative','legend_negative'),('K511:L513','Perto de 0','Near 0','legend_near'),('K514:L517','Positivo','Positive','legend_positive')]:report.text(area,a,b,kind)
  report.text('A519:L522','20 combinações calculadas com o motor de 180 meses e pressupostos ativos, incluindo desconto, fiscalidade, financiamento, fees e reservas. Perto de zero tem precedência em ±limiar. A matriz é do caso operacional ativo, mesmo se o relatório mostrar a referência congelada.','20 combinations calculated with the 180-month engine and active discount, taxes, financing, fees and reserves. Near-zero takes precedence within ±threshold. This matrix uses the active operating case, even when the report displays the frozen reference.','note')
  for lab,area,a,b,f,v,kind in [('A523:F524','A525:F526','ADR mínimo · ocupação ativa','Minimum ADR · active occupancy',minadr,'Não atingido na grelha','adr'),('G523:L524','G525:L526','Ocupação mínima · ADR ativo','Minimum occupancy · active ADR',minocc,'Não atingido na grelha','pct')]:report.text(lab,a,b,'label');report.merge(area,v,f,kind)
  report.text('A527:F527','VAL atual / défice até zero','Current NPV / gap to zero','label');report.merge('A528:F528',src.get(Q,'M15')[1],cellref(Q,'M15'),'money');report.merge('G528:L528',max(0,-src.get(Q,'M15')[1]),f'MAX(0,-\'{Q}\'!M15)','money');changes[pp['DASH 10 V06']]=report.finish()
  # Merge only drawing/chart assets. No engine worksheets are rebuilt.
  chartids=[int(re.search(r'chart(\d+)\.xml$',p)[1]) for p in z.namelist() if re.fullmatch('xl/charts/chart[0-9]+.xml',p)];nextid=max(chartids)+1;mapping={}
  for p in sorted(assets.namelist()):
   if re.fullmatch('xl/charts/chart[0-9]+.xml',p):
    mapping[p]=f'xl/charts/chart{nextid}.xml';nextid+=1;chartroot=E.fromstring(assets.read(p));chartns={'c':'http://schemas.openxmlformats.org/drawingml/2006/chart'}
    if chartroot.find('.//c:doughnutChart',chartns) is not None:
     labels=chartroot.find('.//c:dLbls',chartns)
     for tag,value in [('showLegendKey','0'),('showVal','0'),('showCatName','0'),('showSerName','0'),('showPercent','1'),('showBubbleSize','0')]:
      node=labels.find('{'+chartns['c']+'}'+tag)
      if node is None:node=E.SubElement(labels,'{'+chartns['c']+'}'+tag)
      node.set('val',value)
     order=['dLbl','delete','numFmt','spPr','txPr','dLblPos','showLegendKey','showVal','showCatName','showSerName','showPercent','showBubbleSize','separator','showLeaderLines','leaderLines','extLst']
     labels[:]=sorted(labels,key=lambda n:order.index(E.QName(n).localname))
    added[mapping[p]]=xml(chartroot)
  ct=E.fromstring(z.read('[Content_Types].xml'))
  for p in mapping.values():E.SubElement(ct,'{'+NSC+'}Override',PartName='/'+p,ContentType='application/vnd.openxmlformats-officedocument.drawingml.chart+xml')
  for sn,_,_ in DASHES[:9]:
   ar=E.fromstring(assets.read(ap[sn]));d=ar.find(T('drawing'));ars=E.fromstring(assets.read(relpath(ap[sn])));dr=next(x for x in ars if x.get('Id')==d.get('{'+NSR+'}id'));adraw=resolve(ap[sn],dr.get('Target'))
   target=Sheet(E.fromstring(changes[pp[sn]]),styles);dr0=target.r.find(T('drawing'));rr=E.fromstring(z.read(relpath(pp[sn])));oldrel=next(x for x in rr if x.get('Id')==dr0.get('{'+NSR+'}id'));owner=resolve(pp[sn],oldrel.get('Target'));drawing=E.fromstring(assets.read(adraw))
   # Leave room for chart section title in row 57.
   ns={'xdr':'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing'}
   for anchor in drawing:
    start=anchor.find('xdr:from/xdr:row',ns);end=anchor.find('xdr:to/xdr:row',ns)
    if start is not None:start.text=str(int(start.text)+1)
    if end is not None:end.text='88'
    offset=anchor.find('xdr:to/xdr:rowOff',ns)
    if offset is not None:offset.text=str(11*9525)
   changes[owner]=xml(drawing);rels=E.fromstring(assets.read(relpath(adraw)))
   for rel in rels:
    source=resolve(adraw,rel.get('Target'));rel.set('Target',posixpath.relpath(mapping[source],posixpath.dirname(owner)))
   changes[relpath(owner)]=xml(rels)
  # Copy added formula projections with string remapping; original data retained.
  visual=Sheet(E.fromstring(z.read(pp['V06 Visual'])),styles);ss=[''.join(n.itertext()) for n in E.fromstring(assets.read('xl/sharedStrings.xml'))] if 'xl/sharedStrings.xml' in assets.namelist() else []
  vr=E.fromstring(assets.read(ap['V06 Visual']))
  for c in vr.iter(T('c')):
   addr=c.get('r');_,cn=xl_cell_to_rowcol(addr)
   if cn<240:continue
   node=copy.deepcopy(c);node.set('s',str(styles['money']))
   if node.get('t')=='s':v=ss[int(node.findtext(T('v')))];node.attrib.pop('t');node.remove(node.find(T('v')));node.set('t','inlineStr');E.SubElement(E.SubElement(node,T('is')),T('t')).text=v
   old=visual.cells.get(addr)
   if old is not None:old.getparent().remove(old)
   visual.row(xl_cell_to_rowcol(addr)[0]+1).append(node);visual.cells[addr]=node
  changes[pp['V06 Visual']]=visual.finish();changes['xl/styles.xml']=xml(stylesroot);changes['[Content_Types].xml']=xml(ct)
  # Executive print view omits technical hidden ranges and sidebar by default.
  dn=w.find(T('definedNames'));E.SubElement(dn,T('definedName'),name='KK_NPV_NEAR_ZERO').text="'DASH 05 V06'!$W$116";names=[s.get('name') for s in w.find(T('sheets'))]
  for n in dn:
   if n.get('name')=='_xlnm.Print_Area' and n.get('localSheetId') is not None:
    sn=names[int(n.get('localSheetId'))]
    if sn in [x[0] for x in DASHES[:9]]:n.text=f"'{sn}'!$E$1:$Y${147 if sn==H else 117}"
  cp=w.find(T('calcPr'))
  if cp is not None:cp.set('calcMode','auto');cp.set('fullCalcOnLoad','1');cp.set('forceFullCalc','1')
  changes['xl/workbook.xml']=xml(w)
  building=MASTER.with_suffix('.building.xlsx')
  with ZipFile(building,'w',ZIP_DEFLATED,compresslevel=6) as out:
   for item in z.infolist():out.writestr(copy.copy(item),changes.get(item.filename,z.read(item.filename)))
   for p,data in added.items():out.writestr(p,data)
  building.replace(MASTER)
  # Strong boundary: every original engine part remains byte-identical.
  with ZipFile(MASTER) as result:
   engineparts=[p for sn,p in pp.items() if re.fullmatch(r'V06 Motor|V06 Referencia|V06 [SMO]\d+|V06 F\d+',sn)]
   assert all(z.read(p)==result.read(p) for p in engineparts)
   for input in json.loads((ROOT/'Parametros_Interface.json').read_text()):
    sn,a=input['sheet'],input['cell'];before=src.get(sn,a);after=Source(MASTER).get(sn,a);assert before==after,(sn,a,before,after)
   assert result.testzip() is None
  spec={'release':'PREMIUM-R2-CHECKPOINT','base_sha256':hashlib.sha256(BASE.read_bytes()).hexdigest(),'result_sha256':hashlib.sha256(MASTER.read_bytes()).hexdigest(),'engine_and_cashflow_parts_byte_identical':len(engineparts),'canonical_inputs_preserved':50,'new_ui_input':H+'!W116 (near-zero threshold)','layout':{'font':'Aptos','zoom':85,'freeze':'6 rows / 4 columns','hidden_public_output_rows':'7:21, 42:53, 65:84; existing results retained and displayed by direct links','grid':'E:Y; three charts and three financial tables','sidebar':'separate number/title columns, one selected language'},'native_excel365':'NOT TESTED','official_logo':'NOT TESTED — original image bytes not available','routes':routes,'chart_keys':CHART_KEYS,'display_bindings':newcards,'sensitivity':sensitivity,'sensitivity_exact_axes':{'occupancy':'V06 Matriz ADR!B8 -> DASH 05 V06!I9','adr':'V06 Matriz ADR!G6 -> DASH 05 V06!E9','exact_current_engine':'V06 M10','displayed_matrix':'20 fixed ADR/occupancy combinations; active parameters shared with existing 180-month engines'}}
  (OUT/'Design_Spec.json').write_text(json.dumps(spec,indent=2,ensure_ascii=False));print(MASTER,len(newcards),'display bindings',len(sensitivity),'matrix cells',len(routes),'navigation routes')
if __name__=='__main__':run()
