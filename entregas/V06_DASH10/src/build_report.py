"""Add a reporting-only sheet to the current V06. No financial engine generation.
Financial cells reference existing dashboards, schedules and controls exclusively.
"""
from pathlib import Path
import json,sys,hashlib,shutil,os
import openpyxl,xlsxwriter
from xlsxwriter.utility import xl_col_to_name as col
ROOT=Path(__file__).resolve().parents[1];WORK=Path(os.environ.get('KK01_REPORT_WORKDIR','/workspace/kk01_dash10'));WORK.mkdir(exist_ok=True)
BASE=ROOT.parent/'V06/KK01_V06_Executive_Master.xlsx'
REF=json.loads((ROOT/'Esquema_Fontes.json').read_text())
SECTIONS=[('Síntese Executiva de Investimento','Executive Investment Summary'),('Programa, Orçamento e Financiamento','Project Scope, Development Budget & Funding'),('Promoção Imobiliária e Recebimentos','Real Estate Development, Sales & Collections'),('Exploração Hoteleira e Centro MICE','Hotel & MICE Operating Performance'),('Exploração Residencial e Sell & Leaseback','Residential Operations & Sell & Leaseback'),('Fiscalidade, Incentivos e Seguros','Taxation, Incentives & Insurance'),('Tesouraria, Dívida e Reservas Protegidas','Treasury, Debt & Protected Reserves'),('Acionistas, Capital e Dividendos','Shareholders, Capital & Dividends'),('Retornos, Valor Terminal e Viabilidade','Returns, Terminal Value & Feasibility'),('Sensibilidades, Riscos e Condições de Investimento','Sensitivities, Risks & Investment Conditions')]
class Report:
 def __init__(self,base=BASE):
  self.base=Path(base);self.source=openpyxl.load_workbook(self.base,read_only=True,data_only=True);self.sources={};self.cells={};self.bindings=[]
  self.wb=xlsxwriter.Workbook(str(WORK/'DASH10_modulo.xlsx'));self.wb.set_calc_mode('auto');self.ws=self.wb.add_worksheet('DASH 10 V06');self.ws.hide_gridlines(2);self.ws.hide_row_col_headers();self.ws.set_zoom(100);self.ws.freeze_panes(6,0)
  self.ws.set_column('A:L',7);self.ws.set_default_row(14);self.ws.set_portrait();self.ws.set_paper(9);self.ws.set_print_scale(100);self.ws.set_margins(.3,.3,.3,.3);self.ws.set_header('&LJSCV · KK-01&R V06 · DASH 10',{'margin':.1});self.ws.set_footer('&LPT / EN · Excel 365: native test pending&R&P / &N',{'margin':.1});self.ws.print_area('A1:L500');self.ws.set_h_pagebreaks([50*i for i in range(1,10)])
  styles={'brand':{'bold':True,'font_color':'FFFFFF','bg_color':'B31F2B','font_size':13},'title':{'font_size':18,'bold':True,'text_wrap':True,'valign':'vcenter'},'section':{'font_size':15,'bold':True,'font_color':'B31F2B','text_wrap':True},'sub':{'font_size':9,'font_color':'69717C','text_wrap':True},'label':{'font_size':10,'font_color':'69717C','bg_color':'F5F6F8','text_wrap':True,'valign':'vcenter'},'value':{'font_size':20,'bold':True,'bg_color':'F5F6F8','num_format':'#,##0;[Red](#,##0)'},'value_pct':{'font_size':20,'bold':True,'bg_color':'F5F6F8','num_format':'0.00%'},'value_x':{'font_size':20,'bold':True,'bg_color':'F5F6F8','num_format':'0.00"×"'},'head':{'bold':True,'font_color':'FFFFFF','bg_color':'252932','font_size':9,'text_wrap':True,'valign':'vcenter'},'text':{'font_size':10,'text_wrap':True,'valign':'vcenter'},'num':{'font_size':10,'num_format':'#,##0;[Red](#,##0)','bottom':1,'bottom_color':'E5E7EB'},'pct':{'font_size':10,'num_format':'0.00%'},'note':{'font_size':10,'text_wrap':True,'bg_color':'F5F6F8','valign':'vcenter'},'alert':{'font_size':11,'bold':True,'font_color':'B31F2B','text_wrap':True,'valign':'vcenter'}}
  self.fm={k:self.wb.add_format({'font_name':'Aptos','font_color':'202833',**v}) for k,v in styles.items()};self.lang="'V06 Controlos'!$B$5";self.page=0
 def read(self,sn,addr):
  if sn not in self.sources:self.sources[sn]={c.coordinate:c.value for row in self.source[sn] for c in row if hasattr(c,'coordinate')}
  return self.sources[sn].get(addr)
 def ref(self,sn,addr):return "'"+sn+"'!"+addr
 def p(self,key):return self.ref('V06 Controlos','B'+str(REF['param_rows'][key]))
 def m(self,key,month):return self.ref('V06 Motor',col(REF['main_columns'][key])+str(month+4))
 def total(self,key):c=col(REF['main_columns'][key]);return f"SUM('V06 Motor'!{c}5:{c}184)"
 def tr(self,pt,en):return 'IF('+self.lang+'="PT","'+pt.replace('"','""')+'","'+en.replace('"','""')+'")'
 def value(self,formula):
  # Only evaluate display aggregates of existing source cells to seed valid caches.
  import re
  direct=re.fullmatch(r"'([^']+)'!\$?([A-Z]+)\$?(\d+)",formula)
  if direct:return self.read(direct[1],direct[2]+direct[3])
  total=re.fullmatch(r"SUM\('([^']+)'!([A-Z]+)(\d+):([A-Z]+)(\d+)\)",formula)
  if total:
   sn,c,a,c2,b=total.groups();assert c==c2;return sum(self.read(sn,c+str(r)) or 0 for r in range(int(a),int(b)+1))
  count=re.fullmatch(r"COUNTA\('([^']+)'!([A-Z]+)(\d+):([A-Z]+)(\d+)\)",formula)
  if count:return sum(self.read(count[1],count[2]+str(r)) is not None for r in range(int(count[3]),int(count[5])+1))
  return None
 def merged(self,row,c1,c2,formula,fmt='text',cached=None,height=1,financial=False):
  self.ws.merge_range(row,c1,row+height-1,c2,'',self.fm[fmt]);v=self.value(formula) if cached is None else cached
  self.ws.write_formula(row,c1,'='+formula,self.fm[fmt],v if v is not None else '')
  addr=col(c1)+str(row+1);self.cells[addr]={'formula':formula,'financial':financial,'cached':v}
  if financial:self.bindings.append({'cell':addr,'formula':formula,'base_value':v})
 def txt(self,row,pt,en,fmt='text',height=2):self.merged(self.offset+row,0,11,self.tr(pt,en),fmt,pt,height)
 def start(self,i):
  self.page=i;self.offset=(i-1)*50;o=self.offset
  self.ws.merge_range(o,0,o+1,2,'JSCV · KK-01',self.fm['brand']);self.ws.merge_range(o,3,o+1,11,'V06 · EXECUTIVE FINANCIAL CONTROL CENTER',self.fm['sub'])
  title=self.tr('10 — Relatório Executivo Consolidado de Viabilidade e Investimento','10 — Consolidated Investment & Feasibility Report');self.merged(o+3,0,11,title,'title','10 — Relatório Executivo Consolidado de Viabilidade e Investimento',3)
  self.merged(o+7,0,11,self.tr(f'{i:02} — '+SECTIONS[i-1][0],f'{i:02} — '+SECTIONS[i-1][1]),'section',f'{i:02} — '+SECTIONS[i-1][0],2)
  self.ws.merge_range(o+10,0,o+10,2,'',self.fm['sub']);self.ws.merge_range(o+10,4,o+10,7,'',self.fm['sub']);self.ws.merge_range(o+10,9,o+10,11,'',self.fm['sub']);self.ws.write_url(o+10,0,"internal:'EXEC V06'!A1",self.fm['sub'],string='← Executive');self.ws.write_url(o+10,4,"internal:'V06 Controlos'!B5",self.fm['sub'],string='PT / EN · Controlos / Controls');self.ws.write_url(o+10,9,"internal:'V06 Dicionario'!A1",self.fm['sub'],string='KPI ⓘ')
  self.txt(46,'Fonte: dashboards V06 e motor central · EUR · 180 meses · relatório dinâmico.','Source: V06 dashboards and central engine · EUR · 180 months · dynamic report.','sub')
  self.txt(48,'Candidato de estudo · Teste nativo Microsoft Excel 365 pendente.','Study candidate · Native Microsoft Excel 365 acceptance pending.','sub',1)
 def cards(self,items):
  for i,(pt,en,f,fmt) in enumerate(items):
   r=self.offset+13+(i//2)*5;c=(i%2)*6
   self.merged(r,c,c+5,self.tr(pt,en),'label',pt,2);self.merged(r+2,c,c+5,f,fmt,financial=True,height=2)
 def table(self,row,headers,records,widths=None,formats=None):
  widths=widths or [12//len(headers)]*len(headers);formats=formats or ['text']+['num']*(len(headers)-1);offsets=[];c=0
  for width,(pt,en) in zip(widths,headers):offsets.append(c);self.merged(self.offset+row,c,c+width-1,self.tr(pt,en),'head',pt,2);c+=width
  for j,record in enumerate(records):
   for k,f in enumerate(record):
    ptcache=None
    if isinstance(f,tuple):f,ptcache=f
    self.merged(self.offset+row+2+j*2,offsets[k],offsets[k]+widths[k]-1,f,formats[k],cached=ptcache,height=2,financial=formats[k] in ['num','pct'])
 def chart(self,row,specs,source='V06 Anual',first=5,last=19):
  ch=self.wb.add_chart({'type':'line'})
  for c,pt,en in specs:
   ch.add_series({'name':pt+' / '+en,'categories':f"='{source}'!$A${first}:$A${last}",'values':f"='{source}'!${c}${first}:${c}${last}",'values_data':[self.read(source,c+str(r)) or 0 for r in range(first,last+1)],'categories_data':[self.read(source,'A'+str(r)) for r in range(first,last+1)],'line':{'width':1.6}})
  ch.set_size({'width':620,'height':205});ch.set_chartarea({'border':{'none':True}});ch.set_y_axis({'num_format':'0,," M€"'});ch.set_legend({'position':'bottom','font':{'size':9}});self.ws.insert_chart(self.offset+row,0,ch)
 def generate(self):
  D=lambda sn,c:self.ref(sn,c);S=lambda c:D('V06 Cenarios',c+'5');P=self.p
  self.start(1);self.cards([('VGV','Development sales value',D('DASH 04A V06','I9'),'value'),('CAPEX futuro','Future CAPEX',D('DASH 01 V06','I9'),'value'),('TIR novos investidores','New investor IRR',D('DASH 07 V06','E15'),'value_pct'),('MOIC novos investidores','New investor MOIC',D('DASH 07 V06','I15'),'value_x'),('VAL investidores','Investor NPV',D('DASH 07 V06','M15'),'value'),('Equity Value de saída','Exit Equity Value',D('DASH 07 V06','M9'),'value')])
  alert='IF('+S('G')+'>0,'+self.tr('ALERTA: existe défice de financiamento. Os retornos são condicionais.','WARNING: an unfunded financing gap exists. Returns are conditional.')+',IF('+S('M')+'<0,'+self.tr('VAL negativo: o retorno não alcança a taxa de desconto selecionada.','Negative NPV: returns do not meet the selected discount rate.')+','+self.tr('VAL não negativo no cenário ativo; confirmar contratos, capital e risco.','Non-negative NPV in the active case; confirm contracts, capital and risk.')+'))'
  self.merged(30,0,11,alert,'alert','VAL negativo: o retorno não alcança a taxa de desconto selecionada.',3)
  self.table(35,[('Hotel','Hotel'),('Desconto equity','Equity discount'),('Saída','Exit multiple')],[[P('hotel_scenario'),P('discount'),P('exit_multiple')]],formats=['text','pct','num'])
  self.txt(42,'Dividendos, incentivos e saída são simulações; o relatório não certifica viabilidade nem benefícios concedidos.','Dividends, incentives and exit values are simulations; this report does not certify feasibility or granted benefits.','note',3)
  self.start(2);self.cards([('Investimento histórico, sem caixa futura','Historical investment, no future cash outflow',D('DASH 01 V06','E9'),'value'),('Orçamento futuro base','Future development budget',D('DASH 01 V06','I9'),'value'),('Capital alvo','Equity target',D('DASH 06 V06','M9'),'value'),('Capital KK de cenário','Scenario KK capital',D('DASH 06 V06','E15'),'value')])
  self.txt(23,'Programa: 200 quartos · MICE até 1.500 participantes · 94 frações.','Programme: 200 rooms · MICE up to 1,500 participants · 94 apartments.','sub',2)
  self.table(25,[('Construção / FF&E inicial','Construction / initial FF&E'),('Fiscalização','Supervision'),('Contingência','Contingency')],[[D('DASH 01 V06','M9'),D('DASH 01 V06','E15'),D('DASH 01 V06','I15')]],formats=['num','num','num'])
  self.chart(31,[('B','CAPEX trimestral','Quarterly CAPEX'),('C','Capital KK','KK capital'),('D','Banco','Bank')],'V06 Trimestral',5,14)
  self.start(3);self.cards([('Frações','Apartments',D('DASH 04A V06','E9'),'value'),('Recebimentos 180M','180M collections',D('DASH 04A V06','M9'),'value'),('Custos comerciais','Sales costs',D('DASH 04A V06','E15'),'value'),('Margem promoção — proxy','Development margin — proxy',D('DASH 04A V06','M15'),'value')])
  self.chart(25,[('E','Recebimentos trimestrais','Quarterly collections')],'V06 Trimestral',5,14)
  self.txt(41,'94 frações: inventário único em V06 Unidades. Previsões não equivalem a contratos assinados ou recebimentos comprovados. As frações 504 e 810 conservam o M29 herdado, sujeito a validação.','94 apartments: single inventory in V06 Unidades. Forecasts are not signed contracts or verified receipts. Units 504 and 810 retain inherited M29, subject to validation.','note',4)
  self.start(4);self.cards([('ADR de abertura','Opening ADR',D('DASH 05 V06','E9'),'value'),('Ocupação quartos — abertura','Room occupancy — opening',D('DASH 05 V06','I9'),'value_pct'),('Receita MICE — 180 meses','MICE revenue — 180 months',self.total('mice_revenue'),'value'),('EBITDA Ano15','Year15 EBITDA',D('DASH 05 V06','E15'),'value')])
  self.table(25,[('Fee base / receita','Base fee / revenue'),('Sistemas / receita','Systems / revenue'),('Incentivo / GOP ajustado','Incentive / adjusted GOP')],[[P('base_fee'),P('systems'),P('incentive')]],formats=['pct','pct','pct'])
  self.chart(31,[('B','Receita Hotel','Hotel revenue'),('C','EBITDA','EBITDA')])
  self.start(5);self.cards([('Receita residencial 15A','Residential revenue 15Y',D('DASH 04B V06','E9'),'value'),('Rendas suportáveis proprietários','Affordable owner distributions',D('DASH 04B V06','I9'),'value'),('Yield alvo não coberto','Uncovered owner yield target',D('DASH 04B V06','M9'),'value'),('Residual KK','KK residual cash',D('DASH 04B V06','E15'),'value'),('Dotações FF&E residencial','Residential FF&E allocations',D('DASH 04B V06','I15'),'value'),('Fundo adicional residencial','Additional residential fund',D('DASH 04B V06','M15'),'value')])
  self.txt(31,'O modelo separa a exploração do pool da venda das frações. A renda paga é limitada à caixa disponível; o yield é objetivo, não garantia.','The model separates pool operations from apartment development. Owner payments are limited to available cash; target yield is not guaranteed.','note',4)
  self.table(37,[('Yield inicial','Initial yield'),('Incremento anual','Annual increase'),('Imposto proprietários — hipótese','Owner tax — assumption')],[[P('res_yield'),P('res_yield_step'),P('owner_tax')]],formats=['pct','pct','pct'])
  self.start(6);self.cards([('IVA suportado','VAT incurred',D('DASH 02 V06','E9'),'value'),('Pico crédito IVA','Peak VAT credit',D('DASH 02 V06','I9'),'value'),('IRPC pago','Corporate tax paid',D('DASH 02 V06','M9'),'value'),('Seguro obra adicional','Additional construction insurance',D('DASH 02 V06','E15'),'value')])
  self.table(25,[('IRPC de estudo','Study corporate tax'),('IVA recuperável','VAT recoverable fraction'),('Prazo IVA, meses','VAT lag, months')],[[P('tax_rate'),P('vat_recovery'),P('vat_lag')]],formats=['pct','pct','num'])
  self.table(32,[('Estado da Convenção','Convention status'),('Isenção simulada, anos','Simulated exemption, years')],[[P('convention'),P('exempt_years')]],widths=[8,4],formats=['text','num'])
  self.txt(39,'Convenção de Estabelecimento Turístico: incentivos pretendidos sujeitos a negociação, aprovação e formalização. O estado da candidatura não concede benefícios. IPI, direitos aduaneiros e selo a zero representam exposições por validar, não isenções confirmadas.','Tourism Investment Convention: requested incentives require negotiation, approval and formalisation. Application status does not grant benefits. Zero property tax, customs and stamp inputs are unresolved exposures, not confirmed exemptions.','note',6)
  self.start(7);self.cards([('Caixa M30','Cash M30',D('DASH 03 V06','E9'),'value'),('Caixa M42','Cash M42',D('DASH 03 V06','I9'),'value'),('Caixa M54','Cash M54',D('DASH 03 V06','M9'),'value'),('Pico dívida de fecho','Peak closing debt',S('F'),'value')])
  self.merged(self.offset+23,0,7,self.tr('FF&E Hotel — dotações 180M','Hotel FF&E — 180M allocations'),'text','FF&E Hotel — dotações 180M');self.merged(self.offset+23,8,11,self.total('ffe'),'num',financial=True)
  self.table(25,[('Reserva HOTUSA ativa','Active HOTUSA reserve'),('Défice máximo','Maximum funding gap'),('Falta de reserva','Reserve shortfall')],[[self.m('hotusa_reserve',54),S('G'),S('H')]],formats=['num','num','num'])
  self.chart(31,[('H','Caixa fecho','Closing cash'),('I','Dívida fecho','Closing debt')])
  self.start(8);self.cards([('Blocos totais','Ownership blocks',D('DASH 06 V06','E9'),'value'),('Tickets novos','New investor tickets',D('DASH 06 V06','I9'),'value'),('Distribuição M54','M54 distribution',D('DASH 07 V06','E9'),'value'),('Dividendos até saída','Dividends through exit',D('DASH 07 V06','I9'),'value')])
  self.table(25,[('JSCV','JSCV'),('JALS','JALS'),('Novos investidores','New investors')],[[D('DASH 06 V06','I15'),D('DASH 06 V06','M15'),"SUM('V06 Blocos'!D12:D19)"]],formats=['pct','pct','pct'])
  self.table(32,[('Política dividendos','Dividend policy'),('Cobertura 12M','12M coverage'),('Lucros: modo','Profit mode')],[[P('div_policy'),P('coverage'),P('profit_mode')]],formats=['text','num','text'])
  self.txt(40,'Onze datas elegíveis, M54–M174. Cobertura dos 12 meses anteriores e liquidez acumulada são testes distintos. Reservas e lucros condicionam distribuições. Calendários de capital são hipóteses, não entradas comprovadas.','Eleven eligible dates, M54–M174. Prior-12-month coverage and cumulative liquidity are separate tests. Reserves and profit constraints govern distributions. Capital schedules are assumptions, not verified receipts.','note',5)
  self.start(9);self.cards([('VAL projeto','Project NPV',S('T'),'value'),('VAL novos investidores','New investor NPV',S('M'),'value'),('TIR projeto','Project IRR',S('U'),'value_pct'),('TIR novos investidores','New investor IRR',S('K'),'value_pct'),('MOIC novos investidores','New investor MOIC',S('L'),'value_x'),('Equity Value de saída','Exit Equity Value',S('J'),'value')])
  # EV above must follow the selected horizon; source summary N is Year15 only.
  # Use the already calculated scenario exit cash-flow (equity) and EV formula in its dependency graph.
  # The source can be M120: replace EV card with the active horizon control, no duplicate valuation.
  self.txt(31,'A saída selecionada é alternativa: M120 ou M180, nunca ambas. Dívida, custos, obrigações e caixa elegível são tratados no motor. Apartamentos vendidos não são somados novamente.','The selected exit is an alternative: M120 or M180, never both. Debt, costs, obligations and eligible cash are handled by the engine. Sold apartments are not added again.','note',4)
  self.table(37,[('Mês saída','Exit month'),('Múltiplo','Exit multiple'),('Desconto equity','Equity discount')],[[P('exit_month'),P('exit_multiple'),P('discount')]],formats=['num','num','pct'])
  self.txt(43,'Custos e datas históricos individuais dos promotores não disponíveis: não se inventa TIR histórica.','Individual promoter historical costs and dates are unavailable: historical IRR is not invented.','sub',2)
  self.start(10)
  records=[]
  labels=[(5,'Base','Base'),(7,'IVA diferido','VAT delayed'),(9,'IRPC 3 anos','IRPC 3 years'),(13,'Seguros +50%','Insurance +50%'),(23,'Stress combinado','Combined stress')]
  # Source rows are looked up by scenario name, preserving exact current ordering.
  for _,pt,en in labels:
   lookup={'Base':'Base','IVA diferido':'VAT delayed','IRPC 3 anos':'IRPC 3 years','Seguros +50%':'Insurance +50%','Stress combinado':'Combined stress'}[pt]
   sr=next(r for r in range(5,30) if self.read('V06 Cenarios','A'+str(r))==lookup)
   records.append([(self.tr(pt,en),pt),D('V06 Cenarios','K'+str(sr)),D('V06 Cenarios','M'+str(sr)),D('V06 Cenarios','G'+str(sr))])
  self.table(13,[('Cenário','Scenario'),('TIR','IRR'),('VAL investidores','Investor NPV'),('Défice máximo','Maximum gap')],records,widths=[4,2,3,3],formats=['text','pct','num','num'])
  self.txt(28,'Condições para investir: capital e vendas documentados; orçamento e banco confirmados; fiscalidade validada; reservas protegidas; retorno compatível com a taxa exigida.','Investment conditions: documented capital and sales; confirmed budget and financing; validated tax; protected reserves; returns consistent with the required rate.','note',5)
  self.txt(35,'O relatório reflete o cenário ativo e não otimiza pressupostos para tornar o VAL positivo. Não substitui parecer fiscal, jurídico, bancário ou avaliação independente.','The report reflects the active scenario and does not optimise assumptions to force a positive NPV. It does not replace tax, legal, banking or independent valuation advice.','note',4)
  self.txt(40,'Limites atuais: eventos MICE sem calendário operacional documentado; FF&E sem razão autónoma de gastos e saldo protegido V06; compatibilidade Excel 365 e paridade exata V05.1 não certificadas.','Current limitations: MICE lacks a documented event schedule; FF&E lacks a standalone V06 expenditure/protected-balance ledger; Excel 365 compatibility and exact V05.1 parity are not certified.','sub',5)
  self.wb.close();self.source.close();manifest={'base':str(self.base),'base_sha256':hashlib.sha256(self.base.read_bytes()).hexdigest(),'sections':[{'page':i+1,'pt':a,'en':b} for i,(a,b) in enumerate(SECTIONS)],'financial_bindings':self.bindings,'report_cells':self.cells,'pages':10,'paper':'A4 portrait','engine_changes':0};(ROOT/'Fontes_DASH10.json').write_text(json.dumps(manifest,ensure_ascii=False,indent=2));print('DASH10 module generated',len(self.bindings),'financial references')
if __name__=='__main__':Report(Path(sys.argv[1]) if len(sys.argv)>1 else BASE).generate()
