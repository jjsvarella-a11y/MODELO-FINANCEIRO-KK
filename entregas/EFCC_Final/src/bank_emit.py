"""Controlled bank-report emitter; never exports cached or unreconciled figures.
Full engine recalculation for selected and comparison scenarios. Emits immutable
JSON/180-month CSV + source/settings hashes + snapshot Excel + A4 bank PDFs.
LibreOffice is an independent calculation engine, NOT native Excel acceptance.
"""
from bank_r6 import *
from test_closure_r5 import mutate,recalculate,load,near
from bank_read import read_bank
from assemble import inject_caches
import argparse,csv,shutil,math,io
from reportlab.platypus import SimpleDocTemplate,Paragraph,Spacer,Table,TableStyle,PageBreak,Image,KeepTogether
from reportlab.lib.styles import getSampleStyleSheet,ParagraphStyle
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from xml.sax.saxutils import escape

def sha(path):return hashlib.sha256(Path(path).read_bytes()).hexdigest()
def evaluate(file):
 c,errors,stats=read_bank(file)
 if errors:raise ValueError('Formula errors: '+repr(errors[:12]))
 m=c['V06 Motor'];b=c[DAT];params={key:b['B'+str(P[key])] for key,*_ in PARAMS};params['effective_tan']=b['B42'];params['effective_interest_annual']=b['B43']
 essential=['facility','tan','grace','term','draw_start','draw_end','period','dscr_min']
 for key in essential:
  if not isinstance(params[key],(int,float)) or not math.isfinite(params[key]):raise ValueError('Missing essential assumption '+key)
 if params['term']<=0 or params['period']<=0 or params['dscr_min']<=0 or params['draw_start']>params['draw_end'] or params['draw_end']>180:raise ValueError('Invalid essential financing assumptions')
 if params['status']=='Approved' and params['evidence'] in ['Not supplied','Not saved','']:raise ValueError('Approved terms lack documentary evidence')
 rows=[];checks={'cash_max_error':0.,'debt_max_error':0.,'matrix_errors':len(errors)}
 for rr in range(5,185):
  r={key:m[col(k)+str(rr)] for key,k in CC.items()};r.update(interest_due=m['EK'+str(rr)],capitalized_interest=m['EL'+str(rr)],early_fee=m['EM'+str(rr)],closing_fee=m['EN'+str(rr)],all_fees=m['EO'+str(rr)],contractual_service=m['EP'+str(rr)],bank_payments=m['EQ'+str(rr)])
  for key in ['cash_check','debt_check','ffe_fund_check','inventory_check']:near(r[key],0,.005)
  near(r['debt_open']+r['draw']+r['capitalized_interest']-r['principal']-r['sweep'],r['debt'],.005)
  if r['debt_open']+r['draw']<.005:near(r['interest_due'],0,.005)
  if r['cumulative_draw']>params['facility']+.005:raise ValueError('Facility exceeded')
  if r['sweep']>r['debt_open']+r['draw']+r['capitalized_interest']-r['principal']+.005:raise ValueError('Early repayment exceeds outstanding principal')
  near(r['all_fees'],r['commitment_fee']+r['bank_fees']+r['early_fee']+r['closing_fee'],.005)
  near(r['bank_payments'],r['interest']+r['principal']+r['sweep']+r['all_fees'],.005)
  if r['contractual_service']>0:near(r['dscr'],r['cfads']/r['contractual_service'])
  r['dscr_display']=r['cfads']/r['contractual_service'] if r['contractual_service']>0 else 'N/A'
  checks['cash_max_error']=max(checks['cash_max_error'],abs(r['cash_check']));checks['debt_max_error']=max(checks['debt_max_error'],abs(r['debt_check']));rows.append(r)
 discount=c['V06 Registo Ativo']['B65'];npv=sum(c['V06 F01'].get('B'+str(i+5),0)/(1+discount)**(i/12) for i in range(181));near(npv,c['V06 Cenarios']['M5'])
 annual=[]
 for y in range(15):
  block=rows[y*12:(y+1)*12];annual.append({'year':y+1,**{key:sum(r[key] for r in block) for key in ['draw','principal','sweep','interest','interest_due','all_fees','contractual_service','bank_payments','cfads','capex','receipts','equity','dividend']},'debt':block[-1]['debt'],'min_liquidity':min(r['cash']-r['reserve'] for r in block)})
 metrics={c[SCH]['A'+str(rr)]:c[SCH]['B'+str(rr)] for rr in range(210,227)}
 # A cost IRR on an unpaid loan is not a valid full-life financing cost.
 if rows[-1]['debt']>.005:metrics['Effective annual financing cost']='N/A - debt outstanding beyond 180 months'
 return {'params':params,'metrics':metrics,'monthly':rows,'annual':annual,'npv':npv,'irr':c['V06 Cenarios']['K5'],'moic':c['V06 Cenarios']['L5'],'discount_rate':discount,'checks':checks,'errors':0,'stats':stats,'active_profile':c[BN]['H33'],'global_mode':c['EXEC V06']['H23']}

def compute(source,scenario,folder,recipient=None,overrides=None,need_snapshot=True):
 folder=Path(folder);folder.mkdir(parents=True,exist_ok=True);inputfile=folder/'Input.xlsx';changes=[('EXEC V06','H23','Test'),(BN,'H33',scenario)]
 cfg=Source(source)
 if cfg.get('EXEC V06','H23')[1]=='Reference':
  # Financing alternatives inherit the currently effective non-bank R4 case,
  # never silently revive dormant Test assumptions from another scenario.
  changes += [(i['sheet'],i['cell'],i['value']) for i in INPUTS if i['sheet']!=BN]
  changes += [('DASH 07 V06','H29','Custom'),('DASH 07 V06','N25',cfg.get('V06 Registo Ativo','B65')[1])]
 if recipient:changes.append((CENTER,'H122',recipient))
 if overrides:changes+=overrides
 mutate(source,inputfile,changes);calculated=recalculate(inputfile,folder/'recalc');data=evaluate(calculated)
 snapshot=folder/'Snapshot.xlsx' if need_snapshot else calculated
 if need_snapshot:inject_caches(inputfile,calculated,snapshot)
 data['source_sha256']=sha(source);data['snapshot_sha256']=sha(snapshot);data['calculation_platform']='LibreOffice full engine recalculation';data['calculated_at']=datetime.now(timezone.utc).isoformat();data['native_excel365']='NOT TESTED'
 return snapshot,data

def display(x,kind='money'):
 if not isinstance(x,(int,float)):return str(x)
 if kind=='pct':return f'{x*100:,.2f}%'
 if kind=='ratio':return f'{x:,.2f}x'
 if kind=='whole':return f'{x:,.0f}'
 return f'{x:,.0f} EUR'

def bank_pdf(data,base,path,recipient,language='PT',brief=False,selected_only=False,stresses=None):
 en=language=='EN';say=lambda pt,eng:eng if en else pt;styles=getSampleStyleSheet();styles.add(ParagraphStyle(name='SmallKK',fontName='Helvetica',fontSize=8,leading=11,textColor=colors.HexColor('#142337')));styles.add(ParagraphStyle(name='TitleKK',fontName='Helvetica-Bold',fontSize=18,leading=22,textColor=colors.HexColor('#142337')));flow=[];M=data['metrics'];p=data['params']
 translations={'Maximum annual debt service':'Serviço anual máximo da dívida','Minimum DSCR':'DSCR mínimo','Peak outstanding debt':'Pico de dívida','Final debt repayment month':'Mês de liquidação da dívida','Total interest paid':'Total de juros pagos','Total financing fees':'Total de comissões e encargos','Total financing cost':'Custo total do financiamento','Effective annual financing cost':'Custo efetivo anual do financiamento','Funding gap':'Necessidade adicional de financiamento','Minimum liquidity headroom':'Folga mínima de liquidez','Remaining non-revolving capacity':'Capacidade remanescente da linha não renovável','New investor NPV':'VAL dos novos investidores','New investor IRR':'TIR dos novos investidores','Exit equity value':'Valor dos capitais próprios na saída','Cumulative dividends':'Dividendos acumulados','Interest accrued / due':'Juros devidos','Year':'Ano','Principal':'Capital programado','Interest paid':'Juros pagos','Fees':'Comissões','Debt service':'Serviço da dívida','Month':'Mês','Cash':'Caixa','Reserves':'Reservas','Headroom':'Folga','Debt':'Dívida','Draw':'Utilizações','Repayment':'Amortizações','Value':'Valor','Parameter':'Parâmetro','Check':'Verificação','Result':'Resultado','Sales receipts':'Recebimentos imobiliários','Equity':'Capital próprio','Selected':'Selecionado','Test':'Teste'}
 def text(s,style='SmallKK'):
  s=translations.get(str(s),str(s)) if not en else str(s)
  return Paragraph(escape(str(s)),styles[style])
 def para(s):flow.extend([text(s),Spacer(1,7)])
 def title(i,pt,eng):
  if i>1:flow.append(PageBreak())
  flow.extend([text(f'{i:02} | '+say(pt,eng),'TitleKK'),Spacer(1,12)])
  para(say('Destinatário: ','Recipient: ')+recipient+' | '+data['active_profile']+' | '+data['id'])
 def table(headers,rows,widths=None):
  items=[[text(x) for x in headers]]+[[text(x) for x in row] for row in rows];t=Table(items,colWidths=widths or ([495/len(headers)]*len(headers)),repeatRows=1,hAlign='LEFT');t.setStyle(TableStyle([('BACKGROUND',(0,0),(-1,0),colors.HexColor('#E8EEF5')),('VALIGN',(0,0),(-1,-1),'TOP'),('BOTTOMPADDING',(0,0),(-1,-1),5),('TOPPADDING',(0,0),(-1,-1),5),('LINEBELOW',(0,0),(-1,0),.8,colors.HexColor('#142337')),('LINEBELOW',(0,1),(-1,-1),.25,colors.HexColor('#DCE3EB'))]));flow.extend([t,Spacer(1,10)])
 def comparison(keys):
  rows=[]
  for key in keys:
   kind='pct' if any(x in key for x in ['IRR','annual financing cost']) else 'ratio' if key in ['Minimum DSCR','MOIC'] else 'whole' if 'month' in key else 'money';x=M[key];v=base['metrics'][key];delta=x-v if isinstance(x,(int,float)) and isinstance(v,(int,float)) else 'N/A'
   rows.append([key,display(x,kind)] if selected_only else [key,display(v,kind),display(x,kind),display(delta,kind)])
  table([say('Indicador','Metric'),say('Selecionado','Selected')] if selected_only else [say('Indicador','Metric'),'Base',say('Teste','Test'),'Delta'],rows,[210,285] if selected_only else [190,105,105,95])
 def chart(keys,labels,title_,annual=True):
  fig,ax=plt.subplots(figsize=(7,2.4));rr=data['annual'] if annual else data['monthly'];xs=[r['year' if annual else 'month'] for r in rr]
  for j,(key,label) in enumerate(zip(keys,labels)):ax.plot(xs,[r[key]/1e6 for r in rr],label=label,color=['#1765A8','#C8101E','#142337'][j%3],linewidth=1.7)
  ax.set_title(title_,fontsize=10,loc='left');ax.set_ylabel('M EUR');ax.grid(alpha=.15);ax.legend(fontsize=8);fig.tight_layout();buf=io.BytesIO();fig.savefig(buf,format='png',dpi=145);plt.close(fig);buf.seek(0);flow.extend([Image(buf,width=495,height=170),Spacer(1,8)])
 def risk_text():
  lines=[]
  if data['npv']<0:lines.append(say('O VAL dos novos investidores é negativo à taxa de desconto ativa.','New-investor NPV is negative at the active discount rate.'))
  if M['Funding gap']>.005:lines.append(say('Existe necessidade adicional de financiamento.','Additional financing is required.'))
  if M['Minimum liquidity headroom']<-.005:lines.append(say('Há meses sem cobertura integral das reservas.','Some months do not fully cover required reserves.'))
  ds=M['Minimum DSCR']
  if isinstance(ds,(int,float)) and ds<p['dscr_min']:lines.append(say('O DSCR mínimo é inferior à meta interna configurada; não é um covenant aprovado.','Minimum DSCR is below the configured internal target, which is not an approved covenant.'))
  if data['monthly'][-1]['debt']>.005:lines.append(say('Existe dívida por liquidar após o horizonte de 180 meses.','Debt remains beyond the 180-month horizon.'))
  lines.append(say('Estes cálculos não constituem aprovação bancária nem demonstram, isoladamente, bancabilidade.','These calculations are not a bank approval and do not by themselves establish bankability.'))
  return ' '.join(lines)
 title(1,'Síntese de negociação bancária','Bank negotiation summary');para(say('KK-01 | Hotel & Residences — Praia, Cabo Verde. Base Oficial bancária: TAN 6% e carência 42 meses, alinhada com a R4 preservada. Condições sem aprovação documental.','KK-01 | Hotel & Residences — Praia, Cabo Verde. Official bank base: TAN 6% and grace 42 months, aligned with the preserved R4. Terms have no documentary bank approval.'))
 table(['TAN Base','TAN '+say('Teste','Test'),'Grace Base / Test','Facility'],[[display(base['params']['effective_tan'],'pct'),display(p['effective_tan'],'pct'),f"{base['params']['grace']:.0f} / {p['grace']:.0f} months",display(p['facility'])]])
 comparison(['Peak outstanding debt','Total financing cost','Maximum annual debt service','Minimum DSCR','Minimum liquidity headroom','New investor NPV','New investor IRR']);para(risk_text());para(say('Condições: ','Terms: ')+p['status']+'; '+say('documentação: ','documentation: ')+p['evidence']);para(say('Data UTC dos cálculos: ','Calculation UTC date: ')+data['calculated_at'])
 if brief:
  title(2,'Condições, liquidez e riscos','Terms, liquidity and risks');table(['TAN','Tenor / Grace','Facility','Fees'],[[display(p['effective_tan'],'pct'),f"{p['term']} / {p['grace']} months",display(p['facility']),display(M['Total financing fees'])]]);chart(['contractual_service','cfads'],['Debt service','CFADS'],'Coverage / repayment');para(risk_text());para(say('Consulte o relatório integral, pressupostos congelados e reconciliação de 180 meses incluídos na mesma emissão.','See the full report, frozen assumptions and 180-month reconciliation included with this issuance.'))
 else:
  title(2,'Condições bancárias e estado documental','Bank terms and documentary status')
  rows=[]
  keys=['facility','effective_tan','rate_type','index','spread','term','grace','interest_grace','draw_start','draw_end','amort','sweep','opening','opening_fixed','broker','availability','commitment','commit_base','management','period','prepay','closing','guarantee','expenses','insurance','convention','dscr_min','include_fees','status','evidence']
  for a1,a2 in zip(keys[::2],keys[1::2]):rows.append([a1,str(p[a1]),a2,str(p[a2])])
  table(['Parameter','Value','Parameter','Value'],rows,[128,119,128,120]);para(say('Prazo exclui carência de capital. Indexante variável é constante neste cenário, sem curva futura. Juros devidos sobre capital efetivamente utilizado, com capitalização durante carência de juros. Comissões sem documento são hipóteses.','Term excludes principal grace. A variable index is constant in this scenario, without a forward curve. Interest accrues on actual drawn debt, capitalised during interest grace. Undocumented fees are hypotheses.'))
  title(3,'Construção e fontes de reembolso','Construction and repayment sources');chart(['capex','receipts','equity'],['CAPEX','Property receipts','Equity'],'Construction / sales / equity');table(['Year','CAPEX','Sales receipts','Equity'],[[r['year'],display(r['capex']),display(r['receipts']),display(r['equity'])] for r in data['annual'][:10]]);para(say('O financiamento depende do calendário de entradas de capital, recebimentos imobiliários e caixa operacional. A linha não renovável não permite reutilizar amortizações. O atraso de obra não prorroga a janela bancária.','Financing depends on equity timing, property collections and operating cash. The non-revolving facility does not recycle repayments. Construction delays do not extend the bank availability period.'))
  title(4,'Desembolsos e fontes/aplicações','Drawdowns and sources/uses');chart(['draw','principal','sweep'],['Drawdowns','Scheduled principal','Early repayment'],'Annual bank flows');rr=data['monthly'];sources=sum(r['equity']+r['draw']+r['project_cf'] for r in rr);uses=sum(r['all_fees']+r['interest']+r['principal']+r['sweep']+r['dividend'] for r in rr);delta=rr[-1]['cash']-rr[0]['cash_open'];near(sources-uses,delta,.01);table(['Cash bridge','EUR'],[['Equity + draw + net project CF',display(sources)],['Fees + interest + repayments + dividends',display(uses)],['Closing less opening cash',display(delta)],['Reconciliation residual',display(sources-uses-delta)]]);para(say('Fluxo de projeto líquido inclui CAPEX, vendas, operação, fiscalidade e seguros conforme motor. Evita a dupla contabilização de receitas e custos.','Net project cash flow includes CAPEX, sales, operations, taxes and insurance as modelled. This avoids double counting receipts and costs.'))
  title(5,'Serviço da dívida e custo global','Debt service and financing cost');table(['Year','Principal','Interest paid','Fees','Debt service'],[[r['year'],display(r['principal']),display(r['interest']),display(r['all_fees']),display(r['contractual_service'])] for r in data['annual']]);comparison(['Total interest paid','Total financing fees','Total financing cost','Effective annual financing cost']);para(say('Serviço contratual exclui amortizações antecipadas; pagamentos totais incluem-nas. Juros efetivos anuais e custo efetivo do financiamento não substituem TAN. Custo efetivo é IRR dos fluxos de financiamento mensais, anualizada; N/A se existir saldo por liquidar no horizonte.','Contractual service excludes early principal; total bank payments include it. Effective annual interest and financing cost do not replace TAN. Effective cost is annualised IRR of monthly net financing flows; N/A if debt remains beyond the horizon.'))
  title(6,'DSCR e capacidade de reembolso','DSCR and repayment capacity');chart(['cfads','contractual_service'],['CFADS','Contractual debt service'],'CFADS vs debt service');comparison(['Minimum DSCR','Maximum annual debt service','Final debt repayment month']);table(['Year','CFADS','Service','Annual ratio'],[[r['year'],display(r['cfads']),display(r['contractual_service']),display(r['cfads']/r['contractual_service'],'ratio') if r['contractual_service']>0 else 'N/A'] for r in data['annual'][:10]]);para(say('DSCR segue CFADS do motor (antes de imposto e caixa residencial) / principal programado + juros pagos + compromisso; demais comissões apenas se incluídas no seletor. Meta interna editável: ','DSCR follows engine CFADS (before tax and residential cash) / scheduled principal + paid interest + commitment; other fees only if selected. Editable internal target: ')+display(p['dscr_min'],'ratio')+'.')
  title(7,'Liquidez e meses críticos','Liquidity and critical months');chart(['cash','reserve'],['Cash','Protected reserves'],'Cash and protected reserves',False);ordered=sorted(data['monthly'],key=lambda r:r['cash']-r['reserve'])[:8];table(['Month','Cash','Reserves','Headroom','Debt'],[[r['month'],display(r['cash']),display(r['reserve']),display(r['cash']-r['reserve']),display(r['debt'])] for r in ordered]);comparison(['Funding gap','Minimum liquidity headroom','Remaining non-revolving capacity']);para(risk_text())
  title(8,'Retornos e distribuições','Investor returns and distributions');chart(['dividend'],['Company dividends'],'Projected company dividends');comparison(['Cumulative dividends','New investor NPV','New investor IRR','MOIC','Exit equity value']);para(say('Taxa de desconto anual ativa: ','Active annual equity discount rate: ')+display(data['discount_rate'],'pct')+'. '+say('É independente da TAN bancária. VAL validado a partir dos 181 fluxos datados, incluindo investimento no mês zero. Direitos adquiridos dos promotores não são tratados como investimento histórico conhecido.','It is independent of the bank TAN. NPV is verified from 181 dated cash flows, including month-zero investment. Acquired promoter rights are not treated as documented historical outlays.'))
  title(9,'Stress e riscos','Stress tests and risks');para(risk_text())
  if stresses:table(['Stress','NPV','Peak debt','Min DSCR','Funding gap'],[[x['label'],display(x['npv']),display(x['metrics']['Peak outstanding debt']),display(x['metrics']['Minimum DSCR'],'ratio'),display(x['metrics']['Funding gap'])] for x in stresses])
  else:para(say('Não foram calculados stresses adicionais nesta emissão. Não se extrapolam conclusões; executar o emissor com --stress para incluir recálculos integrais.','No additional stresses were calculated for this issuance. No conclusions are extrapolated; run the emitter with --stress to include full recalculations.'))
  para(say('Riscos relevantes: receitas/ocupação, prazo da obra, vendas e recebimentos, recuperação de IVA, condições da operadora, reservas, custo bancário, refinanciamento e valor de saída. Condições de aprovação, garantias e covenants dependem de documentação bancária.','Relevant risks: revenue/occupancy, construction timing, sales/collections, VAT recovery, operator terms, reserves, bank costs, refinancing and exit value. Approval conditions, security and covenants depend on bank documentation.'))
  title(10,'Reconciliação e rastreabilidade','Reconciliation and traceability');table(['Check','Result'],[['Monthly engine','180 months, recalculated; all scenarios retained'],['Formula error scan','0 errors across all worksheets'],['Maximum cash residual',display(data['checks']['cash_max_error'])],['Maximum debt residual',display(data['checks']['debt_max_error'])],['Independent investor NPV',display(data['npv'])],['Version / scenario',data['id']+' / '+data['active_profile']],['Calculation reference date',data['calculated_at']],['Calculation platform',data['calculation_platform']],['Native Excel 365','NOT TESTED in this environment']]);para('Source SHA256: '+data['source_sha256']);para('Snapshot SHA256: '+data['snapshot_sha256']);para(say('Pressupostos e resultados integrais: Snapshot.json, Monthly_180.csv e Snapshot.xlsx. O perfil bancário oficial e a R4 são preservados; adotar novo perfil oficial exige ação expressa e nova versão. Este documento é um checkpoint calculado, sujeito a validação nativa e confirmação documental.','Complete assumptions/results: Snapshot.json, Monthly_180.csv and Snapshot.xlsx. The official bank profile and R4 remain preserved; adopting a new official profile requires an express action and a new version. This document is a calculated checkpoint, subject to native validation and documentary confirmation.'))
 def footer(canvas,doc):
  canvas.setFillColor(colors.HexColor('#C8101E'));canvas.rect(38,28,519,2,fill=1,stroke=0);canvas.setFont('Helvetica',7);canvas.setFillColor(colors.HexColor('#142337'));canvas.drawString(40,18,'KK-01 | JS-CV | '+data['id']);canvas.drawRightString(555,18,str(doc.page));canvas.setFont('Helvetica',8);canvas.drawString(40,817,'KK-01 EXECUTIVE FINANCIAL CONTROL CENTER')
 doc=SimpleDocTemplate(str(path),pagesize=A4,leftMargin=40,rightMargin=40,topMargin=42,bottomMargin=42,title='KK-01 Bank Negotiation '+data['id'],author='JS-CV');doc.build(flow,onFirstPage=footer,onLaterPages=footer)
 import fitz
 pdf=fitz.open(path);expected=2 if brief else 10
 if len(pdf)!=expected:raise ValueError(f'PDF page count {len(pdf)}, expected {expected}: {path}')
 for pg in pdf:
  txt=pg.get_text()
  if any(x in txt for x in ['#VALUE!','#REF!','#NAME?','###']):raise ValueError('PDF contains formula/rendering errors')
 return len(pdf)

def main():
 ap=argparse.ArgumentParser(description=__doc__);ap.add_argument('--source',type=Path,default=BM);ap.add_argument('--scenario');ap.add_argument('--comparison');ap.add_argument('--recipient');ap.add_argument('--output',type=Path);ap.add_argument('--selected-only',action='store_true');ap.add_argument('--format',choices=['auto','bank','brief','consolidated','all'],default='auto');ap.add_argument('--stress',action='store_true');args=ap.parse_args()
 cfg=Source(args.source)
 args.scenario=args.scenario or cfg.get(BN,'H33')[1]
 args.comparison=args.comparison or cfg.get(BN,'N33')[1]
 args.recipient=args.recipient or cfg.get(CENTER,'H122')[1]
 if not args.recipient or args.recipient.strip() in ['','To identify','Not supplied']:raise ValueError('Recipient must be identified')
 args.selected_only=args.selected_only or cfg.get(CENTER,'H106')[1]=='Selected only'
 if args.format=='auto':args.format={'Bank 10 pages':'bank','Executive 2 pages':'brief','Consolidated 10 pages':'consolidated'}.get(cfg.get(CENTER,'H102')[1],'bank')
 stamp=datetime.now(timezone.utc);ident=stamp.strftime('KK01-BANK-%Y%m%dT%H%M%SZ');dest=args.output or BOUT/'Emissoes'/ident
 if dest.exists():raise FileExistsError('Never overwrite an issued report: '+str(dest))
 dest.mkdir(parents=True);sourcehash=sha(args.source)
 try:
  work=BW/'emissions'/ident
  snap,data=compute(args.source,args.scenario,work/'selected',args.recipient);bsnap,base=compute(args.source,args.comparison,work/'comparison',args.recipient,need_snapshot=False)
  for dct in [data,base]:dct.update(id=ident,recipient=args.recipient)
  stress=[]
  if args.stress:
   # Explicit debt-independent stresses preserve chosen bank terms.
   cases=[('Construction +6m',[('DASH 01 V06','H23',6)]),('Operational Conservative',[('DASH 05 V06','H23','Conservative')]),('Sales collections +6m',[('DASH 04A V06','N23',6)])]
   for idx,(label,changes) in enumerate(cases):
    _,r=compute(args.source,args.scenario,work/f'stress_{idx+1}',overrides=changes,need_snapshot=False);r['label']=label;stress.append(r)
  from bank_snapshot import stamp as stamp_comparison
  data['comparison_guard']=stamp_comparison(snap,data,base);data['snapshot_sha256']=sha(snap)
  shutil.copy2(snap,dest/'Snapshot.xlsx');(dest/'Snapshot.json').write_text(json.dumps({'selected':data,'comparison':base,'stress':stress,'selection_only':args.selected_only},indent=2,ensure_ascii=False))
  with (dest/'Monthly_180.csv').open('w',newline='') as f:
   writer=csv.DictWriter(f,fieldnames=list(data['monthly'][0]));writer.writeheader();writer.writerows(data['monthly'])
  for lang in ['PT','EN']:
   if args.format in ['bank','all']:bank_pdf(data,base,dest/f'Bank_Report_10_{lang}.pdf',args.recipient,lang,selected_only=args.selected_only,stresses=stress)
   if args.format in ['brief','all']:bank_pdf(data,base,dest/f'Executive_2_{lang}.pdf',args.recipient,lang,brief=True,selected_only=args.selected_only)
   if args.format in ['consolidated','all']:
    import closure_presentation as cp
    from export_all import export
    cp.MASTER=snap;cp.WORK=work/('consolidated_'+lang);cp.WORK.mkdir(parents=True,exist_ok=True)
    presentation=cp.lab(lang);cp.chart_cache(presentation)
    pdf=export(presentation,['DASH 10 V06'],'Consolidated',work/('pdf_'+lang));shutil.copy2(pdf,dest/f'Consolidated_10_{lang}.pdf')
    import fitz
    doc=fitz.open(pdf);assert len(doc)==10
    assert all(not any(err in pg.get_text() for err in ['#VALUE!','#REF!','#NAME?','###']) for pg in doc)
  if sourcehash!=sha(args.source):raise ValueError('Source workbook changed during issuance')
  files={str(f.relative_to(dest)):sha(f) for f in dest.rglob('*') if f.is_file() and not any(x in str(f) for x in ['profile/','recalc/'])};(dest/'Manifest.json').write_text(json.dumps({'status':'CALCULATED CHECKPOINT - native Excel acceptance pending','id':ident,'files':files,'source_unchanged':True},indent=2));print('ISSUED',dest,flush=True)
 except Exception as ex:
  # No report may remain published after any validation/export failure.
  for f in dest.rglob('*.pdf'):f.unlink()
  (dest/'BLOCKED.json').write_text(json.dumps({'status':'BLOCKED','reason':repr(ex)},indent=2));raise
if __name__=='__main__':main()
