"""Four executive report formats. Current SAVED Base governs every default report.
Full recalc + reconciliation before issue; reproducible immutable JSON and PDF.
"""
from scenario_v07 import *
import argparse,os,copy
os.environ.setdefault('MPLCONFIGDIR','/tmp/kk_matplotlib');os.environ.setdefault('XDG_CACHE_HOME','/tmp/kk_cache')
from reportlab.platypus import SimpleDocTemplate,Paragraph,Spacer,Table,TableStyle,PageBreak,Image,KeepTogether
from reportlab.lib.styles import getSampleStyleSheet,ParagraphStyle
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from xml.sax.saxutils import escape
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
FONT='/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf'
pdfmetrics.registerFont(TTFont('KK',FONT));pdfmetrics.registerFont(TTFont('KKB','/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf'))
TYPES=['Executive One Page','Bank Financing','Equity Investor','Comprehensive Project']
REPORTDIR=OUT7/'Relatorios';REPORTDIR.mkdir(exist_ok=True)
def number(v,kind='money'):
 if not isinstance(v,(int,float)):return str(v)
 if kind=='pct':return f'{v*100:.2f}%'
 if kind=='ratio':return f'{v:.2f}x'
 return f'{v:,.0f} EUR'
def capture_tables(workbook,context='Base'):
 source=Source(workbook);matrix=values(source,'V06 Matriz Valores');rows=[(0,.5),(2,.6),(3,.65),(4,.7),(5,.75)];sn=MAPS[context][CON];un=MAPS[context]['V06 Unidades'];out={'matrix_rows':rows,'matrix':[[matrix['C'+str(5+row*5+j)] for j in range(4)] for row,occ in rows],'break_even':[source.get('DASH 05 V06',a)[1] for a in ['E135','J135','O135','T135']],'budget':[[source.get(sn,c+str(r))[1] for c in ['A','B','G','H']] for r in range(12,42)],'units':[[source.get(un,c+str(r))[1] for c in ['A','B','D','H','O']] for r in range(5,99)],'near_zero_threshold':source.get('DASH 05 V06','W116')[1]};source.z.close();return out

def pdf_report(file,data,context,comparisons,kind,lang,settings,identifier):
 en=lang=='EN';say=lambda pt,eng:eng if en else pt;ctx=data['contexts'][context];m=ctx['metrics'];params=ctx['params'];inputs=ctx['effective_inputs'];monthly=ctx['monthly'];styles=getSampleStyleSheet()
 for name,size,lead in [('KKBody',9,13),('KKSmall',9,12),('KKTitle',19,24),('KKSection',12,17)]:styles.add(ParagraphStyle(name=name,fontName='KKB' if name in ['KKTitle','KKSection'] else 'KK',fontSize=size,leading=lead,textColor=colors.HexColor('#142337'),spaceAfter=6))
 tables=data.get('report_tables') or capture_tables(data['workbook'],context)
 flow=[];width=515
 def number(value,kind='money'):
  out=globals()['number'](value,kind)
  return out if en else out.replace(',', '\u202f').replace('.', ',')
 names={key:(pt,eng) for key,pt,eng,*_ in PARAMS};names['effective_tan']=('TAN ativa','Active TAN')
 opnames={d['key']:(d['label_pt'],d['label_en']) for d in DEFS}
 status=say('Hipóteses','Hypothetical') if params['status']=='Hypothetical' else say('Proposta bancária','Bank proposal') if params['status']=='Bank proposal' else say('Aprovadas','Approved')
 evidence=say('Não disponibilizada','Not supplied') if params['evidence']=='Not supplied' else params['evidence']
 heat_limit=settings.get('near_zero_threshold',500000)
 def text(s,style='KKBody'):return Paragraph(escape(str(s)),styles[style])
 def para(s):flow.append(text(s))
 def section(pt,eng):flow.extend([Spacer(1,9),text(say(pt,eng),'KKSection')])
 def table(headers,rows,small=False,heat=None):
  items=[[text(x,'KKSmall') for x in headers]]+[[text(x,'KKSmall') for x in row] for row in rows];t=Table(items,colWidths=[width/len(headers)]*len(headers),repeatRows=1,hAlign='LEFT');t.setStyle(TableStyle([('BACKGROUND',(0,0),(-1,0),colors.HexColor('#E9EFF5')),('VALIGN',(0,0),(-1,-1),'TOP'),('TOPPADDING',(0,0),(-1,-1),4),('BOTTOMPADDING',(0,0),(-1,-1),4),('LINEBELOW',(0,0),(-1,-1),.25,colors.HexColor('#DCE3EB'))]));
  if heat is not None:
   for row,entries in enumerate(heat,1):
    for column,value in enumerate(entries,1):
     fill='#FFF0C2' if abs(value)<=heat_limit else '#FBE5E7' if value<0 else '#E0F1E5';t.setStyle(TableStyle([('BACKGROUND',(column,row),(column,row),colors.HexColor(fill))]))
  flow.extend([t,Spacer(1,7)])
 titles={'Executive One Page':('Síntese Executiva','Executive One-Page Report'),'Bank Financing':('Relatório de Financiamento Bancário','Bank Financing Report'),'Equity Investor':('Relatório dos Investidores','Equity Investor Report'),'Comprehensive Project':('Relatório Integral do Projeto','Comprehensive Project Report')}
 flow.append(text('KK-01 | HOTEL & RESIDENCES','KKTitle'));flow.append(text(say(*titles[kind]),'KKSection'));para('KK Hotelaria e Turismo, S.A. | Praia, Cabo Verde');para(settings['recipient']+' | '+settings['confidentiality']+' | '+identifier)
 para(say('Base atual guardada: ','Current saved Base: ')+str(ctx['id'])+' | '+data['calculated_at']+' | V07')
 summary=[(say('CAPEX futuro','Future CAPEX'),'Future CAPEX','money'),(say('Pico da dívida','Peak bank debt'),'Peak Bank Debt','money'),(say('Folga mínima de liquidez','Minimum liquidity headroom'),'Minimum Liquidity Headroom','money'),(say('Dividendos previstos','Planned dividends'),'Planned Dividends','money'),(say('VAL novos investidores','New investor NPV'),'Investor NPV','money'),(say('TIR novos investidores','New investor IRR'),'Investor IRR','pct'),('MOIC','MOIC','ratio'),(say('Equity na saída','Exit equity value'),'Exit Equity Value','money')]
 heads=[say('Indicador','Indicator'),say('Base atual','Current Base')]+[x['id'] for x in comparisons]
 table(heads,[[label,number(m[key],nk)]+[number(x['metrics'][key],nk)+' | Δ '+number(x['metrics'][key]-m[key],nk) for x in comparisons] for label,key,nk in summary])
 para(say('Hipóteses ativas: ','Active assumptions: ')+f'TAN {number(params["effective_tan"],"pct")}; '+say('carência ','grace ')+f'{params["grace"]:.0f}m; '+say('prazo após carência ','tenor following grace ')+f'{params["term"]:.0f}m; ADR {inputs["adr"]:.2f} EUR; '+say('ocupação inicial ','opening occupancy ')+number(inputs['occ'],'pct')+'; '+say('desconto investidores ','investor discount ')+number(inputs['discount'],'pct'))
 conclusion=say('O VAL dos novos investidores é negativo às hipóteses e taxa de desconto selecionadas; estes resultados não sustentam uma conclusão favorável automática sobre bancabilidade.','New investor NPV is negative at the selected assumptions and discount rate; these results do not support an automatic favourable bankability conclusion.') if m['Investor NPV']<0 else say('O VAL é positivo no cenário modelado; não constitui aprovação bancária nem garantia de rendibilidade.','NPV is positive in the modelled scenario; this is not bank approval or a return guarantee.')
 para(conclusion)
 para(say('Custo efetivo anual do financiamento: ','Effective annual financing cost: ')+number(ctx['bank_metrics']['Effective annual financing cost'],'pct'))
 para(say('Condições bancárias: ','Bank terms: ')+status+' | '+str(evidence))
 if kind=='Executive One Page':
  para(say('Fonte: motor de 180 meses recalculado e reconciliado. Reservas e limites preservados. Validação estrutural e recálculo independente não substituem teste nativo no Excel 365.','Source: recalculated and reconciled 180-month engine. Reserves and limits retained. Structural validation and independent calculation do not replace native Excel 365 testing.'))
 else:
  flow.append(PageBreak());section('Pressupostos e condições ativas','Active assumptions and conditions')
  keys=['tan','effective_tan','facility','grace','term','sweep','opening','opening_fixed','broker','availability','commitment','commit_base','management','period','prepay','closing','interest_grace','draw_start','draw_end','rate_type','index','spread','convention','dscr_min','include_fees','status','evidence']
  table([say('Parâmetro','Parameter'),say('Valor','Value')],[[say(*names[k]),number(params[k],'pct' if k in ['tan','effective_tan','sweep','opening','broker','availability','commitment','prepay','index','spread'] else 'ratio' if k=='dscr_min' else 'money' if k in ['facility','opening_fixed','management','closing'] else 'text') if isinstance(params[k],(int,float)) and k not in ['grace','term','period','draw_start','draw_end','interest_grace','include_fees'] else str(params[k])] for k in keys])
  para(say('O prazo de amortização segue a carência de capital. A carência de juros é independente. Juros incidem sobre dívida efetivamente utilizada; as comissões seguem as bases contratuais ou hipóteses identificadas.','Amortisation tenor follows principal grace. Interest grace is separate. Interest applies to actual utilised debt; fees follow the documented contractual basis or identified assumptions.'))
  if comparisons:
   section('Comparação das condições financeiras','Financing terms comparison')
   compare_rows=[]
   for key in ['effective_tan','facility','term','grace','opening','commitment','prepay']:
    nk='pct' if key in ['effective_tan','opening','commitment','prepay'] else 'money' if key=='facility' else 'ratio';compare_rows.append([key,number(params[key],nk)]+[number(x['params'][key],nk) for x in comparisons])
   for key in ['Minimum DSCR','Total financing cost','Effective annual financing cost','Final debt repayment month']:
    nk='pct' if key=='Effective annual financing cost' else 'money' if key=='Total financing cost' else 'ratio';compare_rows.append([key,number(ctx['bank_metrics'][key],nk)]+[number(x['bank_metrics'][key],nk) for x in comparisons])
   table(heads,compare_rows)
  section('Operação e remuneração','Operations and operator remuneration')
  table([say('Parâmetro','Parameter'),say('Valor','Value')],[[say(*opnames[k]),number(inputs[k],'money' if k=='adr' else 'pct')] for k in ['adr','occ','growth','mice','basefee','incentive','systems','ffe','discount']])
  para(say('Ocupação inicial mantém ramp-up de 1,2 pontos percentuais por ano até ao teto do motor. ADR cresce por idade operacional. A comissão de incentivo incide no GOP líquido das restantes comissões e prioridade aplicável, conforme motor. FF&E é reserva, distinta da manutenção.','Opening occupancy preserves the engine ramp-up of 1.2 percentage points per year to its cap. ADR grows by operating age. Incentive fees apply to GOP net of the other fees and applicable priority under the existing engine. FF&E is a reserve, separate from maintenance.'))
  section('Evolução de receitas e EBITDA','Revenue and EBITDA evolution')
  annual=[]
  for y in range(15):
   rows=monthly[y*12:y*12+12];annual.append({'year':y+1,**{k:sum(r[k] for r in rows) for k in ['revenue','ebitda','cfads','draw','principal','sweep','interest','all_fees','contractual_service','dividend','capex','vat_incurred','vat_refund','receipts']},'debt':rows[-1]['debt'],'headroom':min(r['headroom'] for r in rows)})
  chart=WORK7/(identifier+'_'+lang+'.png');fig,ax=plt.subplots(figsize=(7.2,2.5));ax.bar([a['year'] for a in annual],[a['revenue']/1e6 for a in annual],color='#183B65',label=say('Receita','Revenue'));ax.plot([a['year'] for a in annual],[a['ebitda']/1e6 for a in annual],color='#C8101E',label='EBITDA');ax.set_ylabel('M EUR');ax.set_xlabel(say('Ano do projeto','Project year'));ax.legend(frameon=False);fig.tight_layout();fig.savefig(chart,dpi=150);plt.close(fig);flow.append(Image(str(chart),width=width,height=180))
  table([say('Ano','Year'),say('Receitas','Revenue'),'EBITDA',say('Dividendos','Dividends')],[[r['year'],number(r['revenue']),number(r['ebitda']),number(r['dividend'])] for r in annual])
  if kind in ['Bank Financing','Comprehensive Project']:
   section('Fontes de reembolso e serviço da dívida','Repayment sources and debt service')
   para(say('As fontes de reembolso modeladas são recebimentos imobiliários, cash flow operacional e capital previsto. Capital e vendas previstos não são comprovativos de recebimento. Cash sweep pode liquidar a dívida antes do fim da carência; prazo e carência isolados podem assim não alterar resultados.','Modelled repayment sources are real estate collections, operating cash flow and planned equity. Forecast equity and sales are not proof of actual receipts. Cash sweep may settle debt before grace ends; changing tenor or grace alone can therefore legitimately leave results unchanged.'))
   table([say('Ano','Year'),'CFADS',say('Serviço contratual','Contractual service'),say('Juros','Interest'),say('Comissões','Fees'),say('Dívida final','Closing debt')],[[r['year']]+[number(r[k]) for k in ['cfads','contractual_service','interest','all_fees','debt']] for r in annual])
   services=[r['cfads']/r['contractual_service'] for r in monthly if r['contractual_service']>.005];section('Cobertura, liquidez e restrições','Coverage, liquidity and constraints')
   para('DSCR min: '+(f'{min(services):.2f}x' if services else 'N/A')+' | '+say('Défice máximo: ','Maximum funding gap: ')+number(m['Funding Gap'])+' | '+say('Custo financiamento: ','Financing cost: ')+number(m['Financing Cost']))
   para(say('Serviço contratual e pagamentos bancários totais são distintos. DSCR usa o CFADS e denominador configurados no motor; o custo total inclui juros devidos e todas as comissões.','Contractual service and total bank cash payments are distinct. DSCR uses the engine CFADS and configured denominator; total financing cost includes accrued interest and all fees.'))
   critical=sorted(monthly,key=lambda r:r['headroom'])[:12];table([say('Mês','Month'),say('Caixa','Cash'),say('Reservas','Reserves'),say('Folga','Headroom')],[[r['month'],number(r['cash']),number(r['reserve']),number(r['headroom'])] for r in critical])
   section('Construção, desembolsos e IVA','Construction, disbursements and VAT')
   table([say('Trimestre','Quarter'),'CAPEX',say('IVA incorrido','VAT incurred'),say('Bruto','Gross'),say('IVA recuperado','VAT recovered'),say('Utilizações','Drawdowns')],[[j//3+1]+[number(sum(x[k] for x in monthly[j:j+3])) for k in ['capex','vat_incurred']]+[number(sum(x['capex']+x['vat_incurred'] for x in monthly[j:j+3]))]+[number(sum(x[k] for x in monthly[j:j+3])) for k in ['vat_refund','draw']] for j in range(0,180,3) if any(x['capex'] or x['vat_refund'] for x in monthly[j:j+3])])
  if kind in ['Equity Investor','Comprehensive Project']:
   section('Capital, distribuição e valor terminal','Equity, distributions and terminal value')
   para(say('A estrutura preservada tem 15 blocos: JSCV 40%, JALS 10%, e novos investidores 50% / 25 M EUR. Direitos adquiridos não são apresentados como novo capital subscrito. TIR e VAL dos promotores exigem fluxos históricos completos, não disponíveis.','The retained structure has 15 blocks: JSCV 40%, JALS 10%, and new investors 50% / EUR 25m. Acquired rights are not presented as newly subscribed equity. Promoter IRR and NPV require complete historical cash flows, which are not available.'))
   table([say('Ano','Year'),say('Fluxo novos investidores','New investor cash flow')],[[y+1,number(sum(ctx['investor_cashflows'][1+y*12:1+(y+1)*12]))] for y in range(15)])
  section('Sensibilidade real ADR × Ocupação','Actual ADR × Occupancy sensitivity')
  matrixrows=tables['matrix_rows'];heat=tables['matrix'];table(['Occup. / ADR','150 EUR','175 EUR','200 EUR','225 EUR'],[[f'{occ:.0%}']+[number(value) for value in row] for (_,occ),row in zip(matrixrows,heat)],heat=heat);breakeven=tables['break_even'];para(say('Escala: rosa = negativo; dourado = próximo de zero; verde = positivo. Limiar: ','Scale: pink = negative; gold = near zero; green = positive. Threshold: ')+number(heat_limit));table([say('Equilíbrio na grelha','Break-even within tested grid'),say('Resultado','Result')],[[say('ADR mínimo à ocupação ativa','Minimum ADR at active occupancy'),number(breakeven[0]) if isinstance(breakeven[0],(int,float)) else say('Não atingido na grelha','Not reached within tested grid')],[say('Ocupação mínima ao ADR ativo','Minimum occupancy at active ADR'),number(breakeven[1],'pct') if isinstance(breakeven[1],(int,float)) else say('Não atingido na grelha','Not reached within tested grid')],[say('VAL atual dos investidores','Current investor NPV'),number(breakeven[2])],[say('Défice de VAL para zero','NPV gap to break-even'),number(breakeven[3])]])
  para(say('Cada ponto usa o motor mensal com os restantes pressupostos ativos; a matriz acompanha o contexto ativo identificado no Excel. Se o contexto ativo for Comparativo, esta matriz representa esse Comparativo e não a Base. As cores e o limiar são os definidos no Excel.','Each point uses the monthly engine with the other active assumptions; the matrix follows the active context identified in Excel. If Comparative is active, this matrix represents that Comparative rather than Base. Colours and threshold follow Excel.'))
  section('Riscos, dados em falta e metodologia','Risks, missing evidence and methodology')
  para(conclusion);para(say('Condições não documentadas permanecem hipóteses. Não foi disponibilizado o ficheiro binário do logótipo oficial nem medição certificada da área bruta exclusiva do Aparthotel; 47 000 m² é área global/proxy de imputação. Não se inventam contratos nem pagamentos efetivos.','Undocumented conditions remain assumptions. The official logo binary and certified Aparthotel-only gross area were not supplied; 47,000 m² is a global area/allocation proxy. No contracts or actual receipts are invented.'))
  para(say('Recálculo integral independente: LibreOffice. Verificação estrutural: Open XML SDK. Execução nativa Microsoft Excel 365 Desktop indisponível neste ambiente; não se afirma aprovação nativa.','Independent full recalculation: LibreOffice. Structural verification: Open XML SDK. Native Microsoft Excel 365 Desktop execution is unavailable in this environment; native acceptance is not claimed.'))
  if settings['appendices']=='Yes':
   if kind=='Comprehensive Project':
    section('Orçamento integral — 30 capítulos','Complete budget — 30 chapters');table([say('Código','Code'),say('Capítulo','Chapter'),say('Orçamento','Budget'),say('ETC','ETC')],[[row[0],row[1],number(row[2]),number(row[3])] for row in tables['budget']])
    section('Frações — registo de 94 unidades','Apartments — 94-unit register');table([say('Fração','Unit'),say('Tipo','Type'),say('Área venda','Saleable area'),say('Preço ativo','Active price'),say('Estado documental','Documented status')],[[row[0],row[1],row[2],number(row[3]),row[4]] for row in tables['units']])
   if kind in ['Bank Financing','Comprehensive Project']:
    section('Cronograma mensal completo — 180 meses','Complete monthly schedule — 180 months');table([say('Mês','Month'),say('Dívida inicial','Opening debt'),say('Utilização','Draw'),say('Capital + sweep','Principal + sweep'),say('Juros','Interest'),say('Dívida final','Closing debt')],[[r['month'],number(r['debt_open']),number(r['draw']),number(r['principal']+r['sweep']),number(r['interest']),number(r['debt'])] for r in monthly])
 def footer(canvas,doc):
  canvas.setFont('KK',9);canvas.setFillColor(colors.HexColor('#687586'));canvas.drawString(40,22,'KK-01 | JS-CV | '+settings['confidentiality']);canvas.drawRightString(A4[0]-40,22,str(doc.page));canvas.setStrokeColor(colors.HexColor('#C8101E'));canvas.line(40,34,A4[0]-40,34)
 SimpleDocTemplate(str(file),pagesize=A4,rightMargin=40,leftMargin=40,topMargin=30,bottomMargin=44).build(flow,onFirstPage=footer,onLaterPages=footer)

def issue(file=FINAL7,types=None,languages=None,data=None):
 file=Path(file);data=data if data is not None and data.get('_validated_file_sha256')==sha(file) else full_recalc(file,WORK7/('Issue_'+datetime.now().strftime('%H%M%S%f')))
 if not data['contexts']['Base']['saved']:raise ValueError('BLOCKED: unsaved Reference Base; save-base and recalculate first')
 source=Source(file)
 if source.get('EXEC V06','H23')[1]=='Comparative':
  basecopy=WORK7/'Report_Base_Context.xlsx';source.z.close();shutil.copy2(file,basecopy);patch_cells(basecopy,[('EXEC V06','H23','Base')]);baserecalculated=full_recalc(basecopy,WORK7/('Report_Base_'+datetime.now().strftime('%H%M%S%f')));data['workbook']=str(basecopy);source=Source(file)
 settings={'type':source.get(REPORT7,'B6')[1],'comparison':source.get(REPORT7,'B7')[1],'language':source.get(REPORT7,'B8')[1],'recipient':source.get(REPORT7,'B9')[1],'ids':source.get(REPORT7,'B10')[1] or '', 'confidentiality':source.get(REPORT7,'B11')[1],'appendices':source.get(REPORT7,'B12')[1],'near_zero_threshold':source.get('DASH 05 V06','W116')[1]};data.setdefault('calculated_at',datetime.now(timezone.utc).isoformat());data.setdefault('workbook',str(file));comparisons=[]
 if settings['comparison']!='Base Only':
  if settings['comparison']=='Base vs Selected':
   if not data['contexts']['Comparative']['saved']:raise ValueError('Save selected comparative before issue')
   comparisons=[data['contexts']['Comparative']]
  else:
   identifiers=[x.strip() for x in settings['ids'].split(',') if x.strip()]
   if not identifiers:raise ValueError('Select saved comparative IDs')
   for identifier in identifiers:
    snapshot=json.loads((VERSIONS/(identifier+'.json')).read_text(encoding='utf-8-sig'))
    if snapshot['classification']!='Comparative':raise ValueError('Not a comparative scenario: '+identifier)
    temp=WORK7/(identifier+'.xlsx');shutil.copy2(file,temp);edits=[]
    for key,raw in snapshot['raw_inputs'].items():
     sn,addr=selection_location(source,'Comparative',key)
     if key not in ALIASES:edits.append((sn,addr,raw['selection']))
     edits.append((CMP7,'C'+str(ROWS[key]),raw['custom']))
    patch_cells(temp,edits);d=full_recalc(temp,WORK7/('Comparison_'+identifier));c=d['contexts']['Comparative'];c['id']=identifier;comparisons.append(c)
 data['report_tables']=capture_tables(data['workbook'],'Base')
 types=types or (TYPES if settings['type']=='All Formats' else [settings['type']]);languages=languages or (['PT','EN'] if settings['language']=='Both' else [settings['language']]);outputs=[];edits=[];last=max([int(re.sub('[A-Z]','',a)) for a in source.cells(REPORT7) if a.startswith('A')]+[24]);source.z.close()
 for kind in types:
  for lang in languages:
   identifier='KK01-'+datetime.now(timezone.utc).strftime('%Y%m%dT%H%M%S%fZ')+'-'+kind.replace(' ','_')+'-'+lang;pdf=REPORTDIR/(identifier+'.pdf');pdf_report(pdf,data,'Base',comparisons,kind,lang,settings,identifier)
   if kind=='Executive One Page':
    import fitz
    with fitz.open(pdf) as doc:
     if len(doc)!=1:pdf.unlink();raise ValueError('Executive report exceeds one page')
   snapshot={'report_tables':data['report_tables'],'id':identifier,'base_id':data['contexts']['Base']['id'],'comparative_ids':[c['id'] for c in comparisons],'settings':{**settings,'type':kind,'language':lang},'format':kind,'language':lang,'calculated_at':data['calculated_at'],'source_sha256':data['source_sha256'],'pdf_sha256':sha(pdf),'base':data['contexts']['Base'],'comparisons':comparisons,'calculation_platform':data['calculation_platform'],'native_excel365':'NOT TESTED'};jsonfile=pdf.with_suffix('.json');jsonfile.write_text(json.dumps(snapshot,ensure_ascii=False,indent=2));outputs.append(str(pdf));last+=1;row=[identifier,data['calculated_at'],kind,lang,settings['recipient'],snapshot['base_id'],','.join(snapshot['comparative_ids']),data['contexts']['Base']['params']['effective_tan'],data['contexts']['Base']['effective_inputs']['discount'],'Reconciled; native not tested',sha(pdf),sha(jsonfile)];edits +=[(REPORT7,col(j)+str(last),v) for j,v in enumerate(row)]
 patch_cells(file,edits);return outputs
if __name__=='__main__':
 p=argparse.ArgumentParser();p.add_argument('--file',type=Path,default=FINAL7);p.add_argument('--all',action='store_true');p.add_argument('--both-languages',action='store_true');a=p.parse_args();print('\n'.join(issue(a.file,TYPES if a.all else None,['PT','EN'] if a.both_languages else None)))
