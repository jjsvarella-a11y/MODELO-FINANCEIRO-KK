"""Replace demonstrably generic charts with approved purpose-specific charts.
Hotel manually adjusted charts and the four banking charts are not regenerated.
All data are linked to existing monthly engines; support cells only aggregate.
"""
from finalize_v08 import *
import xlsxwriter
import tempfile

def chart_refine(source,target):
 src=Source(source);changes={};state=Sheet(E.fromstring(src.z.read(src.pp[STATUS7])),{});assets=Path(tempfile.mkdtemp(prefix='kk08_charts_'))/'Chart_Assets.xlsx';book=xlsxwriter.Workbook(assets);ws=book.add_worksheet(STATUS7);nextrow=2100;charts=[];proof=[];monthly=[{k:src.get('V06 Motor',col(v)+str(r))[1] for k,v in CC.items()} for r in range(5,185)]
 def chart(sn,index,title_pt,title_en,labels,series,ctype='column',percent=False):
  nonlocal nextrow
  start=nextrow;nextrow+=max(len(labels)+3,20);cols=[]
  def put(r,c,v,f=None):
   addr=col(c)+str(r);state.put(addr,v,f);state.row(r).set('hidden','1')
   if f:ws.write_formula(r-1,c,'='+f,None,v)
   else:ws.write(r-1,c,v)
  for j,item in enumerate(labels):
   v,f=item if isinstance(item,tuple) else (item,None);put(start+j,0,v,f)
  for c,(pt,en,items) in enumerate(series,1):
   for j,(f,v) in enumerate(items):put(start+j,c,v,f)
   put(start-1,c,pt,tr(pt,en));cols.append(c)
  ch=book.add_chart({'type':ctype});colors=['142337','1765A8','C8101E','9BBB59']
  for c,(_,_,items) in zip(cols,series):
   opts={'name':[STATUS7,start-2,c],'categories':[STATUS7,start-1,0,start+len(labels)-2,0],'values':[STATUS7,start-1,c,start+len(labels)-2,c],'line':{'color':'#'+colors[(c-1)%4]},'fill':{'color':'#'+colors[(c-1)%4]}}
   ch.add_series(opts)
  # Title is a linked bilingual cell rather than a cached English-only string.
  put(start-2,0,title_pt,tr(title_pt,title_en));ch.set_title({'name':[STATUS7,start-3,0],'name_font':{'name':'Aptos','size':11,'bold':True}});ch.set_chartarea({'border':{'none':True},'fill':{'color':'FFFFFF'}});ch.set_plotarea({'border':{'none':True}});ch.set_legend({'position':'bottom','font':{'name':'Aptos','size':9}});ch.set_x_axis({'num_font':{'name':'Aptos','size':8},'major_tick_mark':'none'});ch.set_y_axis({'num_font':{'name':'Aptos','size':8},'num_format':'0.0%' if percent else '#,##0.0','major_gridlines':{'visible':True,'line':{'color':'E9EFF5'}}});ws.insert_chart(0,len(charts)*9,ch);charts.append((sn,index));proof.append({'sheet':sn,'chart_index':index,'purpose_pt':title_pt,'purpose_en':title_en,'support_rows':[start,len(labels)],'financial_source':'Existing engine outputs; display aggregation only'})
 def agg(key,period=12,sn='V06 Motor',divide=1e6):
  vals=[];cc=col(CC[key])
  for st in range(0,180,period):
   lo=st+5;hi=min(st+period,180)+4;f='SUM('+q(sn,cc+str(lo)+':'+cc+str(hi))+')/'+str(divide)
   v=sum((src.get(sn,cc+str(r))[1] or 0) for r in range(lo,hi+1))/divide;vals.append((f,v))
  return vals
 years=[(i,tr('Ano '+str(i),'Year '+str(i))) for i in range(1,16)];quarters=[(i,tr('T'+str(i),'Q'+str(i))) for i in range(1,61)]
 def single(key):return ('SUM('+q('V06 Motor',col(CC[key])+'5:'+col(CC[key])+'184')+')/1000000',sum((r[key] or 0) for r in monthly)/1e6)
 chart('EXEC V06',1,'Fontes de financiamento e CAPEX · M€','Funding sources and CAPEX · EUR m',[(pt,tr(pt,en)) for pt,en in [('Capital previsto','Planned equity'),('Vendas previstas','Forecast collections'),('Utilização bancária','Bank drawdowns'),('CAPEX futuro','Future CAPEX')]], [('Montante','Amount',[single(k) for k in ['equity','receipts','draw','capex']])])
 # Construction: approved budget, quarterly execution/VAT and cumulative S curve.
 labels=[src.get(CON,'A'+str(r))[1] for r in range(12,42)];series=[]
 for mode,title in [('Base','Base'),('Comparative','Comparativo')]:
  sn=MAPS[mode][CON];series.append((title,'Comparative' if mode=='Comparative' else 'Base',[(q(sn,'G'+str(r))+'/1000000',src.get(sn,'G'+str(r))[1]/1e6) for r in range(12,42)]))
 chart('DASH 01 V06',0,'Orçamento por capítulo · M€','Budget by chapter · EUR m',labels,series)
 chart('DASH 01 V06',1,'Desembolso trimestral · CAPEX e IVA','Quarterly disbursement · CAPEX and VAT',quarters,[('CAPEX sem IVA','CAPEX net of VAT',agg('capex',3)),('IVA suportado','Input VAT',agg('vat_incurred',3))])
 ss=[]
 for mode in ['Base','Comparative']:
  sn=MAPS[mode]['V06 Motor'];ss.append((mode,mode, [('SUM('+q(sn,'$L$5:L'+str(7+3*i))+')/1000000',sum(src.get(sn,'L'+str(r))[1] for r in range(5,8+3*i))/1e6) for i in range(60)]))
 chart('DASH 01 V06',2,'Curva S · CAPEX acumulado','S curve · cumulative CAPEX',quarters,ss,'line')
 chart('DASH 02 V06',1,'IRPC anual · Base e Comparativo','Annual corporate tax · Base and Comparative',years,[(m,m,agg('tax_paid',12,MAPS[m]['V06 Motor'])) for m in ['Base','Comparative']])
 chart('DASH 02 V06',2,'Seguros e folga de liquidez · M€','Insurance and liquidity headroom · EUR m',years,[('Seguros de obra','Construction insurance',agg('construction_insurance')),('Seguros de exploração','Operating insurance',agg('operating_insurance')),('Caixa menos reservas','Cash less reserves',[(q('V06 Motor','DG'+str(16+12*i))+'-'+q('V06 Motor','BV'+str(16+12*i)),monthly[11+12*i]['cash']-monthly[11+12*i]['reserve']) for i in range(15)])],'line')
 # Correct the scale of the cash line to the same EUR million unit as insurance.
 for c in range(2100,nextrow):
  addr='D'+str(c);node=state.cells.get(addr)
  if node is not None:
   f=node.find(T('f'));v=node.find(T('v'))
   if f is not None and f.text and "'V06 Motor'!DG" in f.text:
    f.text='('+f.text+')/1000000';v.text=str(float(v.text)/1e6);ws.write_formula(c-1,3,'='+f.text,None,float(v.text))
 types=['T1','T2','T3','Penthouse'];un='V06 Unidades'
 # Use actual source typology names; do not invent an apartment classification.
 types=list(dict.fromkeys(str(src.get(un,'B'+str(r))[1]) for r in range(5,99)))
 data=[];price=[]
 for typ in types:
  rs=[r for r in range(5,99) if str(src.get(un,'B'+str(r))[1])==typ];lit='"'+typ+'"';query=q('DASH 04A V06','H35');data.append(('COUNTIFS('+q(un,'B5:B98')+','+lit+','+q(un,'L5:L98')+',">"&'+query+')',sum(1 for r in rs if src.get(un,'L'+str(r))[1]>(src.get('DASH 04A V06','H35')[1] or 0))))
  price.append(('SUMIF('+q(un,'B5:B98')+','+lit+','+q(un,'H5:H98')+')/SUMIF('+q(un,'B5:B98')+','+lit+','+q(un,'D5:D98')+')',sum(src.get(un,'H'+str(r))[1] for r in rs)/sum(src.get(un,'D'+str(r))[1] for r in rs)))
 chart('DASH 04A V06',0,'Disponibilidade prevista por tipologia','Forecast availability by typology',types,[('Disponíveis no mês consultado','Available at selected month',data)])
 chart('DASH 04A V06',1,'Pricing por tipologia · €/m²','Pricing by typology · EUR/m2',types,[('Preço médio ponderado','Weighted average price',price)])
 chart('DASH 04A V06',2,'Recebimentos trimestrais · M€','Quarterly collections · EUR m',quarters,[('Recebimentos previstos','Forecast collections',agg('receipts',3))])
 # Residential owners remain separate from shareholders.
 total=sum(src.get(un,'H'+str(r))[1] for r in range(5,99));yieldrows=[(f.rsplit('/',1)[0]+'/SUM(\'V06 Unidades\'!H5:H98)',v/total) for f,v in agg('owner_distribution',divide=1)]
 chart('DASH 04B V06',0,'Yield anual dos proprietários','Annual owner cash yield',years,[('Distribuição / investimento','Distribution / investment',yieldrows)],'line',True)
 chart('DASH 04B V06',1,'Receita e OPEX residencial · M€','Residential revenue and OPEX · EUR m',years,[('Receita','Revenue',agg('res_revenue')),('OPEX','OPEX',agg('res_opex'))])
 chart('DASH 04B V06',2,'Distribuições e FF&E residencial · M€','Owner distributions and residential FF&E · EUR m',years,[('Distribuições proprietários','Owner distributions',agg('owner_distribution')),('Reserva FF&E','FF&E allocations',agg('res_ffe'))])
 chart('DASH 06 V06',1,'Captação prevista e capital documentado · M€','Planned equity and documented paid capital · EUR m',[(pt,tr(pt,en)) for pt,en in [('Previsto','Planned'),('Documentado','Documented')]], [('Capital','Capital', [('SUM(\'V06 Capital\'!C5:C12)/1000000',sum(src.get('V06 Capital','C'+str(r))[1] for r in range(5,13))/1e6),('SUM(\'V06 Capital\'!H5:H12)/1000000',sum(src.get('V06 Capital','H'+str(r))[1] or 0 for r in range(5,13))/1e6)])])
 chart('DASH 06 V06',2,'Blocos novos por classe','New investment blocks by class',['Platinum','Gold','Silver'], [('Blocos','Blocks',[('COUNTIF(\'V06 Capital\'!B5:B12,"'+t+'")',sum(src.get('V06 Capital','B'+str(r))[1]==t for r in range(5,13))) for t in ['Platinum','Gold','Silver']])])
 opening=monthly[0]['open'];operating_years=(180-int(opening)+12)//12;oplabels=[(i,tr('Expl. '+str(i),'Op. '+str(i))) for i in range(1,operating_years+1)]
 def op_sum(key,i):
  cc=col(CC[key]);f='SUMIFS('+q('V06 Motor',cc+'5:'+cc+'184')+','+q('V06 Motor','E5:E184')+','+str(i-1)+','+q('V06 Motor','A5:A184')+',">="&'+q('V06 Motor','D5')+')/1000000';return f,sum(r[key] for r in monthly if r['op_age']==i-1 and r['month']>=opening)/1e6
 chart('DASH 07 V06',0,'Dividendos por ano de exploração · M€','Dividends by operating year · EUR m',oplabels,[('Dividendos','Dividends',[op_sum('dividend',i) for i in range(1,operating_years+1)])])
 chart('DASH 07 V06',1,'Dividendos e valor de saída · M€','Dividends and exit value · EUR m',[(pt,tr(pt,en)) for pt,en in [('Dividendos','Dividends'),('Equity na saída','Exit equity')]], [('Projeto KK','KK project',[(q('V06 Cenarios','O5')+'/1000000',src.get('V06 Cenarios','O5')[1]/1e6),(q('V06 Cenarios','J5')+'/1000000',src.get('V06 Cenarios','J5')[1]/1e6)])])
 book.close()
 with ZipFile(assets) as az:
  # Native chart caches are generated from formula caches written above.
  for i,(sn,index) in enumerate(charts,1):
   r=E.fromstring(src.z.read(src.pp[sn]));drawing=r.find(T('drawing'));rels=E.fromstring(src.z.read(relpath(src.pp[sn])));rid=drawing.get('{'+NSR+'}id');dp=resolve(src.pp[sn],next(x.get('Target') for x in rels if x.get('Id')==rid));dr=E.fromstring(changes.get(dp,src.z.read(dp)));anchor=list(dr)[index];chartnode=next(n for n in anchor.iter() if E.QName(n).localname=='chart');cr=E.fromstring(src.z.read(relpath(dp)));cp=resolve(dp,next(x.get('Target') for x in cr if x.get('Id')==chartnode.get('{'+NSR+'}id')));changes[cp]=az.read('xl/charts/chart'+str(i)+'.xml')
   if sn=='EXEC V06':
    dn='{http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing}'
    # Funding first, liquidity/debt second, global performance third.
    for idx,col0,col1 in [(0,11,18),(1,4,11)]:
     n=list(dr)[idx];n.find(dn+'from').find(dn+'col').text=str(col0);n.find(dn+'to').find(dn+'col').text=str(col1)
    changes[dp]=xml(dr)
 dimension=state.r.find(T('dimension'));dimension.set('ref','A1:BP'+str(max(2000,nextrow)));changes[src.pp[STATUS7]]=state.finish();rewrite(source,target,changes);src.z.close();return proof
if __name__=='__main__':
 p=argparse.ArgumentParser();p.add_argument('source',type=Path);p.add_argument('target',type=Path);p.add_argument('--proof',type=Path,required=True);a=p.parse_args();a.proof.write_text(json.dumps(chart_refine(a.source,a.target),ensure_ascii=False,indent=2))
