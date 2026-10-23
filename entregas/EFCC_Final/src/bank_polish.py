"""Final R6.1 banking presentation and project-wide financing comparison.
Existing Hotel graphic parts remain untouched. Cash projections are single-source.
"""
from bank_r6 import *
from zip_copy import rewrite

def apply():
 changes={};src=Source(BM)
 with ZipFile(BM) as z:
  wb,pp=paths(z);texts=[''.join(n.itertext()) for n in E.fromstring(z.read('xl/sharedStrings.xml'))];sty={};styles=E.fromstring(z.read('xl/styles.xml'))
  for c in E.fromstring(z.read(pp[ACT])).find(T('sheetData'))[0]:
   if c.get('t')=='s':sty[texts[int(c.findtext(T('v')))]]=int(c.get('s'))
  def sh(sn):return Sheet(E.fromstring(changes.get(pp[sn],z.read(pp[sn]))),sty)
  # Editable percentages must display as percentages, not raw fractions.
  s=sh(BN)
  for key,pt,en,rv,addr,kind in PARAMS:
   if kind=='pct' and int(addr[1:])>=450:s.cells[addr].set('s',str(sty['input_pct']))
  fm=styles.find(T('numFmts'));nf=max([163]+[int(n.get('numFmtId')) for n in fm])+1;E.SubElement(fm,T('numFmt'),numFmtId=str(nf),formatCode='yyyy-mm-dd');fm.set('count',str(len(fm)));xfs=styles.find(T('cellXfs'));datexf=copy.deepcopy(xfs[sty['input']]);datexf.set('numFmtId',str(nf));datexf.set('applyNumberFormat','1');dateid=len(xfs);xfs.append(datexf);xfs.set('count',str(len(xfs)));s.cells['N472'].set('s',str(dateid));s.text('K472:M473','Data indicativa M1 · hipótese','Indicative M1 date · assumption','label')
  for addr,area,pt,en in [('H23','E23:G23','Custom · TAN','Custom · TAN'),('N23','K23:M23','Custom · carência','Custom · grace'),('H25','E25:G25','Custom · prazo após carência','Custom · term after grace'),('N25','K25:M25','Custom · cash sweep','Custom · cash sweep'),('H27','E27:G27','Custom · última utilização','Custom · last draw'),('H29','E29:G29','Custom · amortização','Custom · amortisation'),('H31','E31:G31','Custom · linha €','Custom · facility EUR')]:s.text(area,pt,en,'label')
  active='IF('+reference()+',"Reference R4",H33)';s.merge('E36:P37','TAN: 6%','"TAN: "&TEXT('+q(DAT,'B42')+',"0.00%")&" | TEA: "&TEXT('+q(DAT,'B43')+',"0.00%")&" | "&'+active,'note')
  for rr,label,addr,kind in [(472,'TAN nominal',q(DAT,'B42'),'pct'),(474,'TEA juros',q(DAT,'B43'),'pct'),(476,'Custo efetivo anual',q(SCH,'B217'),'pct'),(478,'Comissões / encargos',q(SCH,'B215'),'money'),(480,'Custo total €',q(SCH,'B216'),'money')]:s.text(f'Q{rr}:U{rr+1}',label,{'TAN nominal':'Nominal TAN','TEA juros':'Effective annual interest','Custo efetivo anual':'Effective annual cost','Comissões / encargos':'Fees / charges','Custo total €':'Total cost EUR'}[label],'label');s.merge(f'V{rr}:Y{rr+1}',0,addr,kind)
  dx=styles.find(T('dxfs'));dxfid=len(dx);dxf=E.SubElement(dx,T('dxf'));font=E.SubElement(dxf,T('font'));E.SubElement(font,T('color'),rgb='FF8A949F');pat=E.SubElement(E.SubElement(dxf,T('fill')),T('patternFill'),patternType='solid');E.SubElement(pat,T('fgColor'),rgb='FFE9EDF1');E.SubElement(pat,T('bgColor'),rgb='FFE9EDF1');dx.set('count',str(len(dx)))
  dn=wb.find(T('definedNames'));E.SubElement(dn,T('definedName'),name='KK6_Bank_Custom').text='AND(\'EXEC V06\'!$H$23="Test",\'DASH 03 V06\'!$H$33="Custom")'
  areas=[]
  for _,_,_,_,addr,_ in PARAMS:
   rr=int(addr[1:]);areas.append(addr+':'+('J' if addr.startswith('H') else 'P')+str(rr+1 if rr>=450 else rr))
  cf=E.SubElement(s.r,T('conditionalFormatting'),sqref=' '.join(areas));rule=E.SubElement(cf,T('cfRule'),type='expression',dxfId=str(dxfid),priority='1');E.SubElement(rule,T('formula')).text='NOT(KK6_Bank_Custom)';changes[pp[BN]]=s.finish();changes['xl/styles.xml']=xml(styles)
  s=sh(DAT);old=s.cells['B43'];p='PRODUCT('+','.join('1+B42*DAY(EOMONTH('+d('start_date')+','+str(i)+'))/365' for i in range(12))+')-1';s.put('B43',src.get(DAT,'B43')[1],'IF('+d('convention')+'="ACT/365",'+p+',(1+B42/12)^12-1)',int(old.get('s')));changes[pp[DAT]]=s.finish()
  # Treat sub-cent projection noise as zero only in the presentation helper.
  s=sh(CASH)
  for rr in range(5,185):
   c=s.cells['O'+str(rr)];v=src.get(CASH,'O'+str(rr))[1];s.put('O'+str(rr),0 if abs(v)<.005 else v,'IF(ABS(B'+str(rr)+'-C'+str(rr)+')<0.005,0,B'+str(rr)+'-C'+str(rr)+')',int(c.get('s','0')))
  changes[pp[CASH]]=s.finish()
  # Project-wide bank impact uses immutable R4 outputs plus the active engine.
  refrows=[{k:sr.get('V06 Motor',col(v)+str(rr))[1] for k,v in CC.items()} for rr in range(5,185)];refcost=sum(r['interest']+r['bank_fees']+r['commitment_fee'] for r in refrows);refds=max(sum(r['debt_service'] for r in refrows[y*12:y*12+12]) for y in range(15));refdscr=min(r['dscr'] for r in refrows if r['debt_service']>0)
  reserve=[0 if abs(r['cash']-r['reserve'])<.005 else r['cash']-r['reserve'] for r in refrows];cash=[r['cash'] for r in refrows];enddebt='Beyond horizon' if refrows[-1]['debt']>.005 else max([r['month'] for r in refrows if r['debt']>.005]+[-1])+1
  entries=[('Custo total financiamento','Total financing cost',refcost,q(SCH,'B216'),'money'),('Serviço anual máximo','Maximum annual debt service',refds,q(SCH,'B210'),'money'),('DSCR mínimo · todas fases','Minimum DSCR · all phases',refdscr,q(SCH,'B211'),'ratio'),('Mês menor caixa','Minimum cash month',cash.index(min(cash))+1,q(CASH,'B190'),'number'),('Mês menor folga reservas','Minimum reserve headroom month',reserve.index(min(reserve))+1,q(CASH,'B191'),'number'),('Défice financiamento','Funding gap',max(r['funding_gap'] for r in refrows),'MAX('+motor('funding_gap')+')','money'),('Mês liquidação dívida','Debt repayment month',enddebt,q(SCH,'B213'),'number'),('Limite autorizado não utilizado','Unused non-revolving facility',30000000-refrows[-1]['cumulative_draw'],q(SCH,'B220'),'money')]
  s=sh(IMP)
  for j,(pt,en,v,f,kind) in enumerate(entries,30):s.put('A'+str(j),pt,tr(pt,en));s.put('B'+str(j),v);s.put('C'+str(j),0,f);s.put('D'+str(j),0,'IF(AND(ISNUMBER(B'+str(j)+'),ISNUMBER(C'+str(j)+')),C'+str(j)+'-B'+str(j)+',"N/A")');s.put('E'+str(j),0,'IF(AND(ISNUMBER(D'+str(j)+'),ABS(B'+str(j)+')>0.000001),D'+str(j)+'/ABS(B'+str(j)+'),"N/A")')
  changes[pp[IMP]]=s.finish()
  for sn,_,_ in DASHES:
   s=sh(sn);s.merge('Q40:Y41','Impacto financiamento →','HYPERLINK("#\''+sn+'\'!E650",'+tr('Impacto financiamento →','Financing impact →')+')','input');s.text('E650:Y652','Impacto bancário no projeto · Referência R4 / Ativo','Project financing impact · R4 Reference / Active','section')
   for area,pt,en in [('E654:L655','Indicador','Metric'),('M654:P655','Referência R4','R4 Reference'),('Q654:T655','Ativo','Active'),('U654:W655','Delta','Delta'),('X654:Y655','Delta %','Delta %')]:s.text(area,pt,en,'head')
   for j,(pt,en,v,f,kind) in enumerate(entries):
    rr=656+j*2;ir=30+j;s.text(f'E{rr}:L{rr+1}',pt,en,'label')
    for area,cc in [(f'M{rr}:P{rr+1}','B'),(f'Q{rr}:T{rr+1}','C'),(f'U{rr}:W{rr+1}','D'),(f'X{rr}:Y{rr+1}','E')]:s.merge(area,v if cc in ['B','C'] else 0,q(IMP,cc+str(ir)),'pct' if cc=='E' else kind)
   s.text('E674:Y677','DSCR inclui a fase de construção; valores negativos antes da exploração exigem leitura das fontes de capital e vendas. Para comparar dois perfis bancários livremente, emitir o snapshot Base/Teste no DASH 10.','DSCR includes construction; negative pre-operation values require reading equity and sales funding sources. To compare any two bank profiles, issue the Base/Test snapshot from DASH 10.','note');changes[pp[sn]]=s.finish()
  # Complete drawdown/repayment chart with voluntary repayments.
  s=sh(SCH);s.put('J189','Early repayments')
  for y in range(15):s.put('J'+str(190+y),sum(r['sweep'] for r in refrows[y*12:y*12+12]),'SUM(I'+str(5+y*12)+':I'+str(16+y*12)+')')
  c=s.cells['B230'];f=c.findtext(T('f'));s.put('B230',src.get(SCH,'B230')[1],f.replace('AND(', 'AND('+d('draw_end')+'<='+d('grace')+'+'+d('term')+',',1),int(c.get('s')));changes[pp[SCH]]=s.finish()
  drrels=E.fromstring(z.read(relpath(pp[SCH])));owner=next(r for r in drrels if r.get('Type').endswith('/drawing'));dp=resolve(pp[SCH],owner.get('Target'));rels=E.fromstring(z.read(relpath(dp)));chartpart=resolve(dp,next(r for r in rels if r.get('Type').endswith('/chart')).get('Target'));root=E.fromstring(z.read(chartpart));C='{http://schemas.openxmlformats.org/drawingml/2006/chart}';plot=root.find('.//'+C+'barChart');ser=copy.deepcopy(plot.findall(C+'ser')[-1]);ser.find(C+'idx').set('val','2');ser.find(C+'order').set('val','2');tx=ser.find(C+'tx');tx.clear();E.SubElement(tx,C+'v').text='Early repayment';vr=ser.find(C+'val/'+C+'numRef');vr.find(C+'f').text=q(SCH,'$J$190:$J$204');cache=vr.find(C+'numCache')
  if cache is not None:vr.remove(cache)
  # Series precede data-label/gap/axis settings in the chart schema.
  loc=max(plot.index(x) for x in plot.findall(C+'ser'))+1;plot.insert(loc,ser);changes[chartpart]=xml(root)
  # Informative dimensions include all newly appended cells, without changing data.
  for sn,p in pp.items():
   data=changes.get(p)
   if data is None and (sn=='V06 Motor' or re.fullmatch(r'V06 [SMO]\d+',sn)):
    raw=z.read(p);changes[p]=re.sub(rb'<dimension ref="[^"]+"',b'<dimension ref="A1:EQ184"',raw,count=1)
   elif data is not None:
    root=E.fromstring(data);cells=[rc(c.get('r')) for c in root.iter(T('c'))];dim=root.find(T('dimension'))
    if dim is not None and cells:dim.set('ref','A1:'+col(max(c[1] for c in cells))+str(max(c[0] for c in cells)+1))
    changes[p]=xml(root)
  changes['xl/workbook.xml']=xml(wb)
 temp=BM.with_suffix('.polish.xlsx');rewrite(BM,temp,changes);temp.replace(BM);print('R6.1 controls, dates, cost cards, chart and financing impact complete',flush=True)
if __name__=='__main__':apply()
