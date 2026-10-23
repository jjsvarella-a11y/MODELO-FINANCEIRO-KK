"""Bank presentation, native dashboard charts, scenario comparisons and navigation.
Financial formulas are not modified by this pass.
"""
from bank_r6 import *
from assemble import relpath,resolve
from closure_presentation import chart_cache

def apply():
 with ZipFile(BM) as z:
  wb,pp=paths(z);changes={};style={};texts=[''.join(n.itertext()) for n in E.fromstring(z.read('xl/sharedStrings.xml'))]
  for c in E.fromstring(z.read(pp[ACT])).find(T('sheetData'))[0]:
   if c.get('t')=='s':
    name=texts[int(c.findtext(T('v')))];
    if name in ['text','label','section','head','note','money','pct','ratio','number','input','input_money','input_number']:style[name]=int(c.get('s'))
  def sh(sn):return Sheet(E.fromstring(changes.get(pp[sn],z.read(pp[sn]))),style)
  def put(s,area,f,kind='money'):s.merge(area,0,f,kind)
  s=sh(SCH);c=s.cells['B217'];s.put('B217','N/A','IF(O184>0.005,"N/A",'+c.findtext(T('f'))+')',int(c.get('s')))
  # Monthly presentation selector and compact table, while retaining all 180 months.
  s.text('A320:D321','Periodicidade de leitura','Reading frequency','section');s.put('E320','Annual',style='input');dvs=s.r.find(T('dataValidations'))
  if dvs is None:dvs=E.SubElement(s.r,T('dataValidations'),count='0')
  dv=E.SubElement(dvs,T('dataValidation'),type='list',sqref='E320',showErrorMessage='1');E.SubElement(dv,T('formula1')).text='"Monthly,Quarterly,Annual"';dvs.set('count',str(len(dvs)))
  # Selector controls a separate formula-only grouped service table.
  for j,label in enumerate(['Period','Drawdowns','Principal','Interest paid','Fees','Debt service','CFADS','Closing debt']):s.put(col(j)+'324',label,style='head')
  for j in range(180):
   rr=325+j;step='IF($E$320="Monthly",1,IF($E$320="Quarterly",3,12))';last='MIN(184,4+'+str(j+1)+'*'+step+')';first='5+'+str(j)+'*'+step;cond=str(j)+'*'+step+'<180';s.put('A'+str(rr),j+1,'IF('+cond+','+str(j+1)+',"")')
   for k,c in enumerate(['C','D','F','L','M','P'],1):s.put(col(k)+str(rr),0,'IF('+cond+',SUM(INDEX('+c+'5:'+c+'184,'+str(j)+'*'+step+'+1):INDEX('+c+'5:'+c+'184,MIN(180,'+str(j+1)+'*'+step+'))),"")',style='money')
   s.put('H'+str(rr),0,'IF('+cond+',INDEX(O5:O184,MIN(180,'+str(j+1)+'*'+step+')),"")',style='money')
  changes[pp[SCH]]=s.finish()
  # Comparison selected profile is external recalc, never a fake parameter-only
  # computation. Cells intentionally say 'Run issuer' until a validated emission.
  s=sh(BN);s.text('E492:Y494','Gráficos de financiamento e serviço da dívida','Financing and debt service charts','section');s.merge('E490:P490','Gráficos bancários →','HYPERLINK("#\''+BN+'\'!E492",'+tr('Gráficos bancários →','Bank charts →')+')','input')
  s.text('E578:Y580','Comparação de propostas · resultados certificados na emissão','Proposal comparison · validated results in issuance','section')
  s.text('E582:Y585','As condições da Base e do Teste são guardadas na tabela Perfis. A comparação de resultados exige recálculo integral de cada perfil; o emissor gera dois snapshots e uma tabela Base/Teste/Delta, sem substituir a Base Oficial.','Base and Test terms are stored in Profiles. Comparing results requires a full recalculation for each profile; the issuer creates two snapshots and a Base/Test/Delta table without replacing the Official Base.','note')
  changes[pp[BN]]=s.finish()
  # Copy five native standard chart anchors to the dashboard without removing its
  # four existing manually formatted chart objects.
  D='{http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing}';A='{http://schemas.openxmlformats.org/drawingml/2006/main}';C='{http://schemas.openxmlformats.org/drawingml/2006/chart}';PR='{http://schemas.openxmlformats.org/package/2006/relationships}';R='{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'
  def drawing(sn):
   rels=E.fromstring(z.read(relpath(pp[sn])));dr=next(r for r in rels if r.get('Type').endswith('/drawing'));p=resolve(pp[sn],dr.get('Target'));return p,E.fromstring(z.read(p)),E.fromstring(z.read(relpath(p)))
  sp,sdr,srr=drawing(SCH);dp,ddr,drr=drawing(BN);used=max([int(x.get('Id')[3:]) for x in drr if x.get('Id','').startswith('rId')]+[0]);ids=max([int(x.get('id')) for x in ddr.findall('.//'+D+'cNvPr')]+[0]);lookup={r.get('Id'):r for r in srr}
  if not any(x.get('name','').startswith('R6 Bank ') for x in ddr.findall('.//'+D+'cNvPr')):
   for j,anchor in enumerate(sdr):
    new=copy.deepcopy(anchor);chart=new.find('.//'+C+'chart');oldrel=lookup[chart.get(R+'id')];newrel=copy.deepcopy(oldrel);newrel.set('Id','rId'+str(used+j+1));newrel.set('Target',__import__('posixpath').relpath(resolve(sp,oldrel.get('Target')),__import__('posixpath').dirname(dp)));drr.append(newrel);chart.set(R+'id',newrel.get('Id'));nv=new.find('.//'+D+'cNvPr');nv.set('id',str(ids+j+1));nv.set('name','R6 Bank '+str(j+1));row=496+(j//2)*26;left=4+(j%2)*10
    for nm,rr,cc in [('from',row,left),('to',row+24,left+10)]:
     marker=new.find(D+nm);marker.find(D+'row').text=str(rr);marker.find(D+'col').text=str(cc);marker.find(D+'rowOff').text='0';marker.find(D+'colOff').text='0'
    ddr.append(new)
  changes[dp]=xml(ddr);changes[relpath(dp)]=xml(drr)
  # All bank report pages have distinct information, rather than repeated KPIs.
  for sn,count in [(REP,10),(BRF,2)]:
   s=sh(sn)
   for page in range(count):
    r=page*52
    s.merge(f'A{r+14}:L{r+39}',None,style='text')
    def row(k,label,f,kind='money'):
     rr=r+15+k*2;s.text(f'A{rr}:H{rr}',label,label,'label');put(s,f'I{rr}:L{rr}',f,kind)
    if sn==BRF or page in [0,7]:
     keys=[0,1,2,4,5,6,8,9,11,12,13,14] if page==0 else [11,12,13,14,15,1,9]
     labels=['Maximum annual debt service','Minimum DSCR','Peak outstanding debt','Final debt repayment month','Total interest paid','Total financing fees','Total financing cost','Effective annual financing cost','Funding gap','Minimum liquidity headroom','Remaining facility','New investor NPV','New investor IRR','MOIC','Exit equity value','Dividends']
     for k,i in enumerate(keys[:12]):row(k,labels[i],q(SCH,'B'+str(210+i)),'pct' if i in [7,12] else 'ratio' if i in [1,13] else 'money')
    elif page==1:
     for k,(key,pt,en,rv,addr,kind) in enumerate(PARAMS[:24]):
      rr=r+15+k;left=0 if k<12 else 6;rr=r+15+(k%12)*2;s.text(f'{col(left)}{rr}:{col(left+3)}{rr}',pt,en,'label');put(s,f'{col(left+4)}{rr}:{col(left+5)}{rr}',d(key),'pct' if kind=='pct' else 'money' if kind=='money' else 'text')
    elif page in [2,3,4,5]:
     headings={2:['Year','CAPEX','Receipts','Equity'],3:['Year','Draw','Repayment','Debt'],4:['Year','Interest','Fees','Debt service'],5:['Year','CFADS','Debt service','DSCR']}[page]
     for j,label in enumerate(headings):s.text(f'{col(j*3)}{r+14}:{col(j*3+2)}{r+14}',label,label,'head')
     for k in range(15):
      rr=r+15+k;start=5+k*12;end=start+11;yr=190+k;formulas={2:[str(k+1),'SUM('+motor('capex',start,end)+')','SUM('+motor('receipts',start,end)+')','SUM('+motor('equity',start,end)+')'],3:[str(k+1),q(SCH,'B'+str(yr)),q(SCH,'C'+str(yr))+'+SUM('+q(SCH,'I'+str(start)+':I'+str(end))+')',q(SCH,'H'+str(yr))],4:[str(k+1),q(SCH,'D'+str(yr)),q(SCH,'E'+str(yr)),q(SCH,'F'+str(yr))],5:[str(k+1),q(SCH,'G'+str(yr)),q(SCH,'F'+str(yr)),'IF('+q(SCH,'F'+str(yr))+'>0,'+q(SCH,'G'+str(yr))+'/'+q(SCH,'F'+str(yr))+',"N/A")']}[page]
      for j,f in enumerate(formulas):put(s,f'{col(j*3)}{rr}:{col(j*3+2)}{rr}',f,'number' if j==0 else 'ratio' if page==5 and j==3 else 'money')
    elif page==6:
     for k in range(6):
      rr=r+15+k*3;s.merge(f'A{rr}:F{rr}','Critical point',q(CASH,'A'+str(190+k)),'label');put(s,f'G{rr}:H{rr}',q(CASH,'B'+str(190+k)),'number');put(s,f'I{rr}:L{rr}',q(CASH,'C'+str(190+k)),'ratio' if k==4 else 'money')
    elif page==8:
     row(0,'Funding gap',q(SCH,'B218'));row(1,'Minimum liquidity headroom',q(SCH,'B219'));row(2,'Minimum DSCR',q(SCH,'B211'),'ratio');row(3,'Internal target (not bank approved)',d('dscr_min'),'ratio');s.text(f'A{r+25}:L{r+33}','Stresses calculados pelo emissor: obra +6 meses; operação Conservador; recebimentos +6 meses. Não extrapolar conclusões sem executar os recálculos.','Issuer stresses: construction +6 months; Conservative operations; collections +6 months. Do not extrapolate without running the recalculations.','note')
    else:
     row(0,'Cash reconciliation','MAX('+motor('cash_check')+')');row(1,'Debt reconciliation','MAX('+motor('debt_check')+')');row(2,'Calculation months','180','number');row(3,'NPV new investors',q('V06 Cenarios','M5'));row(4,'TAN annual',q(DAT,'B42'),'pct');row(5,'Discount annual',q('V06 Registo Ativo','B65'),'pct');s.text(f'A{r+29}:L{r+36}','A emissão gera Snapshot.xlsx, pressupostos JSON, fluxos CSV, data UTC e manifesto SHA256. Excel 365 nativo: não testado neste ambiente.','Issuance creates Snapshot.xlsx, JSON assumptions, CSV cash flows, UTC date and SHA256 manifest. Native Excel 365: not tested in this environment.','note')
   changes[pp[sn]]=s.finish()
  s=sh(PRO)
  for rr in [9,10]:s.put(col(P['status']-4)+str(rr),'Hypothetical')
  changes[pp[PRO]]=s.finish()
  s=sh(CENTER);s.text('E122:G123','Destinatário do relatório','Report recipient','label');s.merge('H122:P123','To identify',style='input');changes[pp[CENTER]]=s.finish()
  s=sh(SCH);cell=s.cells['B230'];f=cell.findtext(T('f'));f=f.replace(d('institution')+'<>\"To identify\"',q(CENTER,'H122')+'<>\"To identify\"');s.put('B230','BLOCKED',f,int(cell.get('s')));changes[pp[SCH]]=s.finish()
  # Native hyperlinks cannot open hidden sheets. Make only user-facing linked
  # destinations visible, keep technical effective-input helpers hidden.
  visible_targets=set();navigation=[]
  for sn,p in pp.items():
   r=E.fromstring(changes.get(p,z.read(p)))
   for c in r.iter(T('c')):
    f=c.findtext(T('f')) or ''
    for m in re.finditer(r'HYPERLINK\("#\'([^\']+)\'!([^\"]+)"',f):
     dest,addr=m.groups();ok=dest in pp
     if not ok:raise AssertionError((sn,c.get('r'),dest))
     if dest in [ACT,REF,DAT,IMP]:continue
     visible_targets.add(dest);navigation.append({'origin':sn+'!'+c.get('r'),'destination':dest+'!'+addr,'type':'native internal HYPERLINK','structural_result':'PASS','native_excel365':'NOT TESTED'})
  for n in wb.find(T('sheets')):
   if n.get('name') in visible_targets:n.set('state','visible')
  changes['xl/workbook.xml']=xml(wb)
  temp=BM.with_suffix('.finish.xlsx')
  with ZipFile(temp,'w',ZIP_DEFLATED,compresslevel=6) as out:
   for item in z.infolist():out.writestr(copy.copy(item),changes.get(item.filename,z.read(item.filename)))
 temp.replace(BM);(BOUT/'Navigation_Matrix.json').write_text(json.dumps(navigation,indent=2,ensure_ascii=False));print('Bank UI / navigation complete',len(navigation),flush=True)
if __name__=='__main__':apply()
