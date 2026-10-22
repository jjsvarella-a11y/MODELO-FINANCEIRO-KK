"""Integrate premium UI without renaming existing technical sheets.
Expose four existing residential constants, with exactly the same initial values.
"""
from pathlib import Path
from zipfile import ZipFile,ZIP_DEFLATED
from lxml import etree as E
import json,re,hashlib,copy,posixpath,sys
from assemble import paths,T,xml,resolve,relpath
from premium_ui import ROOT,WORK,BASE,Source,DASHES,CENTER,LANG
sys.path.insert(0,str(ROOT.parents[0]/'V06/src'))
from merge import merge,R
BACK={'V06 Identidade':'EXEC V06','V06 Controlos':'EXEC V06','V06 Unidades':'DASH 04A V06','V06 Capitulos':'DASH 01 V06','V06 Capital':'DASH 06 V06','V06 Hipoteses':CENTER,'V06 Cobrancas':'DASH 04A V06','V06 Motor':'EXEC V06','V06 Cenarios':'DASH 03 V06','V06 Fluxos Invest':'DASH 07 V06','V06 Blocos':'DASH 06 V06','V06 Anual':'DASH 03 V06','V06 Trimestral':'DASH 01 V06','V06 Dicionario':'EXEC V06','V06 Matriz ADR':'DASH 05 V06','V06 Equilibrio':'DASH 05 V06','V06 Otimizacao':'DASH 07 V06','V06 Operacao':'DASH 05 V06','V06 Disponibilidade':'DASH 04A V06','V06 Conciliacao Vendas':'DASH 04A V06','V06 Imprimir':CENTER,'V06 Registo Ativo':CENTER,'DASH 10 V06':CENTER,'V06 Brief':CENTER}
BACK.update({f'V06 PDF {i+1:02}':n[0] for i,n in enumerate(DASHES[:9])})
TITLES={'V06 Unidades':('Inventário de unidades','Unit inventory'),'V06 Capitulos':('Capítulos e orçamento','Development budget'),'V06 Capital':('Calendário de capital','Equity subscription calendar'),'V06 Hipoteses':('Registo de pressupostos','Assumption register'),'V06 Cobrancas':('Recebimentos documentais','Documented collections'),'V06 Motor':('Motor financeiro central','Central financial engine'),'V06 Cenarios':('Cenários financeiros','Financial scenarios'),'V06 Fluxos Invest':('Fluxos e dividendos dos investidores','Investor cash flows and dividends'),'V06 Blocos':('Blocos de investimento','Investment blocks'),'V06 Anual':('Resultados anuais','Annual results'),'V06 Trimestral':('Resultados trimestrais','Quarterly results'),'V06 Dicionario':('Dicionário de indicadores','KPI dictionary'),'V06 Matriz ADR':('Sensibilidade ADR × Ocupação','ADR × Occupancy sensitivity'),'V06 Equilibrio':('Equilíbrio na grelha testada','Break-even in tested grid'),'V06 Otimizacao':('Simulações de dividendos, reservas e fees','Dividend, reserve and fee simulations'),'V06 Operacao':('Operação, fees e reservas','Operations, fees and reserves'),'V06 Disponibilidade':('Disponibilidade de unidades','Unit availability'),'V06 Conciliacao Vendas':('Conciliação de vendas','Sales reconciliation'),'V06 Imprimir':('Impressão e exportação PDF','Print and PDF export'),'V06 Registo Ativo':('Parâmetros ativos','Active parameters'),'V06 Controlos':('Pressupostos complementares','Additional assumptions')}
def prune(z,changed):
 pending=[''];reachable={'[Content_Types].xml'}
 while pending:
  owner=pending.pop();rp='_rels/.rels' if owner=='' else relpath(owner)
  if rp in z.namelist() and rp not in reachable:
   reachable.add(rp)
   for r in E.fromstring(changed.get(rp,z.read(rp))):
    if r.get('TargetMode')=='External':continue
    target=resolve(owner,r.get('Target'))
    if target not in reachable:reachable.add(target);pending.append(target)
 ct=E.fromstring(z.read('[Content_Types].xml'))
 for n in list(ct):
  if n.get('PartName') and n.get('PartName').lstrip('/') not in reachable:ct.remove(n)
 changed['[Content_Types].xml']=xml(ct);return reachable

def integrate():
 spec=json.loads((WORK/'Premium_Spec.json').read_text());replaced={x[0] for x in DASHES[:9]}|{'DASH 10 V06','V06 Brief'};trim=WORK/'Premium_Backend.xlsx';combined=WORK/'Premium_Combined.xlsx';dest=ROOT/'KK01_V06_Executive_Master.xlsx'
 with ZipFile(BASE) as z:
  w,pp=paths(z);ss=w.find(T('sheets'));old=[s.get('name') for s in ss];removed=set()
  for s in list(ss):
   if s.get('name') in replaced:removed.add(s.get('{'+R+'}id'));ss.remove(s)
  rels=E.fromstring(z.read('xl/_rels/workbook.xml.rels'))
  for r in list(rels):
   if r.get('Id') in removed:rels.remove(r)
  names=w.find(T('definedNames'));kept=[s.get('name') for s in ss]
  if names is not None:
   for n in list(names):
    if n.get('localSheetId') is not None:
     name=old[int(n.get('localSheetId'))]
     if name in replaced:names.remove(n)
     else:n.set('localSheetId',str(kept.index(name)))
  changes={'xl/workbook.xml':xml(w),'xl/_rels/workbook.xml.rels':xml(rels)};reachable=prune(z,changes)
  with ZipFile(trim,'w',ZIP_DEFLATED,compresslevel=6) as zz:
   for p in z.infolist():
    if p.filename in reachable:zz.writestr(p.filename,changes.get(p.filename,z.read(p.filename)))
 merge(trim,WORK/'Premium_UI.xlsx',combined)
 changes={};exposure_count=0;engine_names=[];modified_headers=[];source=Source(BASE)
 with ZipFile(combined) as z:
  w,pp=paths(z);sheet_nodes=w.find(T('sheets'));oldnames=[s.get('name') for s in sheet_nodes];order=[x[0] for x in DASHES]+[n for n in oldnames if n not in {x[0] for x in DASHES}];byname={s.get('name'):s for s in sheet_nodes};sheet_nodes[:]=[byname[n] for n in order]
  byname['V06 Visual'].set('state','hidden')
  dn=w.find(T('definedNames'))
  if dn is not None:
   for n in dn:
    if n.get('localSheetId') is not None:n.set('localSheetId',str(order.index(oldnames[int(n.get('localSheetId'))])))
  for v in w.find(T('bookViews')):v.set('activeTab','0');v.set('firstSheet','0')
  changes['xl/workbook.xml']=xml(w)
  navstyle=E.fromstring(z.read(pp['EXEC V06'])).find('.//'+T('c')+'[@r="E5"]').get('s')
  def root(sn):return E.fromstring(changes.get(pp[sn],z.read(pp[sn])))
  def write(r,addr,f,value):
   sd=r.find(T('sheetData'));rn=int(re.search(r'\d+',addr)[0]);row=sd.find(T('row')+f'[@r="{rn}"]')
   if row is None:row=E.SubElement(sd,T('row'),r=str(rn));sd[:]=sorted(sd,key=lambda x:int(x.get('r')))
   c=row.find(T('c')+f'[@r="{addr}"]')
   if c is None:c=E.SubElement(row,T('c'),r=addr)
   for n in list(c):c.remove(n)
   if isinstance(value,str):c.set('t','str')
   else:c.attrib.pop('t',None)
   E.SubElement(c,T('f')).text=f;E.SubElement(c,T('v')).text=str(value)
   def ci(c):
    out=0
    for a in re.match('[A-Z]+',c.get('r'))[0]:out=out*26+ord(a)-64
    return out
   row[:]=sorted(row,key=ci);return c
  # One canonical editable input per operational assumption.
  for a in spec['aliases']:
   sn=a['source_sheet'];r=root(sn);write(r,a['source_cell'],f"'{a['destination_sheet']}'!{a['destination_cell']}",a['value'])
   validations=r.find(T('dataValidations'))
   if validations is not None:
    for n in list(validations):
     if n.get('sqref')==a['source_cell']:validations.remove(n)
    validations.set('count',str(len(validations)))
   changes[pp[sn]]=xml(r)
  # Exposing existing coefficients is necessary for direct residential controls.
  # Frozen Reference Case is deliberately excluded. Original values remain exact.
  for sn in pp:
   if sn!='V06 Motor' and not re.fullmatch(r'V06 [SMO]\d+',sn):continue
   r=root(sn);local=0
   for c in r.iter(T('c')):
    f=c.find(T('f'));addr=c.get('r');rownum=int(re.search(r'\d+',addr)[0]);column=re.match('[A-Z]+',addr)[0]
    if f is None or rownum<5 or column not in ['AP','AQ','AS','AU']:continue
    before=f.text
    if column=='AP':f.text=f.text.replace('(225*',"('DASH 04B V06'!$N$27*")
    elif column=='AQ':f.text=f.text.replace('(0.581+',"('DASH 04B V06'!$H$29+")
    elif column=='AS':f.text=f.text.replace('*0.35)',"*'DASH 04B V06'!$N$29)")
    else:f.text=f.text.replace('*0.04)',"*'DASH 04B V06'!$N$31)")
    assert f.text!=before,(sn,addr,before);local+=1
   assert local==720,(sn,local);exposure_count+=local;engine_names.append(sn);changes[pp[sn]]=xml(r)
  assert len(engine_names)==68
  # Compact, translated return header for every visible auxiliary sheet.
  for s in sheet_nodes:
   sn=s.get('name');r=root(sn);visible=s.get('state','visible')=='visible';new_ui=sn in replaced or sn==CENTER or sn=='V06 Visual' or sn.startswith('V06 PDF ')
   if new_ui:continue
   if visible and sn=='V06 Identidade':
    routes=[('B5','EXEC V06','A1','Início','Home'),('E5','EXEC V06','A1','Voltar ao Dashboard','Back to Dashboard'),('H5',CENTER,'H7','Parâmetros','Assumptions'),('K5','V06 Dicionario','A4','Ajuda','Help')]
   elif visible:
    assert sn in BACK,sn;owner=BACK[sn];routes=[('A2','EXEC V06','A1','Início','Home'),('B2',owner,'A1','Voltar ao Dashboard','Back to Dashboard'),('C2',owner,'H23' if owner!=CENTER else 'H7','Parâmetros','Assumptions'),('D2','V06 Dicionario','A4','Ajuda','Help')]
    if sn in TITLES:
     a,b=TITLES[sn];addr='B1' if sn=='V06 Matriz ADR' else 'A1';c=write(r,addr,f'IF({LANG}="PT","{a}","{b}")',a);c.set('s',navstyle)
     hs=r.find(T('hyperlinks'))
     if hs is not None:
      for h in list(hs):
       if h.get('ref')==addr:hs.remove(h)
   else:
    occupied=r.find('.//'+T('c')+'[@r="A1"]')
    if occupied is not None and (occupied.find(T('f')) is not None or occupied.find(T('v')) is not None):continue
    routes=[('A1','EXEC V06','A1','← Início','← Home')]
   for addr,sh,cell,a,b in routes:
    c=write(r,addr,f'HYPERLINK("#\'{sh}\'!{cell}",IF({LANG}="PT","{a}","{b}"))',a);c.set('s',navstyle)
    hs=r.find(T('hyperlinks'))
    if hs is not None:
     for h in list(hs):
      if h.get('ref')==addr:hs.remove(h)
    row=c.getparent();row.set('ht','30' if visible else '16');row.set('customHeight','1');spec['navigation'].append({'origin_sheet':sn,'origin_cell':addr,'destination_sheet':sh,'destination_cell':cell,'type':'auxiliary Back' if addr in ['B2','E5'] else 'auxiliary internal link','PT':a,'EN':b})
   for view in r.findall(T('sheetViews')+'/'+T('sheetView')):
    view.set('showGridLines','0');view.set('tabSelected','0')
   dim=r.find(T('dimension'))
   if dim is not None:dim.set('ref','A1:'+dim.get('ref').split(':')[-1])
   changes[pp[sn]]=xml(r);modified_headers.append(sn)
  # Real native PDF workflow: choose destination, recalculate, then Ctrl+P.
  r=root('V06 Imprimir')
  instructions=[('Escolha o painel ou relatório nos links abaixo.','Choose a dashboard or report using the links below.'),('Na folha de destino: Fórmulas > Calcular Agora; confirme o cenário.','On the destination sheet: Formulas > Calculate Now; confirm the scenario.'),('Prima Ctrl+P. Selecione folhas ativas e respeite a área de impressão definida.','Press Ctrl+P. Select active sheets and use the configured print area.'),('Use Exportar > Criar PDF ou a impressora Microsoft Print to PDF.','Use Export > Create PDF or the Microsoft Print to PDF printer.'),('Painéis: A3 horizontal. Relatório integral: A4 vertical, 10 páginas. Resumo: 2 páginas.','Dashboards: A3 landscape. Complete report: A4 portrait, 10 pages. Brief: 2 pages.'),('Feche a pré-visualização e use Início ou Voltar para continuar.','Close print preview and use Home or Back to continue.')]
  for i,(a,b) in enumerate(instructions,5):
   c=write(r,'A'+str(i),f'IF({LANG}=\"PT\",\"{a}\",\"{b}\")',a);c.set('s',navstyle)
  changes[pp['V06 Imprimir']]=xml(r)
  # Keep exactly one selected tab and preserve all engine cell values.
  for sn in [x[0] for x in DASHES]+['DASH 10 V06','V06 Brief']:
   r=root(sn)
   for view in r.findall(T('sheetViews')+'/'+T('sheetView')):view.set('tabSelected','1' if sn=='EXEC V06' else '0')
   changes[pp[sn]]=xml(r)
  # Audit financial expressions: the four coefficient exposures are the only
  # permitted changes to existing calculation-engine formulas, never values.
  unchanged=0
  for sn in engine_names+['V06 Referencia']:
   before=source.cells(sn);after={c.get('r'):c for c in root(sn).iter(T('c'))}
   for addr,c in before.items():
    if int(re.search(r'\d+',addr)[0])<4:continue
    n=after[addr];f=c.findtext(T('f'));nf=n.findtext(T('f'));column=re.match('[A-Z]+',addr)[0]
    assert c.findtext(T('v'))==n.findtext(T('v')) and c.get('t')==n.get('t'),(sn,addr,'cache changed')
    if sn not in engine_names or column not in ['AP','AQ','AS','AU']:assert f==nf,(sn,addr,'unexpected expression change');unchanged+=f is not None
  # Properties and reachability: remove obsolete UI drawings and chart parts.
  if 'docProps/app.xml' in z.namelist():
   app=E.fromstring(z.read('docProps/app.xml'));ns={'ep':'http://schemas.openxmlformats.org/officeDocument/2006/extended-properties','vt':'http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes'}
   v=app.find('ep:TitlesOfParts/vt:vector',ns)
   if v is not None:
    for n in list(v):v.remove(n)
    for sn in order:E.SubElement(v,'{'+ns['vt']+'}lpstr').text=sn
    v.set('size',str(len(order)))
   sizes=app.findall('ep:HeadingPairs/vt:vector/vt:variant/vt:i4',ns)
   if sizes:sizes[0].text=str(len(order))
   changes['docProps/app.xml']=xml(app)
  reachable=prune(z,changes)
  with ZipFile(dest,'w',ZIP_DEFLATED,compresslevel=6) as zz:
   for p in z.infolist():
    if p.filename in reachable:zz.writestr(p.filename,changes.get(p.filename,z.read(p.filename)))
  assert len(order)==190,len(order)
 proof={'status':'PASS','source_sha256':hashlib.sha256(BASE.read_bytes()).hexdigest(),'sha256':hashlib.sha256(dest.read_bytes()).hexdigest(),'worksheets':190,'legacy_v04_sheets':0,'original_engine_formulas_preserved':unchanged,'residential_coefficient_exposures':exposure_count,'active_engines_with_direct_residential_controls':len(engine_names),'all_existing_engine_caches_unchanged':True,'frozen_reference_engine_unchanged_except_discreet_home_link':True,'canonical_input_transfers':spec['aliases'],'new_residential_controls':spec['new_constant_inputs'],'native_excel365':'NOT TESTED','native_logo_bytes_available':False}
 (ROOT/'Integridade_Premium.json').write_text(json.dumps(proof,indent=2,ensure_ascii=False));(ROOT/'Parametros_Interface.json').write_text(json.dumps(spec['inputs'],indent=2,ensure_ascii=False));(ROOT/'Navegacao_Premium.json').write_text(json.dumps(spec['navigation'],indent=2,ensure_ascii=False));(WORK/'Premium_Spec.json').write_text(json.dumps(spec,indent=2,ensure_ascii=False));print(json.dumps(proof,ensure_ascii=False))
if __name__=='__main__':integrate()
