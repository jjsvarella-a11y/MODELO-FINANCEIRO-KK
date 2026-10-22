"""Muta seletores reais no OOXML, força hard recalc externo e confronta com Python."""
from engine import *
from pathlib import PurePosixPath
import zipfile,xml.etree.ElementTree as ET,subprocess,openpyxl,concurrent.futures,shutil,copy,time
META=json.loads((ROOT/'src/build_metadata.json').read_text());P=ROOT/'entrega/KK01_Modelo_Integrado_V04.xlsx';NS='http://schemas.openxmlformats.org/spreadsheetml/2006/main';REL='http://schemas.openxmlformats.org/officeDocument/2006/relationships'
CASES=[
 ('V04_Final',{},{}),
 ('V04_Capital_Prudente',{'capital_schedule':'Prudente','capital_delay':12},{}),
 ('V04_Atraso_Obra',{'construction_delay':6,'commercial_delay':6,'opening_delay':3},{}),
 ('V04_CP_Nao_Cumprida',{'bank_cp':0},{}),
 ('V04_Pari_Passu',{'funding_kind':'Empréstimo','min_equity_ratio':.3,'pari_passu':1,'subordinated_eligible':0},{}),
 ('V04_Mensal',{'dividend_frequency':1},{}),
 ('V04_Limite_Legal',{'payout_basis':'Documentado'},{}),
 ('V04_Pricing_Motor',{}, {'V04 Pricing':{'H5':1.2},'Pressupostos':{'B'+str(META['param_rows']['pricing_mode']):'Motor'}}),
 ('V04_ETC_Comissao',{'status_month':12,'etc_override':1,'etc_amount':50000000,'commit_fee':.005},{}),
]

def mutate(dest,changes):
 with zipfile.ZipFile(P) as zi:
  w=ET.fromstring(zi.read('xl/workbook.xml'));rels=ET.fromstring(zi.read('xl/_rels/workbook.xml.rels'));paths={r.attrib['Id']:r.attrib['Target'] for r in rels}
  targets={}
  for s in w.find(f'{{{NS}}}sheets'):
   if s.attrib['name'] in changes:
    target=paths[s.attrib[f'{{{REL}}}id']];target=target.lstrip('/') if target.startswith('/') else 'xl/'+target
    targets[target]=changes[s.attrib['name']]
  with zipfile.ZipFile(dest,'w',zipfile.ZIP_DEFLATED) as zo:
   for item in zi.infolist():
    data=zi.read(item.filename)
    if item.filename in targets:
     x=ET.fromstring(data)
     for addr,value in targets[item.filename].items():
      cell=x.find(f'.//{{{NS}}}c[@r="{addr}"]');assert cell is not None
      for child in list(cell):cell.remove(child)
      if isinstance(value,str):
       cell.set('t','inlineStr');node=ET.SubElement(cell,f'{{{NS}}}is');ET.SubElement(node,f'{{{NS}}}t').text=value
      else:
       cell.pop('t',None) if hasattr(cell,'pop') else cell.attrib.pop('t',None)
       ET.SubElement(cell,f'{{{NS}}}v').text=str(value)
     data=ET.tostring(x,encoding='utf-8',xml_declaration=True)
    # Cached values are invalid after input mutation. Invalidate all formula caches so
    # the external engine computes every dependency rather than preserving stale values.
    if item.filename.startswith('xl/worksheets/') and item.filename.endswith('.xml'):
     x=ET.fromstring(data)
     for cell in x.findall(f'.//{{{NS}}}c'):
      if cell.find(f'{{{NS}}}f') is not None:
       v=cell.find(f'{{{NS}}}v')
       if v is not None:cell.remove(v)
     data=ET.tostring(x,encoding='utf-8',xml_declaration=True)
    zo.writestr(item,data)

def worker(case):
 name,params,extras=case;folder=ROOT/'validacao'/name;folder.mkdir(parents=True,exist_ok=True);(folder/'out').mkdir(exist_ok=True);profile=folder/'profile/user';profile.mkdir(parents=True,exist_ok=True)
 (profile/'registrymodifications.xcu').write_text('''<?xml version="1.0"?><oor:items xmlns:oor="http://openoffice.org/2001/registry"><item oor:path="/org.openoffice.Office.Calc/Formula/Load"><prop oor:name="OOXMLRecalcMode" oor:op="fuse"><value>0</value></prop></item></oor:items>''')
 changes=copy.deepcopy(extras)
 for k,v in params.items():changes.setdefault('Pressupostos',{})['B'+str(META['param_rows'][k])]=v
 dest=folder/(name+'.xlsx');mutate(dest,changes)
 cmd=['soffice','-env:UserInstallation='+str((folder/'profile').as_uri()),'--headless','--convert-to','xlsx','--outdir',str(folder/'out'),str(dest)]
 proc=subprocess.run(cmd,capture_output=True,text=True,timeout=1200)
 output=folder/'out'/dest.name
 if proc.returncode or not output.exists():return dict(case=name,ok=False,error=proc.stdout+proc.stderr)
 # sale variant only changes input201; use a temporary process-local copy of UNITS to avoid parallel mutation.
 if name=='Venda_201_M2':
  code="import sys,json;sys.path.insert(0,"+repr(str(ROOT/'src'))+");import engine;engine.UNITS[0]['sale']=2;print(json.dumps(engine.run()['flow']))"
  expected=json.loads(subprocess.check_output(['python','-c',code],text=True))
 else:
  expected_params=dict(params)
  if name=='V04_Pricing_Motor':expected_params.update(pricing_mode='Motor',unit_factors={UNITS[0]['id']:{'view':1.2}})
  expected=run(config(**expected_params))['flow']
 w=openpyxl.load_workbook(output,data_only=True,read_only=True);fail=[];count=0
 for row,e in zip(w['Fluxo 180M'].iter_rows(min_row=5,max_row=184,values_only=True),expected):
  for k,c in META['flow_columns'].items():
   if e[k] is None:continue
   if isinstance(e[k],(int,float)):
    count+=1
    if not isinstance(row[c],(int,float)) or abs(row[c]-e[k])>max(.01,abs(e[k])*1e-9):fail.append(f'm{e["m"]} {k}: {row[c]} != {e[k]}')
 if name=='Filtro_ID_201':
  s=w['DASH Comercial'];count+=2
  if s['B11'].value!=1:fail.append('Filtro não selecionou uma fração')
  if abs(s['B12'].value-UNITS[0]['price'])>.01:fail.append('VGV filtrado errado')
  for row,e in zip(s.iter_rows(min_row=151,max_row=330,values_only=True),receipts(config())['units'][0]):
   count+=1
   if abs(row[5]-e)>.01:fail.append('Recebimento filtrado errado')
 errs=[]
 for s in w:
  for row in s:
   for c in row:
    if c.data_type=='e':errs.append((s.title,c.coordinate,c.value))
 # Return vectors must use investor outlay, not project treasury receipts.
 if name=='Sem_transferencia_KK':
  count+=1
  if w['Fluxos Retorno']['C6'].value!=-DEFAULT['equity']:fail.append('Preço de aquisição confundido com entrada na KK')
 if name=='Operador_proposta':
  count+=1
  if abs(w['Condicoes Operador']['F5'].value-.04)>.00001:fail.append('Fee proposta não selecionado')
 w.close()
 return dict(case=name,ok=not fail and not errs,comparisons=count,failures=fail[:12],formula_errors=errs[:12],params=expected_params if name=='V04_Pricing_Motor' else params,output=str(output))
if __name__=='__main__':
 import sys
 report_path=ROOT/'entrega/Testes_Seletores.json'
 previous=json.loads(report_path.read_text()) if report_path.exists() else []
 retry='--retry-failed' in sys.argv
 if retry:
  known={r['case']:r for r in previous}; selected=[c for c in CASES if c[0] not in known or not known[c[0]]['ok']]
  (ROOT/'entrega/Testes_Seletores_Primeira_Passagem.json').write_text(json.dumps(previous,ensure_ascii=False,indent=2))
 else:selected=CASES;known={}
 with concurrent.futures.ThreadPoolExecutor(max_workers=3) as pool:
  for res in pool.map(worker,selected):known[res['case']]=res;print(json.dumps(res,ensure_ascii=False),flush=True)
 results=[known[c[0]] for c in CASES]
 report_path.write_text(json.dumps(results,ensure_ascii=False,indent=2))
 raise SystemExit(0 if all(r['ok'] for r in results) else 1)
