"""Mutate real input cells, recalculate externally, verify propagation and PT/EN invariance.
Uses LibreOffice only for supplementary calculation, never Excel certification.
"""
from pathlib import Path
from zipfile import ZipFile,ZIP_DEFLATED
from lxml import etree as E
import json,sys,subprocess,concurrent.futures,openpyxl,os
OUT=Path(__file__).resolve().parents[1];WORK=Path(os.environ.get('KK01_WORKDIR','/workspace/kk01_v06'));reference=json.loads((OUT/'Referencia_Python_V06.json').read_text());N='http://schemas.openxmlformats.org/spreadsheetml/2006/main';T=lambda n:'{'+N+'}'+n;R='http://schemas.openxmlformats.org/officeDocument/2006/relationships'
def worker(case):
 key,value,expected=case;folder=WORK/'selectors'/key;folder.mkdir(parents=True,exist_ok=True);(folder/'out').mkdir(exist_ok=True);profile=folder/'profile/user';profile.mkdir(parents=True,exist_ok=True);(profile/'registrymodifications.xcu').write_text('<oor:items xmlns:oor="http://openoffice.org/2001/registry"><item oor:path="/org.openoffice.Office.Calc/Formula/Load"><prop oor:name="OOXMLRecalcMode" oor:op="fuse"><value>0</value></prop></item></oor:items>');source=WORK/'V06_modulos.xlsx';dest=folder/(key+'.xlsx')
 with ZipFile(source) as z,ZipFile(dest,'w',ZIP_DEFLATED) as zz:
  wb=E.fromstring(z.read('xl/workbook.xml'));rels={x.get('Id'):x.get('Target') for x in E.fromstring(z.read('xl/_rels/workbook.xml.rels'))};sheet=next(x for x in wb.find(T('sheets')) if x.get('name')=='V06 Controlos');target=rels[sheet.get('{'+R+'}id')];target=target.lstrip('/') if target.startswith('/') else 'xl/'+target
  for p in z.infolist():
   data=z.read(p.filename)
   if p.filename==target:
    root=E.fromstring(data);cell=root.find('.//'+T('c')+'[@r="B'+str(reference['param_rows'][key])+'"]');assert cell is not None
    for node in list(cell):cell.remove(node)
    if isinstance(value,str):cell.set('t','inlineStr');E.SubElement(E.SubElement(cell,T('is')),T('t')).text=value
    else:cell.attrib.pop('t',None);E.SubElement(cell,T('v')).text=str(value)
    data=E.tostring(root,encoding='UTF-8',xml_declaration=True)
   if p.filename.startswith('xl/worksheets/') and p.filename.endswith('.xml'):
    root=E.fromstring(data)
    for c in root.iter(T('c')):
     if c.find(T('f')) is not None:
      c.attrib.pop('t',None)
      v=c.find(T('v'))
      if v is not None:c.remove(v)
    data=E.tostring(root,encoding='UTF-8',xml_declaration=True)
   zz.writestr(p,data)
 command=['soffice','-env:UserInstallation='+str((folder/'profile').as_uri()),'--headless','--convert-to','xlsx','--outdir',str(folder/'out'),str(dest)]
 native=folder/'out'/dest.name
 if native.exists():native.unlink()
 proc=subprocess.run(command,capture_output=True,text=True,timeout=240);native=folder/'out'/dest.name
 assert proc.returncode==0 and native.exists(),proc.stderr+proc.stdout
 w=openpyxl.load_workbook(native,read_only=True,data_only=True);count=0;fail=[]
 for i,(row,e) in enumerate(zip(w['V06 Motor'].iter_rows(min_row=5,max_row=184,values_only=True),reference['scenarios'][expected]),1):
  for k,c in reference['main_columns'].items():
   count+=1;a=row[c];v=e[k]
   if not isinstance(a,(int,float)) or abs(a-v)>max(.01,abs(v)*1e-9):fail.append([i,k,a,v])
 if key=='language':
  expected_titles=['Executive Overview','Investment & Construction','Tax, Incentives & Insurance','Project Consolidated','Real Estate Development','Residential Sell & Leaseback','Hotel & MICE','Shareholding & Investment Blocks','Equity Returns & Waterfall']
  for sn,title in zip(reference['dashboards'],expected_titles):
   count+=1
   if w[sn]['E2'].value!=title:fail.append([sn,'title',w[sn]['E2'].value])
   for ni,label in enumerate(expected_titles):
    count+=1
    if w[sn].cell(7+ni*4,2).value!=label:fail.append([sn,'sidebar',ni,w[sn].cell(7+ni*4,2).value])
  with ZipFile(native) as z:
   chart=E.fromstring(z.read('xl/charts/chart1.xml'));labels=chart.xpath('//c:ser/c:tx//c:v/text()',namespaces={'c':'http://schemas.openxmlformats.org/drawingml/2006/chart'});count+=2
   if labels!=['Cash','Protected reserve']:fail.append(['chart translations',labels])
 w.close();return {'input':key,'value':value,'expected_scenario':expected,'comparisons':count,'failures':fail[:20],'passed':not fail}
if __name__=='__main__':
 cases=[('language','EN','Base'),('vat_lag',18,'VAT delayed'),('rate',.08,'Interest 8%'),('capital_delay',6,'Capital delayed')]
 with concurrent.futures.ThreadPoolExecutor(max_workers=2) as pool:
  results=[]
  for r in pool.map(worker,cases):results.append(r);print(json.dumps(r),flush=True)
 (OUT/'Testes_Seletores_V06.json').write_text(json.dumps(results,indent=2));raise SystemExit(not all(r['passed'] for r in results))
