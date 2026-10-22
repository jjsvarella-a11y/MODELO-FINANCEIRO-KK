from pathlib import Path
from zipfile import ZipFile
from lxml import etree as E
import json,re,posixpath,hashlib,sys
import fitz,openpyxl
from extend import paths,T,DASHES
ROOT=Path(__file__).resolve().parents[1]
def validate():
 p=ROOT/'KK01_V06_Executive_Master.xlsx';checks=[]
 def ck(name,ok,detail=None):checks.append({'name':name,'passed':bool(ok),'detail':detail})
 manifest=json.loads((ROOT/'Fontes_DASH10.json').read_text());proof=json.loads(p.with_suffix('.preservation.json').read_text());tests=json.loads((ROOT/'Testes_DASH10.json').read_text());ck('Financial engine preserved',proof['financial_cells_changed']==0 and proof['existing_cells_verified']>2000000,proof['existing_cells_verified']);ck('Current workbook matches preservation proof',hashlib.sha256(p.read_bytes()).hexdigest()==proof['result_sha256']);ck('Dynamic and language tests passed',all(x['passed'] for x in tests))
 with ZipFile(p) as z:
  ck('CRC',z.testzip() is None);w,lookup=paths(z);names=[s.get('name') for s in w.find(T('sheets'))];ck('Ten dashboards first, names preserved',names[:10]==DASHES+['DASH 10 V06']);r=E.fromstring(z.read(lookup['DASH 10 V06']));ns={'m':T('x')[1:].split('}')[0]};setup=r.find(T('pageSetup'));ck('A4 portrait at legible 100% print scale',setup.get('paperSize')=='9' and setup.get('orientation')=='portrait' and setup.get('scale','100')=='100');breaks=[int(x.get('id')) for x in r.findall(T('rowBreaks')+'/'+T('brk'))];ck('Nine explicit page breaks',breaks==list(range(50,500,50)),breaks)
  cells={c.get('r'):c for c in r.iter(T('c'))};ck('All financial indicators are source formulas',all(cells[b['cell']].findtext(T('f'))==b['formula'] for b in manifest['financial_bindings']),len(manifest['financial_bindings']))
  issues=[];parts=set(z.namelist())
  for part in parts:
   if not part.endswith(('.xml','.rels')):continue
   doc=E.fromstring(z.read(part))
   if part.endswith('.rels'):
    owner='' if part=='_rels/.rels' else part.rsplit('/_rels/',1)[0]+'/'+part.rsplit('/_rels/',1)[1][:-5]
    for rel in doc:
     if rel.get('TargetMode')=='External':continue
     t=rel.get('Target','').split('#',1)[0];target=t.lstrip('/') if t.startswith('/') else posixpath.normpath(posixpath.join(posixpath.dirname(owner),t))
     if t and target not in parts:issues.append([part,target])
  ck('OPC relationships and XML well formed',not issues,issues[:5])
 for lang in ['PT','EN']:
  pdf=ROOT/f'KK01_DASH10_Relatorio_{lang}.pdf'
  with fitz.open(pdf) as d:
   ck(lang+' entire unfiltered PDF is exactly ten A4 pages',len(d)==10 and all(abs(x.rect.width-595.3)<1 and abs(x.rect.height-841.89)<1 for x in d))
   ck(lang+' PDF includes every section and final page',all(f'{i:02} —' in d[i-1].get_text() for i in range(1,11)) and '10 / 10' in d[9].get_text())
 report={'checks':checks,'passed':sum(x['passed'] for x in checks),'failed':sum(not x['passed'] for x in checks),'microsoft_excel_365':'NOT TESTED','financial_engine_changes':0,'pdf_export_engine':'LibreOffice supplementary export','scope':'DASH10 reporting extension only; no new operating simulators or financial optimization from the broader attachment'};(ROOT/'Validacao_Entrega.json').write_text(json.dumps(report,ensure_ascii=False,indent=2));print(json.dumps(report,ensure_ascii=False,indent=2));return report
if __name__=='__main__':raise SystemExit(validate()['failed']>0)
