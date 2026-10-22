"""Route matrix and interface dependency checks; no claim of native click testing."""
from pathlib import Path
from zipfile import ZipFile
import json,re,csv
from assemble import paths,T,E
from premium_ui import ROOT,WORK,DASHES,CENTER
from premium_integrate import BACK

def validate():
 path=ROOT/'KK01_V06_Executive_Master.xlsx';spec=json.loads((WORK/'Premium_Spec.json').read_text());routes=[];fail=[]
 with ZipFile(path) as z:
  w,pp=paths(z);visible={s.get('name') for s in w.find(T('sheets')) if s.get('state','visible')=='visible'};roots={}
  for sn,p in pp.items():
   if sn not in visible:continue
   r=E.fromstring(z.read(p));roots[sn]=r;seen=set()
   for c in r.iter(T('c')):
    f=c.findtext(T('f')) or ''
    if 'HYPERLINK(' not in f:continue
    for dest,addr in re.findall(r'"#\x27([^\x27]+)\x27!([A-Z]+\d+)"',f):
     seen.add((c.get('r'),dest,addr))
   for h in r.findall(T('hyperlinks')+'/'+T('hyperlink')):
    m=re.fullmatch(r"'([^']+)'!([A-Z]+\d+)",h.get('location',''))
    if m:seen.add((h.get('ref'),m[1],m[2]))
   for cell,dest,addr in sorted(seen):
    ok=dest in visible
    if not ok:fail.append([sn,cell,dest,'missing or hidden destination'])
    routes.append({'origin':sn+'!'+cell,'destination':dest+'!'+addr,'type':'Excel internal link','structural_test':'PASS' if ok else 'FAIL','excel365_click_test':'NOT TESTED'})
  if 'V06 Visual' in visible:fail.append(['V06 Visual','technical chart projection must be hidden'])
  for sn,_,_ in DASHES:
   for target,_,_ in DASHES:
    if not any(x['origin'].startswith(sn+'!B') and x['destination']==target+'!A1' for x in routes):fail.append([sn,'sidebar missing',target])
   for n,(target,_,_) in enumerate(DASHES):
    addr=__import__('xlsxwriter').utility.xl_col_to_name(4+n)+'6'
    if not any(x['origin']==sn+'!'+addr and x['destination']==target+'!A1' for x in routes):fail.append([sn,'missing frozen route',target])
   if not any(x['origin']==sn+'!B6' and x['destination']==sn+'!E6' for x in routes):fail.append([sn,'Menu must open frozen dashboard menu'])
   view=roots[sn].find(T('sheetViews')+'/'+T('sheetView'));pane=view.find(T('pane'))
   if pane is None or pane.get('xSplit')!='4' or pane.get('ySplit')!='6':fail.append([sn,'freeze panes'])
  for sn,owner in BACK.items():
   if sn not in visible:continue
   if not any(x['origin'].startswith(sn+'!') and x['destination']=='EXEC V06!A1' for x in routes):fail.append([sn,'missing Home'])
   if not any(x['origin'].startswith(sn+'!') and x['destination']==owner+'!A1' for x in routes):fail.append([sn,'wrong Back',owner])
  for a in spec['aliases']:
   r=roots.get(a['source_sheet'])
   if r is None:r=E.fromstring(z.read(pp[a['source_sheet']]))
   c=r.find('.//'+T('c')+f'[@r="{a["source_cell"]}"]');expected=f"'{a['destination_sheet']}'!{a['destination_cell']}"
   if c.findtext(T('f'))!=expected:fail.append(['canonical input alias',a,expected])
  for a in spec['inputs']:
   c=roots[a['sheet']].find('.//'+T('c')+f'[@r="{a["cell"]}"]')
   if c is None or c.find(T('f')) is not None:fail.append(['editable input',a])
   if isinstance(a['value'],(int,float)):
    rules=roots[a['sheet']].findall(T('dataValidations')+'/'+T('dataValidation'))
    if not any(d.get('sqref')==a['cell'] for d in rules):fail.append(['missing numeric validation',a])
  for sn,cells in [('DASH 05 V06',['H24','H26','H28','N28','N30']),('DASH 04B V06',['H24','N24']),('DASH 07 V06',['N26'])]:
   for addr in cells:
    node=roots[sn].find('.//'+T('c')+f'[@r="{addr}"]')
    if node is None or node.find(T('f')) is None:fail.append([sn,addr,'missing effective readout'])
  for sn,addresses,selector,choice in [('DASH 05 V06',['H25','H27'],'H23','Custom'),('DASH 05 V06',['H29','H30','N32','H34','N34'],'H32','Custom Terms'),('DASH 04B V06',['H25','N25'],'H23','Custom Terms'),('DASH 07 V06',['N25'],'H29','Custom')]:
   for addr in addresses:
    rules=roots[sn].findall(T('conditionalFormatting'))
    if not any(addr in rule.get('sqref','') and choice in ''.join(rule.itertext()) for rule in rules):fail.append([sn,addr,'missing inactive Custom style'])
 result={'status':'PASS' if not fail else 'FAIL','scope':'Route destinations, ten sidebars, Home/Back ownership, freeze panes, canonical editable inputs','route_count':len(routes),'canonical_input_aliases':len(spec['aliases']),'editable_front_panel_inputs':len(spec['inputs']),'failures':fail,'native_excel365':'NOT TESTED'}
 (ROOT/'Validacao_Navegacao.json').write_text(json.dumps(result,indent=2,ensure_ascii=False))
 with (ROOT/'Matriz_Navegacao.csv').open('w',newline='',encoding='utf-8-sig') as f:
  wr=csv.DictWriter(f,fieldnames=['origin','destination','type','structural_test','excel365_click_test'],delimiter=';',lineterminator='\n');wr.writeheader();wr.writerows(routes)
 print(result);assert not fail,fail
if __name__=='__main__':validate()
