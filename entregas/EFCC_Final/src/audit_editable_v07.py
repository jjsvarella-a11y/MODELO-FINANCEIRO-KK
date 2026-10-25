"""Structural navigation, graphics preservation, presets and bilingual evidence."""
from scenario_v07 import *
import csv

def audit():
 source=Source(FINAL7);old=Source(BM);rows=[];broken=[];untranslated=[]
 for sn,p in source.pp.items():
  root=E.fromstring(source.z.read(p))
  for c in root.iter(T('c')):
   f=c.findtext(T('f')) or ''
   if 'HYPERLINK(' not in f:continue
   targets=re.findall(r'#\'([^\']+)\'!([A-Z]+[0-9]+)',f)
   for target,address in targets:
    passed=target in source.pp and bool(re.fullmatch('[A-Z]{1,3}[1-9][0-9]*',address));rows.append([sn,c.get('r'),target,address,'Internal HYPERLINK','PASS' if passed else 'FAIL','OOXML destination inspection; native Excel execution not available'])
    if not passed:broken.append(rows[-1])
 assert not broken,broken
 with (OUT7/'Matriz_Navegacao.csv').open('w',newline='',encoding='utf-8-sig') as file:
  writer=csv.writer(file,lineterminator="\n");writer.writerow(['Origin sheet','Origin cell','Destination sheet','Destination cell','Link type','Structural result','Test method / limitation']);writer.writerows(rows)
 comparisons=[]
 for p in old.z.namelist():
  if (p.startswith('xl/charts/') or p.startswith('xl/drawings/')) and p.endswith('.xml'):
   assert old.z.read(p)==source.z.read(p),p;comparisons.append(p)
 for sn,*_ in DASHES:
  root=E.fromstring(source.z.read(source.pp[sn]));assert root.find('.//'+T('pane')) is not None,sn
  for r in range(26,32):
   for cc in ['T','V','X']:
    f=source.get(sn,cc+str(r))[0];assert f and IMP in f,(sn,cc,r)
  for addr in ['T24','V24']:
   f=source.get(sn,addr)[0];assert f and '"PT"' in f,(sn,addr)
 for sn in [BASE7,CMP7,HIST7,REPORT7,STATUS7,MET7]:
  assert 'EXEC V06' in source.get(sn,'A1')[0],sn
 actual={}
 for sn,panel in PANELS.items():
  root=E.fromstring(source.z.read(source.pp[sn]));dvs={n.get('sqref'):n.findtext(T('formula1')) for n in root.findall('.//'+T('dataValidation'))}
  for j,(key,pt,en,options) in enumerate(panel):
   for cc in ['H','N']:
    addr=cc+str(23+j*2);assert dvs[addr]=='"'+','.join(options)+'"';actual[sn+'!'+addr]=options
 schema=json.loads((WORK7/'Schema_Final.json').read_text());assert not schema,schema[:3]
 proof={'status':'PASS','internal_destinations_checked':len(rows),'broken_targets':broken,'graphic_parts_byte_identical':len(comparisons),'preserved_parts':comparisons,'mirrored_preset_controls':actual,'openxml_sdk_errors':len(schema),'source_model_commit':'f5409903c326fb65fa353ce7b011f9b3f51c732f','native_excel365':'NOT TESTED','powershell_native_runner':'SUPPLIED; NOT EXECUTED IN LINUX','official_logo_binary':'Absent from source; original V06 Identidade insertion area preserved'};(OUT7/'Auditoria_Estrutural.json').write_text(json.dumps(proof,indent=2,ensure_ascii=False));print(proof['status'],len(rows),'links',len(comparisons),'original graphic parts')
if __name__=='__main__':audit()
