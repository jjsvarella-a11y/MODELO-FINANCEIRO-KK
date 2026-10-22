from pathlib import Path
from zipfile import ZipFile
import hashlib,json,re
from assemble import paths
ROOT=Path(__file__).resolve().parents[1];FINAL=ROOT/'KK01_V06_Executive_Master.xlsx';BEFORE=Path('/workspace/kk01_efcc/Release_Master.xlsx')
exclude=['EXEC V06','DASH 01 V06','DASH 02 V06','DASH 03 V06','DASH 04A V06','DASH 04B V06','DASH 05 V06','DASH 06 V06','DASH 07 V06','DASH 10 V06','V06 Brief']
with ZipFile(BEFORE) as a,ZipFile(FINAL) as b:
 _,pa=paths(a);_,pb=paths(b);checked=[];changed=[];selected=[]
 for sn,part in pa.items():
  if sn in exclude:continue
  ar=re.search(rb'<sheetData(?:\s[^>]*)?>.*?</sheetData>',a.read(part),re.S).group();br=re.search(rb'<sheetData(?:\s[^>]*)?>.*?</sheetData>',b.read(pb[sn]),re.S).group();checked.append(sn)
  if ar!=br:changed.append(sn)
 for sn,part in pb.items():
  header=b.read(part).split(b'<sheetData',1)[0]
  if re.search(rb'<sheetView\b[^>]*tabSelected="(?:true|1)"',header):selected.append(sn)
 assert not changed,changed;assert selected==['EXEC V06'],selected
 result={'status':'PASS','sheetData_blocks_compared':len(checked),'changes':changed,'selected_sheets':selected,'identity_sheet_present':'V06 Identidade' in pb,'financial_changes_from_logo_area':False,'sha256':hashlib.sha256(FINAL.read_bytes()).hexdigest(),'native_excel365':'NOT TESTED'}
 (ROOT/'Prova_Final_Identidade.json').write_text(json.dumps(result,indent=2));p=ROOT/'Preservacao.json';old=json.loads(p.read_text());old['result_sha256']=result['sha256'];old['identity_financial_sheetData_unchanged']=len(checked);p.write_text(json.dumps(old,indent=2));print(result)
