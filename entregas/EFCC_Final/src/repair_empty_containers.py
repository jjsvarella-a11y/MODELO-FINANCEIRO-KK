"""Remove empty optional worksheet containers with mandatory child records."""
from pathlib import Path
from zipfile import ZipFile,ZIP_DEFLATED
from lxml import etree as E
import hashlib,json,shutil
from assemble import paths,T,xml
ROOT=Path(__file__).resolve().parents[1]
# SpreadsheetML CT_Hyperlinks.hyperlink and CT_DataValidations.dataValidation
# have minOccurs=1. Omit the optional parent when there are no records.
REQUIRED_CHILDREN={'hyperlinks':'hyperlink','dataValidations':'dataValidation','mergeCells':'mergeCell','cols':'col','conditionalFormatting':'cfRule'}
def empty_containers(root):
 return [node for tag,child in REQUIRED_CHILDREN.items() for node in root.findall(T(tag)) if not node.findall(T(child))]
def normalize(root):
 removed=[]
 for node in empty_containers(root):removed.append(E.QName(node).localname);root.remove(node)
 return removed
def repair():
 master=ROOT/'KK01_V06_Executive_Master.xlsx'
 evidence=ROOT/'Correcao_Reparacao_XML.json'
 if evidence.exists() and json.loads(evidence.read_text()).get('after_sha256')==hashlib.sha256(master.read_bytes()).hexdigest():
  print('R1 already applied; no files changed');return
 before=Path('/workspace/kk01_efcc/checkpoints/Before_EmptyContainers.xlsx');shutil.copy2(master,before);dest=master.with_suffix('.repair.xlsx');changed=[];parts=[]
 with ZipFile(before) as z,ZipFile(dest,'w',ZIP_DEFLATED,compresslevel=6) as out:
  _,pp=paths(z);names={p:sn for sn,p in pp.items()}
  for item in z.infolist():
   data=z.read(item.filename)
   if item.filename in names:
    root=E.fromstring(data);removed=normalize(root)
    if removed:
     changed.append({'sheet':names[item.filename],'part':item.filename,'removed_empty_elements':removed});parts.append(item.filename);data=xml(root)
   out.writestr(item,data)
 expected={'xl/worksheets/sheet492.xml':['hyperlinks'],'xl/worksheets/sheet323.xml':['hyperlinks'],'xl/worksheets/sheet328.xml':['dataValidations']}
 assert {r['part']:r['removed_empty_elements'] for r in changed}==expected,changed
 with ZipFile(before) as a,ZipFile(dest) as b:
  assert a.namelist()==b.namelist();unchanged=0;cells=0
  for part in a.namelist():
   if part not in parts:assert a.read(part)==b.read(part),part;unchanged+=1
   else:
    old,new=E.fromstring(a.read(part)),E.fromstring(b.read(part));oc=old.findall('.//'+T('c'));nc=new.findall('.//'+T('c'));assert len(oc)==len(nc)
    assert [E.tostring(c,method='c14n') for c in oc]==[E.tostring(c,method='c14n') for c in nc],part;cells+=len(oc)
    normalize(old);assert E.tostring(old,method='c14n')==E.tostring(new,method='c14n'),part
  assert b.testzip() is None
 dest.replace(master)
 proof={'status':'PASS','release':'V06.3-SYNC-R1','reported_native_excel_status_before':'User screenshot confirms Excel repaired all three listed parts','native_excel365_after_fix':'NOT TESTED','before_sha256':hashlib.sha256(before.read_bytes()).hexdigest(),'after_sha256':hashlib.sha256(master.read_bytes()).hexdigest(),'changes':changed,'all_other_zip_parts_byte_identical':unchanged,'cell_records_preserved_in_changed_parts':cells,'all_cell_formulas_values_styles_and_chart_parts_preserved':True,'regression_before_invalid_containers':3,'regression_after_invalid_containers':0,'scope':'Mandatory child cardinality checks; not full Open XML SDK or native Excel certification'}
 (ROOT/'Correcao_Reparacao_XML.json').write_text(json.dumps(proof,ensure_ascii=False,indent=2));print(proof)
if __name__=='__main__':repair()
