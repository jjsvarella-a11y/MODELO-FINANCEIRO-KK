"""Separate the authoritative V04 archive and remove legacy worksheet engines.
Fail before removing anything if a retained formula/chart still references a legacy
sheet. Preserve retained worksheet XML byte-for-byte and prune unreachable parts.
"""
from pathlib import Path
from zipfile import ZipFile,ZIP_DEFLATED
from lxml import etree as E
import re,sys,json,posixpath,hashlib,shutil
from assemble import paths,T,xml,resolve,relpath
ROOT=Path(__file__).resolve().parents[1]
R='{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id'
VT='http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes'
AP='http://schemas.openxmlformats.org/officeDocument/2006/extended-properties'

def separate(source):
 source=Path(source);temp=source.with_suffix('.current_only.xlsx');archive=ROOT.parent/'Arquivo/V04/KK01_Modelo_Integrado_V04_Excel_Fix.xlsx';original=Path('/workspace/kk01_v06/sources/KK01_Modelo_Integrado_V04_Excel_Fix.xlsx')
 if not original.exists():original=archive
 archive.parent.mkdir(parents=True,exist_ok=True)
 if not archive.exists():shutil.copy2(original,archive)
 original_sha=hashlib.sha256(original.read_bytes()).hexdigest();assert hashlib.sha256(archive.read_bytes()).hexdigest()==original_sha
 issues=[];refs=0;preserved=0
 with ZipFile(source) as z:
  w,pp=paths(z);sheets=w.find(T('sheets'));old=list(sheets);kept=[s for s in old if s.get('name').startswith('V06 ') or s.get('name').endswith(' V06')];removed=[s for s in old if s not in kept];removednames={s.get('name') for s in removed};keptnames={s.get('name') for s in kept};assert len(removed)==313
  # Actual formula dependencies, including hyperlinks and chart series.
  scanparts=[pp[n] for n in keptnames]+[p for p in z.namelist() if p.startswith('xl/charts/') and p.endswith('.xml')]
  referenced=set()
  for part in scanparts:
   root=E.fromstring(z.read(part));chart=part.startswith('xl/charts/')
   for node in root.iter():
    local=E.QName(node).localname
    if local=='f' and node.text:
     refs+=1
     # All generated V06 inter-sheet references are explicit and quoted.
     for name in re.findall(r"'((?:[^']|'')+)'!",node.text):
      name=name.replace("''", "'");referenced.add(name)
      if name in removednames and not chart:issues.append([part,node.getparent().get('r'),name])
    elif local=='hyperlink':
     value=node.get('location','');match=re.match(r"'?([^']+)'?!",value)
     if match and match[1] in removednames:issues.append([part,'hyperlink',match[1]])
  assert not issues,issues[:10]
  removedids={s.get(R) for s in removed};indices={i:kept.index(s) for i,s in enumerate(old) if s in kept};sheets[:]=kept
  dn=w.find(T('definedNames'));dropped_names=0
  if dn is not None:
   for n in list(dn):
    i=n.get('localSheetId')
    if i is not None:
     if int(i) not in indices:dn.remove(n);dropped_names+=1
     else:n.set('localSheetId',str(indices[int(i)]))
    else:
     dep={s.replace("''", "'") for s in re.findall(r"'((?:[^']|'')+)'!",n.text or '')}
     assert not dep.intersection(removednames),('Global name depends on legacy sheet',n.get('name'))
  for v in w.find(T('bookViews')):v.set('activeTab','0');v.set('firstSheet','0')
  rels=E.fromstring(z.read('xl/_rels/workbook.xml.rels'))
  for r in list(rels):
   if r.get('Id') in removedids or r.get('Type','').endswith('/calcChain'):rels.remove(r)
  changes={'xl/workbook.xml':xml(w),'xl/_rels/workbook.xml.rels':xml(rels)}
  # Traverse package relationships: deleted legacy worksheets, drawings, images
  # and unused charts become unreachable and are physically omitted from ZIP.
  parts=set(z.namelist());reachable=set();todo=['']
  while todo:
   owner=todo.pop();rp='_rels/.rels' if owner=='' else relpath(owner)
   if rp not in parts or rp in reachable:continue
   reachable.add(rp)
   for r in E.fromstring(changes.get(rp,z.read(rp))):
    if r.get('TargetMode')=='External':continue
    target=r.get('Target')
    if target.startswith('#'):continue
    dest=resolve(owner,target)
    if dest not in reachable:reachable.add(dest);todo.append(dest)
  reachable.add('[Content_Types].xml');ct=E.fromstring(z.read('[Content_Types].xml'))
  for node in list(ct):
   if node.get('PartName') and node.get('PartName').lstrip('/') not in reachable:ct.remove(node)
  changes['[Content_Types].xml']=xml(ct)
  if 'docProps/app.xml' in parts:
   app=E.fromstring(z.read('docProps/app.xml'))
   hp=app.find('{'+AP+'}HeadingPairs');tp=app.find('{'+AP+'}TitlesOfParts')
   if hp is not None:
    for n in list(hp):hp.remove(n)
    vector=E.SubElement(hp,'{'+VT+'}vector',size='2',baseType='variant');E.SubElement(E.SubElement(vector,'{'+VT+'}variant'),'{'+VT+'}lpstr').text='Worksheets';E.SubElement(E.SubElement(vector,'{'+VT+'}variant'),'{'+VT+'}i4').text=str(len(kept))
   if tp is not None:
    for n in list(tp):tp.remove(n)
    vector=E.SubElement(tp,'{'+VT+'}vector',size=str(len(kept)),baseType='lpstr')
    for s in kept:E.SubElement(vector,'{'+VT+'}lpstr').text=s.get('name')
   changes['docProps/app.xml']=xml(app)
  with ZipFile(temp,'w',ZIP_DEFLATED,compresslevel=6) as out:
   for part in z.infolist():
    if part.filename in reachable:out.writestr(part.filename,changes.get(part.filename,z.read(part.filename)))
  with ZipFile(temp) as out:
   aw,ap=paths(out);assert set(ap)==keptnames;assert not set(ap).intersection(removednames);assert out.testzip() is None
   for name in keptnames:assert out.read(ap[name])==z.read(pp[name]),name;preserved+=1
   for part in out.namelist():
    if part not in changes:assert out.read(part)==z.read(part),part
   # Kept chart dependencies must also point exclusively at remaining sheets.
   for part in out.namelist():
    if part.startswith('xl/charts/') and part.endswith('.xml'):
     for node in E.fromstring(out.read(part)).iter():
      if E.QName(node).localname=='f':
       assert not {n.replace("''", "'") for n in re.findall(r"'((?:[^']|'')+)'!",node.text or '')}.intersection(removednames),(part,node.text)
 temp.replace(source)
 result={'status':'PASS','legacy_sheets_physically_removed':len(removed),'removed_sheet_names':sorted(removednames),'retained_worksheet_count':preserved,'retained_worksheet_parts_byte_for_byte_unchanged':True,'formula_dependencies_scanned':refs,'remaining_formula_references_to_removed_sheets':issues,'removed_scoped_names':dropped_names,'unreachable_parts_removed':len(parts-reachable),'archive_path':str(archive.relative_to(ROOT.parent)),'archive_sha256':original_sha,'original_v04_byte_for_byte_preserved':True,'sha256':hashlib.sha256(source.read_bytes()).hexdigest(),'native_excel365':'NOT TESTED'}
 (ROOT/'Separacao_V04.json').write_text(json.dumps(result,indent=2,ensure_ascii=False));print({k:v for k,v in result.items() if not isinstance(v,list)})
 for file,key in [('Preservacao.json','result_sha256'),('Prova_Final_Identidade.json','sha256')]:
  p=ROOT/file;r=json.loads(p.read_text());r[key]=result['sha256'];r['latest_revision']='Separacao_V04.json: legacy moved to standalone authoritative archive; retained worksheet parts unchanged';p.write_text(json.dumps(r,ensure_ascii=False,indent=2))
 # Replace the prior hide-only status with the actual current packaging state.
 p=ROOT/'Organizacao_Folhas.json';r=json.loads(p.read_text());r['superseded_by']='Separacao_V04.json';r['legacy_sheets_in_current_workbook']=0;r['change']='Legacy removed from current workbook; archive preserved separately';r['sha256']=result['sha256'];p.write_text(json.dumps(r,ensure_ascii=False,indent=2))
if __name__=='__main__':separate(sys.argv[1])
