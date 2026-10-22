"""Replace authorised V06 modules, retain authoritative legacy sheets and all C0 inputs.
No openpyxl re-save of the 385-sheet master; relationship-aware OOXML operations.
"""
from pathlib import Path
from zipfile import ZipFile,ZIP_DEFLATED
from lxml import etree as E
import sys,copy,posixpath,json,hashlib,re
sys.path.insert(0,str(Path(__file__).resolve().parents[2]/'V06/src'))
from merge import merge,T,xml,R,P,C
BASE=Path('/workspace/kk01_efcc/checkpoints/C0_V06_DASH10.xlsx');WORK=Path('/workspace/kk01_efcc');OUT=Path(__file__).resolve().parents[1]
def paths(z):
 w=E.fromstring(z.read('xl/workbook.xml'));rels={r.get('Id'):r.get('Target') for r in E.fromstring(z.read('xl/_rels/workbook.xml.rels'))}
 return w,{s.get('name'):resolve('xl/workbook.xml',rels[s.get('{'+R+'}id')]) for s in w.find(T('sheets'))}
def resolve(owner,target):return target.lstrip('/') if target.startswith('/') else posixpath.normpath(posixpath.join(posixpath.dirname(owner),target))
def relpath(owner):
 d,n=posixpath.split(owner);return (d+'/' if d else '')+'_rels/'+n+'.rels'
def prepare_legacy():
 dest=WORK/'C0_legacy_preserved.xlsx'
 with ZipFile(BASE) as z:
  w,pp=paths(z);sheets=w.find(T('sheets'));keep=[s.get('name') for s in sheets if not (s.get('name').endswith(' V06') or s.get('name').startswith('V06 '))];assert len(keep)==313,len(keep)
  removedids=set();indices={};oldnames=[]
  for i,s in enumerate(list(sheets)):
   oldnames.append(s.get('name'))
   if s.get('name') not in keep:removedids.add(s.get('{'+R+'}id'));sheets.remove(s)
   else:indices[i]=len(indices)
  rels=E.fromstring(z.read('xl/_rels/workbook.xml.rels'))
  for r in list(rels):
   if r.get('Id') in removedids:rels.remove(r)
  dn=w.find(T('definedNames'))
  if dn is not None:
   for d in list(dn):
    if d.get('localSheetId') is not None:
     i=int(d.get('localSheetId'))
     if i not in indices:dn.remove(d)
     else:d.set('localSheetId',str(indices[i]))
  for v in w.find(T('bookViews')):v.set('activeTab','0');v.set('firstSheet','0')
  changed={'xl/workbook.xml':xml(w),'xl/_rels/workbook.xml.rels':xml(rels)}
  # Retain only reachable parts, removing orphan old V06 drawings and worksheets.
  todo=[''];reachable=set()
  while todo:
   owner=todo.pop();rp='_rels/.rels' if owner=='' else relpath(owner)
   if rp in z.namelist() and rp not in reachable:
    reachable.add(rp);rr=E.fromstring(changed.get(rp,z.read(rp)))
    for r in rr:
     if r.get('TargetMode')=='External':continue
     target=resolve(owner,r.get('Target'))
     if target not in reachable:reachable.add(target);todo.append(target)
  reachable.add('[Content_Types].xml');ct=E.fromstring(z.read('[Content_Types].xml'))
  for n in list(ct):
   if n.get('PartName') and n.get('PartName').lstrip('/') not in reachable:ct.remove(n)
  changed['[Content_Types].xml']=xml(ct)
  with ZipFile(dest,'w',ZIP_DEFLATED,compresslevel=6) as zz:
   for part in z.infolist():
    if part.filename in reachable:zz.writestr(part.filename,changed.get(part.filename,z.read(part.filename)))
  assert all(pp[n] in reachable for n in keep)
 return dest,keep

def inject_caches(original,recalculated,dest):
 count=0
 with ZipFile(original) as a,ZipFile(recalculated) as b,ZipFile(dest,'w',ZIP_DEFLATED,compresslevel=6) as z:
  _,pa=paths(a);_,pb=paths(b);bstrings=[]
  if 'xl/sharedStrings.xml' in b.namelist():bstrings=[''.join(x.itertext()) for x in E.fromstring(b.read('xl/sharedStrings.xml'))]
  partnames={v:k for k,v in pa.items()}
  for part in a.infolist():
   data=a.read(part.filename)
   if part.filename in partnames:
    sn=partnames[part.filename];ar=E.fromstring(data);br=E.fromstring(b.read(pb[sn]));bc={c.get('r'):c for c in br.iter(T('c'))}
    for c in ar.iter(T('c')):
     if c.find(T('f')) is None:continue
     other=bc.get(c.get('r'));assert other is not None,(sn,c.get('r'));v=other.find(T('v'));typ=other.get('t')
     assert typ!='e',(sn,c.get('r'),v.text if v is not None else '')
     cv=c.find(T('v'))
     if cv is not None:c.remove(cv)
     c.attrib.pop('t',None)
     if typ=='s':c.set('t','str');value=bstrings[int(v.text)]
     elif typ=='inlineStr':c.set('t','str');value=''.join(other.find(T('is')).itertext())
     else:
      value=v.text if v is not None else ''
      if typ in ['str','b']:c.set('t',typ)
     E.SubElement(c,T('v')).text=value;count+=1
    data=xml(ar)
   z.writestr(part.filename,data)
 return count

def assemble(modules):
 legacy,names=prepare_legacy();dest=OUT/'KK01_V06_Executive_Master.xlsx';proof=merge(legacy,modules,dest)
 # Original legacy cell records and XML remain unchanged (view-selection exception only).
 with ZipFile(BASE) as a,ZipFile(dest) as b:
  _,pa=paths(a);w,pb=paths(b);unchanged=[]
  for sn in names:
   ar=E.fromstring(a.read(pa[sn]));br=E.fromstring(b.read(pb[sn]));ac={c.get('r'):E.tostring(c,method='c14n') for c in ar.iter(T('c'))};bc={c.get('r'):E.tostring(c,method='c14n') for c in br.iter(T('c'))}
   # Namespace contexts may change; compare attributes and descendants semantically.
   def record(root):return [(c.get('r'),dict(c.attrib),[(n.tag,dict(n.attrib),n.text) for n in c.iter() if n is not c]) for c in root.iter(T('c'))]
   assert record(ar)==record(br),sn;unchanged.append(sn)
  assert len(set(pb))==len(pb);assert b.testzip() is None
 proof.update(legacy_sheets_preserved=len(unchanged),legacy_cell_records_unchanged=True,c0_sha256=hashlib.sha256(BASE.read_bytes()).hexdigest(),result_sha256=hashlib.sha256(dest.read_bytes()).hexdigest(),current_inputs_preserved='Verified against C0 before generation; original V06 implementation extended from existing source')
 (OUT/'Preservacao.json').write_text(json.dumps(proof,indent=2));print(dest,proof['result_sha256']);return dest
if __name__=='__main__':
 modules=WORK/'V06_modulos.xlsx';recalc=WORK/'final_recalc/out/V06_modulos.xlsx';cached=WORK/'EFCC_Calculado.xlsx';n=inject_caches(modules,recalc,cached);print('Caches verified',n);assemble(cached)
