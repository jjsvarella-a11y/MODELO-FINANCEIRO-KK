"""Insert DASH10 after the nine dashboards and add navigation without touching the engine."""
from pathlib import Path
from zipfile import ZipFile,ZIP_DEFLATED
from lxml import etree as E
import re,json,hashlib,sys,tempfile
from merge import merge
N='http://schemas.openxmlformats.org/spreadsheetml/2006/main';R='http://schemas.openxmlformats.org/officeDocument/2006/relationships';T=lambda n:'{'+N+'}'+n
DASHES=['EXEC V06','DASH 01 V06','DASH 02 V06','DASH 03 V06','DASH 04A V06','DASH 04B V06','DASH 05 V06','DASH 06 V06','DASH 07 V06']
def paths(z):
 w=E.fromstring(z.read('xl/workbook.xml'));rels={r.get('Id'):r.get('Target') for r in E.fromstring(z.read('xl/_rels/workbook.xml.rels'))};out={}
 for s in w.find(T('sheets')):
  t=rels[s.get('{'+R+'}id')];out[s.get('name')]=t.lstrip('/') if t.startswith('/') else 'xl/'+t
 return w,out
def extend(base,module,output):
 output=Path(output);intermediate=output.with_suffix('.building.xlsx');merge(base,module,intermediate)
 with ZipFile(intermediate) as z,ZipFile(output,'w',ZIP_DEFLATED,compresslevel=9) as zz:
  w,lookup=paths(z);sheetlist=w.find(T('sheets'));report=sheetlist[0];assert report.get('name')=='DASH 10 V06';sheetlist.remove(report);sheetlist.insert(9,report)
  for d in w.findall(T('definedNames')+'/'+T('definedName')):
   if d.get('localSheetId') is not None:
    old=int(d.get('localSheetId'));d.set('localSheetId',str(9 if old==0 else old-1 if old<=9 else old))
  for view in w.find(T('bookViews')):view.set('activeTab','0');view.set('firstSheet','0')
  changes={'xl/workbook.xml':E.tostring(w,encoding='UTF-8',xml_declaration=True)}
  for sn in DASHES+['DASH 10 V06']:
   root=E.fromstring(z.read(lookup[sn]))
   for view in root.findall(T('sheetViews')+'/'+T('sheetView')):view.set('tabSelected','1' if sn=='EXEC V06' else '0')
   if sn in DASHES:
    sd=root.find(T('sheetData'));row=sd.find(T('row')+'[@r="43"]');assert row is not None
    cell=row.find(T('c')+'[@r="B43"]');assert cell is not None and cell.find(T('f')) is None
    cell.attrib.pop('t',None)
    for child in list(cell):cell.remove(child)
    f='HYPERLINK("#\'DASH 10 V06\'!A1",IF(\'V06 Controlos\'!$B$5="PT","10 — Relatório de Viabilidade","10 — Investment Report"))'
    E.SubElement(cell,T('f')).text=f;cell.set('t','str');E.SubElement(cell,T('v')).text='10 — Relatório de Viabilidade'
    merges=root.find(T('mergeCells'));assert merges is not None;assert not any(c.get('ref')=='B43:C45' for c in merges);E.SubElement(merges,T('mergeCell'),ref='B43:C45');merges.set('count',str(len(merges)))
   changes[lookup[sn]]=E.tostring(root,encoding='UTF-8',xml_declaration=True)
  for p in z.infolist():zz.writestr(p.filename,changes.get(p.filename,z.read(p.filename)))
 intermediate.unlink()
 # Compare every existing financial cell, including formula attributes and stored values.
 with ZipFile(base) as a,ZipFile(output) as b:
  _,pa=paths(a);_,pb=paths(b);changed=[];checked=0
  for sn,part in pa.items():
   old=E.fromstring(a.read(part));new=E.fromstring(b.read(pb[sn]));oc={c.get('r'):c for c in old.iter(T('c'))};nc={c.get('r'):c for c in new.iter(T('c'))}
   for addr,c in oc.items():
    if sn in DASHES and addr=='B43':continue
    checked+=1
    def record(c):return (dict(c.attrib),[(E.QName(x).localname,dict(x.attrib),x.text,E.tostring(x).decode().split('>')[1:] if len(x) else None) for x in c])
    # Namespace declarations can move; compare the actual cell contents and attributes.
    if record(c)!=record(nc[addr]):changed.append([sn,addr])
  assert not changed,changed[:10];assert b.testzip() is None
 proof={'base_sha256':hashlib.sha256(Path(base).read_bytes()).hexdigest(),'result_sha256':hashlib.sha256(output.read_bytes()).hexdigest(),'existing_cells_verified':checked,'financial_cells_changed':0,'navigation_only_changes':[sn+'!B43' for sn in DASHES],'report_position':10,'existing_sheet_names_preserved':True}
 output.with_suffix('.preservation.json').write_text(json.dumps(proof,indent=2));print(json.dumps(proof))
if __name__=='__main__':extend(*sys.argv[1:4])
