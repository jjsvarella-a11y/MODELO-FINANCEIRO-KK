"""Preserve the active sheet but clear other selected tabs; financial cells unchanged."""
from pathlib import Path
from zipfile import ZipFile
from lxml import etree as E
from assemble import paths,T,xml
from zip_copy import rewrite
import hashlib,json,argparse

def ungroup(source,target):
 changes={};before=[];after=[];proof=[]
 with ZipFile(source) as z:
  wb,parts=paths(z);nodes=list(wb.find(T('sheets')));views=wb.findall('.//'+T('workbookView'));idx=int(views[0].get('activeTab','0')) if views else 0
  if idx>=len(nodes):raise ValueError('Invalid active sheet index')
  active=nodes[idx].get('name')
  if nodes[idx].get('state','visible')!='visible':raise ValueError('Active sheet is hidden')
  for sn,path in parts.items():
   root=E.fromstring(z.read(path));sheetviews=root.findall('.//'+T('sheetView'));dirty=False
   original=E.tostring(root.find(T('sheetData')),method='c14n')
   for view in sheetviews:
    if view.get('tabSelected')=='1':before.append(sn)
    selected='1' if sn==active else '0'
    if view.get('tabSelected','0')!=selected:view.set('tabSelected',selected);dirty=True
    if selected=='1':after.append(sn)
   if dirty:changes[path]=xml(root)
   assert E.tostring(root.find(T('sheetData')),method='c14n')==original
   proof.append(hashlib.sha256(original).hexdigest())
 rewrite(source,target,changes)
 return {'active_sheet':active,'selected_before':before,'selected_after':after,'worksheet_cell_records_unchanged':len(proof),'financial_formula_changes':0}
if __name__=='__main__':
 p=argparse.ArgumentParser();p.add_argument('source',type=Path);p.add_argument('target',type=Path);a=p.parse_args();print(json.dumps(ungroup(a.source,a.target),indent=2))
