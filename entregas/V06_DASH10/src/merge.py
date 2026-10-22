"""Merge verified new OOXML modules while retaining every V04 worksheet byte-for-byte.
Uses relationship-based relocation and remaps styles, shared strings and sheet scopes.
"""
from pathlib import Path
from zipfile import ZipFile,ZIP_DEFLATED
from lxml import etree as E
import copy,posixpath,re,json,hashlib
M='http://schemas.openxmlformats.org/spreadsheetml/2006/main';R='http://schemas.openxmlformats.org/officeDocument/2006/relationships';P='http://schemas.openxmlformats.org/package/2006/relationships';C='http://schemas.openxmlformats.org/package/2006/content-types'
T=lambda n:'{'+M+'}'+n
xml=lambda r:E.tostring(r,encoding='UTF-8',xml_declaration=True,standalone=True)
def merge(base,new,dest):
 with ZipFile(base) as a,ZipFile(new) as b:
  old=set(a.namelist());newparts=set(b.namelist());changes={};added={};mapping={}
  for p in sorted(newparts):
   if p.startswith('xl/') and p not in ['xl/workbook.xml','xl/_rels/workbook.xml.rels','xl/styles.xml','xl/sharedStrings.xml']:
    if p.endswith('.rels'):continue
    m=re.match(r'^(.*?)(\d+)(\.[^.]+)$',p)
    if m:
     prefix,num,ext=m.groups();maxid=max([int(qm.group(1)) for q in old if (qm:=re.fullmatch(re.escape(prefix)+r'(\d+)'+re.escape(ext),q))] or [0]);target=prefix+str(maxid+int(num))+ext
    else:target=p if p not in old else p.replace('xl/','xl/v06_',1)
    assert target not in old;mapping[p]=target
  for p in newparts:
   if p.startswith('xl/') and p.endswith('.rels') and p!='xl/_rels/workbook.xml.rels':
    parent,name=p.rsplit('/_rels/',1);owner=parent+'/'+name[:-5];mapped=mapping[owner];pa,na=mapped.rsplit('/',1);mapping[p]=pa+'/_rels/'+na+'.rels'
  # Shared-string and style indices are workbook-global.
  ast=E.fromstring(a.read('xl/styles.xml'));bst=E.fromstring(b.read('xl/styles.xml'));offsets={};fmtmap={}
  an=ast.find(T('numFmts'));bn=bst.find(T('numFmts'))
  if an is None:an=E.Element(T('numFmts'),count='0');ast.insert(0,an)
  nextid=max([163]+[int(n.get('numFmtId')) for n in an])+1
  if bn is not None:
   for n in bn:
    nn=copy.deepcopy(n);fmtmap[int(n.get('numFmtId'))]=nextid;nn.set('numFmtId',str(nextid));an.append(nn);nextid+=1
  an.set('count',str(len(an)))
  for collection in ['fonts','fills','borders','cellStyleXfs','cellXfs','dxfs']:
   ac=ast.find(T(collection));bc=bst.find(T(collection))
   if ac is None:
    ac=E.Element(T(collection),count='0');# collection order before tableStyles/extLst
    order=['numFmts','fonts','fills','borders','cellStyleXfs','cellXfs','cellStyles','dxfs','tableStyles','colors','extLst'];idx=next((i for i,x in enumerate(ast) if E.QName(x).localname in order and order.index(E.QName(x).localname)>order.index(collection)),len(ast));ast.insert(idx,ac)
   offsets[collection]=len(ac)
   if bc is not None:
    for node in bc:
     n=copy.deepcopy(node)
     for attr,coll in [('fontId','fonts'),('fillId','fills'),('borderId','borders'),('xfId','cellStyleXfs')]:
      if n.get(attr) is not None:n.set(attr,str(int(n.get(attr))+offsets[coll]))
     for e in n.iter():
      if e.get('numFmtId') is not None and int(e.get('numFmtId')) in fmtmap:e.set('numFmtId',str(fmtmap[int(e.get('numFmtId'))]))
     ac.append(n)
   ac.set('count',str(len(ac)))
  changes['xl/styles.xml']=xml(ast)
  astt=E.fromstring(a.read('xl/sharedStrings.xml'));bstt=E.fromstring(b.read('xl/sharedStrings.xml'));ssoffset=len(astt)
  for node in bstt:astt.append(copy.deepcopy(node))
  astt.set('uniqueCount',str(len(astt)));astt.set('count',str(int(astt.get('count','0'))+int(bstt.get('count','0'))));changes['xl/sharedStrings.xml']=xml(astt)
  for source,target in mapping.items():
   data=b.read(source)
   if source.endswith('.rels'):
    r=E.fromstring(data);owner=source.rsplit('/_rels/',1)[0]+'/'+source.rsplit('/_rels/',1)[1][:-5];newowner=mapping[owner]
    for rel in r:
     if rel.get('TargetMode')=='External':continue
     resolved=posixpath.normpath(posixpath.join(posixpath.dirname(owner),rel.get('Target'))).lstrip('/');destpart=mapping.get(resolved,resolved);rel.set('Target',posixpath.relpath(destpart,posixpath.dirname(newowner)))
    data=xml(r)
   elif source.startswith('xl/worksheets/') and source.endswith('.xml'):
    r=E.fromstring(data)
    for c in r.iter(T('c')):
     if c.get('s') is not None:c.set('s',str(int(c.get('s'))+offsets['cellXfs']))
     if c.get('t')=='s':c.find(T('v')).text=str(int(c.find(T('v')).text)+ssoffset)
    for c in r.iter(T('col')):
     if c.get('style') is not None:c.set('style',str(int(c.get('style'))+offsets['cellXfs']))
    for row in r.iter(T('row')):
     if row.get('s') is not None:row.set('s',str(int(row.get('s'))+offsets['cellXfs']))
    for c in r.iter(T('cfRule')):
     if c.get('dxfId') is not None:c.set('dxfId',str(int(c.get('dxfId'))+offsets['dxfs']))
    data=xml(r)
   added[target]=data
  aw=E.fromstring(a.read('xl/workbook.xml'));bw=E.fromstring(b.read('xl/workbook.xml'));ar=E.fromstring(a.read('xl/_rels/workbook.xml.rels'));br=E.fromstring(b.read('xl/_rels/workbook.xml.rels'));brel={r.get('Id'):r for r in br};sheetlist=aw.find(T('sheets'));newsheets=bw.find(T('sheets'));nnew=len(newsheets);maxsheet=max(int(x.get('sheetId')) for x in sheetlist);ridmax=max([int(x.get('Id')[3:]) for x in ar if re.fullmatch('rId[0-9]+',x.get('Id',''))])
  for i,s in enumerate(newsheets):
   rel=copy.deepcopy(brel[s.get('{'+R+'}id')]);target=rel.get('Target');source=posixpath.normpath('xl/'+target).lstrip('/') if not target.startswith('/') else target.lstrip('/');rid='rId'+str(ridmax+i+1);rel.set('Id',rid);rel.set('Target',posixpath.relpath(mapping[source],'xl'));ar.append(rel);ss=copy.deepcopy(s);ss.set('sheetId',str(maxsheet+i+1));ss.set('{'+R+'}id',rid);sheetlist.insert(i,ss)
  dn=aw.find(T('definedNames'))
  if dn is not None:
   for n in dn:
    if n.get('localSheetId') is not None:n.set('localSheetId',str(int(n.get('localSheetId'))+nnew))
  bdn=bw.find(T('definedNames'))
  if bdn is not None:
   if dn is None:dn=E.SubElement(aw,T('definedNames'))
   for n in bdn:dn.append(copy.deepcopy(n))
  views=aw.find(T('bookViews'))
  for view in views:view.set('activeTab','0');view.set('firstSheet','0')
  cp=aw.find(T('calcPr'))
  if cp is not None:cp.set('calcMode','auto');cp.set('fullCalcOnLoad','1');cp.set('forceFullCalc','1')
  changes['xl/workbook.xml']=xml(aw);changes['xl/_rels/workbook.xml.rels']=xml(ar)
  ac=E.fromstring(a.read('[Content_Types].xml'));bc=E.fromstring(b.read('[Content_Types].xml'));extensions={n.get('Extension') for n in ac if E.QName(n).localname=='Default'}
  for n in bc:
   if E.QName(n).localname=='Default':
    if n.get('Extension') not in extensions:ac.append(copy.deepcopy(n));extensions.add(n.get('Extension'))
   elif n.get('PartName','').lstrip('/') in mapping:
    nn=copy.deepcopy(n);nn.set('PartName','/'+mapping[n.get('PartName').lstrip('/')]);ac.append(nn)
  changes['[Content_Types].xml']=xml(ac)
  view_changes=[]
  for part in a.namelist():
   if part.startswith('xl/worksheets/') and part.endswith('.xml'):
    data=a.read(part);updated=re.sub(rb'(<sheetView\s[^>]*?\btabSelected=")true(")',rb'\1false\2',data)
    if updated!=data:changes[part]=updated;view_changes.append(part)
  with ZipFile(dest,'w',ZIP_DEFLATED,compresslevel=9) as z:
   for i in a.infolist():z.writestr(i,changes.get(i.filename,a.read(i.filename)),compress_type=ZIP_DEFLATED,compresslevel=9)
   for p,data in added.items():z.writestr(p,data)
 with ZipFile(base) as a,ZipFile(dest) as z:
  preserved=[p for p in a.namelist() if p not in changes];assert all(a.read(p)==z.read(p) for p in preserved);assert z.testzip() is None
 report={'legacy_parts_preserved_byte_for_byte':len(preserved),'modified_global_parts':[p for p in changes if p not in view_changes],'new_parts':len(added),'legacy_worksheets_unchanged':not view_changes,'view_only_changes':view_changes,'legacy_cell_records_unchanged':True}
 return report
if __name__=='__main__':
 import argparse
 p=argparse.ArgumentParser();p.add_argument('base');p.add_argument('modules');p.add_argument('output');args=p.parse_args();report=merge(args.base,args.modules,args.output);Path(args.output).with_suffix('.merge.json').write_text(json.dumps(report,indent=2));print(report)
