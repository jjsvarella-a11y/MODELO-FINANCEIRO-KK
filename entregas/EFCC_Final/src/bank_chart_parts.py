"""Give every banking chart its own native ChartPart, as standard Excel writers do.
Copies graphics only, never financial calculations or original Hotel objects.
"""
from bank_r6 import *
from zip_copy import rewrite

def apply(file=None):
 file=Path(file or BM);changes={};clones=0
 with ZipFile(file) as z:
  _,pp=paths(z);rels=E.fromstring(z.read(relpath(pp[BN])));dr=next(x for x in rels if x.get('Type').endswith('/drawing'));dp=resolve(pp[BN],dr.get('Target'));drawing=E.fromstring(z.read(dp));rp=relpath(dp);rr=E.fromstring(z.read(rp));RID='{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id';D='{http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing}';C='{http://schemas.openxmlformats.org/drawingml/2006/chart}';ids={}
  for part in z.namelist():
   if part.startswith('xl/drawings/_rels/') and part.endswith('.rels'):
    owner=part.replace('/_rels/','/')[:-5]
    for r in E.fromstring(z.read(part)):
     if r.get('Type').endswith('/chart'):
      target=resolve(owner,r.get('Target'));ids[target]=ids.get(target,0)+1
  ct=E.fromstring(z.read('[Content_Types].xml'));ctns='{http://schemas.openxmlformats.org/package/2006/content-types}';n=max(int(m[1]) for p in z.namelist() if (m:=re.fullmatch(r'xl/charts/chart(\d+)\.xml',p)))
  for anchor in drawing:
   nv=anchor.find('.//'+D+'cNvPr');c=anchor.find('.//'+C+'chart')
   if nv is None or c is None or not nv.get('name','').startswith('R6 Bank '):continue
   rel=next(r for r in rr if r.get('Id')==c.get(RID));old=resolve(dp,rel.get('Target'))
   if ids.get(old,0)<=1:continue
   n+=1;new=f'xl/charts/chart{n}.xml';changes[new]=z.read(old);rel.set('Target',__import__('posixpath').relpath(new,__import__('posixpath').dirname(dp)));clones+=1;E.SubElement(ct,ctns+'Override',PartName='/'+new,ContentType='application/vnd.openxmlformats-officedocument.drawingml.chart+xml')
   assert relpath(old) not in z.namelist(),'Unexpected banking chart dependencies'
  if clones:changes[rp]=xml(rr);changes['[Content_Types].xml']=xml(ct)
 if changes:
  tmp=file.with_suffix('.chartparts.xlsx');rewrite(file,tmp,changes);tmp.replace(file)
 return clones
if __name__=='__main__':print('Distinct native bank chart parts created',apply())
