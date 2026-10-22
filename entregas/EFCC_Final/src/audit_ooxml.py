"""Structural checks relevant to repair risks; not an Open XML SDK certificate."""
from zipfile import ZipFile
from pathlib import Path
from lxml import etree as E
import json,posixpath,re,sys
from repair_empty_containers import empty_containers
ROOT=Path(__file__).resolve().parents[1];N='http://schemas.openxmlformats.org/spreadsheetml/2006/main';T=lambda n:'{'+N+'}'+n

def audit(path):
 failures=[];counts={'parts':0,'relationships':0,'worksheets':0,'cells':0,'formulas':0,'charts':0};long=[]
 with ZipFile(path) as z:
  names=z.namelist();parts=set(names)
  if len(parts)!=len(names):failures.append('Duplicate ZIP part names')
  if z.testzip() is not None:failures.append('ZIP CRC failure')
  styles=E.fromstring(z.read('xl/styles.xml'));nstyles=len(styles.find(T('cellXfs')));strings=E.fromstring(z.read('xl/sharedStrings.xml'));nstrings=len(strings);wb=E.fromstring(z.read('xl/workbook.xml'));sheets=wb.find(T('sheets'));sheetnames=[s.get('name') for s in sheets]
  for index,font in enumerate(styles.find(T('fonts'))):
   bold=font.find(T('b'))
   if bold is not None and font.index(bold)>0:failures.append(['xl/styles.xml','Bold font element must precede size/color/name/family',index])
  if len(sheetnames)!=len(set(n.casefold() for n in sheetnames)):failures.append('Duplicate case-insensitive sheet names')
  for name in names:
   counts['parts']+=1
   if not (name.endswith('.xml') or name.endswith('.rels')):continue
   try:r=E.fromstring(z.read(name))
   except Exception as ex:failures.append([name,'XML parse',str(ex)]);continue
   if name.endswith('.rels'):
    if name=='_rels/.rels':owner=''
    else:d,n=name.rsplit('/_rels/',1);owner=d+'/'+n[:-5]
    ids=[]
    for rel in r:
     ids.append(rel.get('Id'));counts['relationships']+=1
     if rel.get('TargetMode')=='External':continue
     target=rel.get('Target')
     if rel.get('Type','').endswith('/hyperlink') and target.startswith('#'):
      # Workbook-location fragments are hyperlink destinations, not ZIP parts.
      dest=target[1:].split('!',1)[0].strip("'")
      if dest not in sheetnames:failures.append([name,'Missing internal hyperlink sheet',dest])
      continue
     resolved=target.lstrip('/') if target.startswith('/') else posixpath.normpath(posixpath.join(posixpath.dirname(owner),target))
     if resolved not in parts:failures.append([name,'Missing target',resolved])
    if len(ids)!=len(set(ids)):failures.append([name,'Duplicate relationship IDs'])
   elif name.startswith('xl/worksheets/') and '/_rels/' not in name:
    counts['worksheets']+=1;seen=set()
    for invalid in empty_containers(r):failures.append([name,'Optional container missing mandatory child records',E.QName(invalid).localname])
    for row in r.findall(T('sheetData')+'/'+T('row')):
     lastci=-1
     for c in row:
      if c.tag!=T('c'):continue
      addr=c.get('r');counts['cells']+=1
      if addr in seen:failures.append([name,'Duplicate cell',addr])
      seen.add(addr)
      if int(re.search(r'\d+',addr)[0])!=int(row.get('r')):failures.append([name,'Row mismatch',addr])
      ci=0
      for a in re.match('[A-Z]+',addr)[0]:ci=ci*26+ord(a)-64
      if ci<=lastci:failures.append([name,'Cells not in increasing order',addr])
      lastci=ci
      if not 0<=int(c.get('s','0'))<nstyles:failures.append([name,'Invalid style',addr])
      typ=c.get('t');v=c.findtext(T('v'));f=c.find(T('f'))
      if typ=='s' and (v is None or not 0<=int(v)<nstrings):failures.append([name,'Invalid shared string',addr,v])
      if typ=='e':failures.append([name,'Cached cell error',addr,v])
      if f is not None:
       counts['formulas']+=1
       if len(f.text or '')>8192:long.append([name,addr,len(f.text)])
       if f.text and '#REF!' in f.text:failures.append([name,'Broken formula reference',addr])
      if typ in [None,'n'] and v not in [None,'']:
       try:float(v)
       except ValueError:failures.append([name,'Non-numeric value with numeric type',addr,v])
    for merge in r.findall(T('mergeCells')+'/'+T('mergeCell')):
     if ':' not in merge.get('ref'):failures.append([name,'Invalid merge',merge.get('ref')])
   elif name.startswith('xl/charts/') and '/_rels/' not in name:counts['charts']+=1
  for n in wb.findall(T('definedNames')+'/'+T('definedName')):
   if n.get('localSheetId') is not None and not 0<=int(n.get('localSheetId'))<len(sheets):failures.append(['Invalid localSheetId',n.attrib])
   if n.text and '#REF!' in n.text:failures.append(['Broken defined name',n.attrib,n.text])
  # Workbook-level names cannot resolve to deleted/new duplicate identities.
  dids=[s.get('sheetId') for s in sheets]
  if len(dids)!=len(set(dids)):failures.append('Duplicate sheetId')
 result={'status':'PASS' if not failures and not long else 'FAIL','scope':'ZIP CRC, XML parsing, relationship targets, style/shared-string indices, numeric types, cell order, required child cardinality for worksheet containers, named scopes, formula length/references','native_excel365':'NOT TESTED','counts':counts,'failures':failures,'formulas_over_8192':long}
 (ROOT/'Auditoria_OOXML.json').write_text(json.dumps(result,indent=2,ensure_ascii=False));print({'status':result['status'],**counts,'failures':len(failures),'long_formulas':len(long)});assert not failures and not long,(failures[:5],long[:5]);return result
if __name__=='__main__':audit(Path(sys.argv[1]))
