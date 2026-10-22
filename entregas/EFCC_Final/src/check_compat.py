from pathlib import Path
from zipfile import ZipFile
from lxml import etree as E
import json,posixpath
N='http://schemas.openxmlformats.org/spreadsheetml/2006/main';T=lambda n:'{'+N+'}'+n
p=Path('/workspace/kk01_efcc/compat_recalc/out/C1_Compat_Modules.xlsx');ref=json.loads(Path('/workspace/kk01_efcc/checkpoints/C1_Referencia.json').read_text());errors=[];diff=[];n=0
with ZipFile(p) as z:
 w=E.fromstring(z.read('xl/workbook.xml'));rr={r.get('Id'):r.get('Target') for r in E.fromstring(z.read('xl/_rels/workbook.xml.rels'))};ss=E.fromstring(z.read('xl/sharedStrings.xml'));strings=[''.join(x.itertext()) for x in ss]
 for s in w.find(T('sheets')):
  sn=s.get('name');target=rr[s.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')];part=target.lstrip('/') if target.startswith('/') else posixpath.normpath('xl/'+target);root=E.fromstring(z.read(part))
  for c in root.iter(T('c')):
   if c.get('t')=='e':errors.append([sn,c.get('r'),c.findtext(T('v'))])
  if sn=='V06 Motor' or sn.startswith('V06 S') and sn[5:].isdigit():
   idx=0 if sn=='V06 Motor' else int(sn[5:]);label=list(ref['scenarios'])[idx];colkeys={i:k for k,i in ref['main_columns'].items()}
   for row in root.find(T('sheetData')):
    r=int(row.get('r'));m=r-5
    if not 0<=m<180:continue
    for c in row:
     addr=c.get('r');letters=''.join(filter(str.isalpha,addr));ci=0
     for a in letters:ci=ci*26+ord(a)-64
     ci-=1
     if ci not in colkeys:continue
     k=colkeys[ci];v=c.findtext(T('v'));expected=ref['scenarios'][label][m][k]
     if not isinstance(expected,(float,int)):continue
     n+=1
     if v is None or c.get('t')=='e' or abs(float(v)-expected)>max(.00001,abs(expected)*1e-10):diff.append([sn,addr,k,expected,v])
result={'status':'PASS' if not errors and not diff else 'FAIL','recalculation':'LibreOffice, all formula caches invalidated before calculation','excel_native':'NOT TESTED','monthly_numeric_comparisons':n,'errors':errors,'differences':diff,'xlsx':str(p)}
Path('entregas/EFCC_Final/Compat_Recalculo.json').write_text(json.dumps(result,indent=2));print({k:(len(v) if isinstance(v,list) else v) for k,v in result.items()});assert not errors and not diff,(errors[:3],diff[:3])
