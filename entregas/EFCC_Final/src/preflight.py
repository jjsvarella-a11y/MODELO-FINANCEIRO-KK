from pathlib import Path
from zipfile import ZipFile
from lxml import etree as E
import sys,json
sys.path.insert(0,str(Path(__file__).resolve().parents[2]/'V06/src'))
from core import DEFAULTS,legacy
N='http://schemas.openxmlformats.org/spreadsheetml/2006/main';T=lambda n:'{'+N+'}'+n
p=Path('/workspace/kk01_efcc/checkpoints/C0_V06_DASH10.xlsx');diff=[]
with ZipFile(p) as z:
 w=E.fromstring(z.read('xl/workbook.xml'));rels={r.get('Id'):r.get('Target') for r in E.fromstring(z.read('xl/_rels/workbook.xml.rels'))};strings=[''.join(n.itertext()) for n in E.fromstring(z.read('xl/sharedStrings.xml'))]
 def cells(sn):
  s=next(s for s in w.find(T('sheets')) if s.get('name')==sn);part=rels[s.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')];root=E.fromstring(z.read(part.lstrip('/') if part.startswith('/') else 'xl/'+part));out={}
  for c in root.iter(T('c')):
   v=c.findtext(T('v'))
   out[c.get('r')]=strings[int(v)] if c.get('t')=='s' else (float(v) if v is not None and c.get('t') not in ['str','e'] else v)
  return out
 cc=cells('V06 Controlos')
 for i,x in enumerate(DEFAULTS,5):
  if cc.get('B'+str(i))!=x[3]:diff.append([x[0],cc.get('B'+str(i)),x[3]])
 uu=cells('V06 Unidades')
 for r,u in enumerate(legacy.UNITS,5):
  for c,k in [('H','price'),('I','sale'),('L','deed')]:
   if uu.get(c+str(r))!=u[k]:diff.append([u['id'],k,uu.get(c+str(r)),u[k]])
 result={'status':'PASS' if not diff else 'FAIL','source':str(p),'controls_checked':len(DEFAULTS),'unit_price_sale_deed_checked':282,'differences':diff}
 Path('entregas/EFCC_Final/Preservacao_Inputs.json').write_text(json.dumps(result,indent=2));print(result);assert not diff
