"""Read required banking/UI cells; independently scan every worksheet for errors."""
from bank_r6 import *
def read_bank(file):
 needed={'V06 Motor',DAT,SCH,'V06 Cenarios','V06 F01',BN,'EXEC V06',CENTER,'V06 Registo Ativo',IMP,CON,INV,SALES,'DASH 10 V06'}|{v[0] for v in DASHES}
 with ZipFile(file) as z:
  _,pp=paths(z);strings=[''.join(x.itertext()) for x in E.fromstring(z.read('xl/sharedStrings.xml'))] if 'xl/sharedStrings.xml' in z.namelist() else [];out={};errors=[];fc=0
  for sn,p in pp.items():
   data=z.read(p);fc+=data.count(b'<f>')+data.count(b'<f ')
   has_error=b't="e"' in data or b"t='e'" in data
   if sn not in needed and not has_error:continue
   r=E.fromstring(data);cells={}
   for c in r.iter(T('c')):
    typ=c.get('t');v=c.findtext(T('v'))
    if typ=='e':errors.append([sn,c.get('r'),v])
    if sn not in needed:continue
    if typ=='s':v=strings[int(v)] if v is not None else None
    elif typ=='inlineStr':v=''.join(c.find(T('is')).itertext())
    elif typ not in ['str','e'] and v is not None:
     try:v=float(v)
     except ValueError:errors.append([sn,c.get('r'),'invalid numeric '+v])
    cells[c.get('r')]=v
   if sn in needed:out[sn]=cells
 return out,errors,{'formula_count':fc,'sheets':len(pp),'error_scan':'all worksheet parts','numeric_read':'bank engine and public dashboards only'}
