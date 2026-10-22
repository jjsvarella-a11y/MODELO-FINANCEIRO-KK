"""Fix Excel2019 MINIFS namespace and undefined IRR deviations; never fixes cached results."""
from engine import ROOT
import zipfile,xml.etree.ElementTree as E,json,re,os
P=ROOT/'entrega/KK01_Modelo_Integrado_V04.xlsx';M=json.loads((ROOT/'src/build_metadata.json').read_text());etc="'Pressupostos'!$B$"+str(M['param_rows']['etc_amount']);N='http://schemas.openxmlformats.org/spreadsheetml/2006/main';T=lambda n:'{'+N+'}'+n;count=0
with zipfile.ZipFile(P) as zi,zipfile.ZipFile(P.with_suffix('.compat'),'w',zipfile.ZIP_DEFLATED) as zo:
 wb=E.fromstring(zi.read('xl/workbook.xml'));rels={x.attrib['Id']:x.attrib['Target'] for x in E.fromstring(zi.read('xl/_rels/workbook.xml.rels'))};names={}
 for sn in wb.find(T('sheets')):
  target=rels[sn.attrib['{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id']];names[target.lstrip('/') if target.startswith('/') else 'xl/'+target]=sn.attrib['name']
 for inf in zi.infolist():
  data=zi.read(inf.filename)
  if inf.filename in names:
   x=E.fromstring(data);changed=False;sn=names[inf.filename]
   for c in x.iter(T('c')):
    f=c.find(T('f'))
    if f is None or f.text is None:continue
    old=f.text;new=re.sub(r'(?<![\w.])MINIFS\(',r'_xlfn.MINIFS(',old,flags=re.I)
    if (sn=='V04 Retornos' or re.fullmatch(r'V04 R\d+Retornos',sn)) and c.attrib['r'] in ['E5','E6','E7']:
     r=c.attrib['r'][1:];new=f'IF(ISNUMBER(D{r}),D{r}-C{r},"Indisponível")'
    if sn=='V04 R06CAPEX' and f'ISNUMBER({etc}),{etc}*' not in new:new=new.replace(etc,f'IF(ISNUMBER({etc}),{etc}*1.1,"")')
    if new!=old:
     f.text=new;changed=True;count+=1
     v=c.find(T('v'))
     if v is not None:c.remove(v)
   if changed:data=E.tostring(x,encoding='utf-8',xml_declaration=True)
  zo.writestr(inf,data)
os.replace(P.with_suffix('.compat'),P);print(count,'formulas corrected; full external recalc required')
