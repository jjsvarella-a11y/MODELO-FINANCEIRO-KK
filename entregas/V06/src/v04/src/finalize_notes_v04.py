"""Presentation-only XML pass, preserves all formulas, values, drawings and links."""
from engine import ROOT
from v04_formula_helpers import note_method
from returns10 import payback_formula
import html
import json,zipfile,xml.etree.ElementTree as E,re,os,hashlib
P=ROOT/'entrega/KK01_Modelo_Integrado_V04.xlsx';N='http://schemas.openxmlformats.org/spreadsheetml/2006/main';T=lambda x:'{'+N+'}'+x;V='urn:schemas-microsoft-com:vml';X='urn:schemas-microsoft-com:office:excel';MC='http://schemas.openxmlformats.org/markup-compatibility/2006'
def clean(t):
 t=re.sub(r'([a-zà-ÿ]{2,})(\d)',r'\1 \2',t);t=re.sub(r'(\d)([A-Za-zÀ-ÿ]{2,})',r'\1 \2',t);t=re.sub(r'([,;])(?=[A-Za-zÀ-ÿ])',r'\1 ',t)
 for word in ['CAPEX','EBITDA','CFADS','GOP','VGV','DSCR','TIR','VPL','EAC','BAC','ETC','KK','MICE','Hotel','Board','Actual Cost','Forecast']:
  t=re.sub(re.escape(word)+r'(?=[a-zà-ÿ])',word+' ',t)
 return t
D=json.loads((ROOT/'src/v04_kpi_dictionary.json').read_text())
for d in D:
 d['method']=note_method(d['key']);d['formula']=re.sub(r'(?<![\w.])MINIFS\(',r'_xlfn.MINIFS(',d['formula'],flags=re.I)
 if d['key']=='investor_payback':d['formula']=payback_formula()
(ROOT/'src/v04_kpi_dictionary.json').write_text(json.dumps(D,ensure_ascii=False,indent=2))
bykey={x['key']:x for x in D};modified=0;dictcount=0;vmlcount=0;before={}
with zipfile.ZipFile(P) as zi,zipfile.ZipFile(P.with_suffix('.notes'),'w',zipfile.ZIP_DEFLATED) as zo:
 wb=E.fromstring(zi.read('xl/workbook.xml'));rs={r.attrib['Id']:r.attrib['Target'] for r in E.fromstring(zi.read('xl/_rels/workbook.xml.rels'))};dictsheet=next(s for s in wb.find(T('sheets')) if s.attrib['name']=='V04 Dicionario KPI');target=rs[dictsheet.attrib['{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id']];dictpath=target.lstrip('/') if target.startswith('/') else 'xl/'+target
 for inf in zi.infolist():
  data=zi.read(inf.filename)
  if inf.filename.startswith('xl/worksheets/') and inf.filename.endswith('.xml'):
   before[inf.filename]=hashlib.sha256(data).hexdigest()
  if inf.filename==dictpath:
   raw=data.decode('utf-8')
   for i,d in enumerate(D,5):
    for cc,content in [('H',d['formula']),('I',d['method'])]:
     addr=cc+str(i);pattern=r'<c\b([^>]*\br="'+addr+r'"[^>]*)>.*?</c>';m=re.search(pattern,raw)
     if m:
      attr=re.sub(r'\s+t="[^"]*"','',m.group(1));raw=raw[:m.start()]+'<c'+attr+' t="inlineStr"><is><t>'+html.escape(content,quote=False)+'</t></is></c>'+raw[m.end():]
   data=raw.encode('utf-8')
  if inf.filename.startswith('xl/comments') and inf.filename.endswith('.xml'):
   x=E.fromstring(data);x.attrib.pop('{'+MC+'}Ignorable',None)
   for c in x.iter(T('comment')):
    for child in list(c):
     if child.tag=='{'+MC+'}AlternateContent':c.remove(child)
    old=''.join(t.text or '' for t in c.iter(T('t')));match=re.search(r'Dicionário: V04 Dicionario KPI / ([^\n]+)',old)
    if not match:continue
    key=re.sub(r'\s+','',match.group(1));d=bykey[key];source=old.split('Motor/fonte: ')[-1].split('\n')[0] if 'Motor/fonte: ' in old else d['source'];lines=[d['title'],'','Definição: '+d['definition'],'','Interpretação: '+d['interpretation'],'','Unidade: '+d['unit']+'. '+d['reference'],'','AUDITORIA','Motor/fonte: '+source,'Método: '+note_method(key),'Dicionário: V04 Dicionario KPI / '+key];content=clean('\n'.join(lines[:-1]))+'\n'+lines[-1];text=c.find(T('text'))
    for child in list(text):text.remove(child)
    for i,line in enumerate(content.split('\n')):
     r=E.SubElement(text,T('r'));rp=E.SubElement(r,T('rPr'));E.SubElement(rp,T('sz'),{'val':'10'});E.SubElement(rp,T('rFont'),{'val':'Arial'})
     if i==0 or line=='AUDITORIA':E.SubElement(rp,T('b'))
     E.SubElement(r,T('t'),{'{http://www.w3.org/XML/1998/namespace}space':'preserve'}).text=line+('\n' if i<len(content.split('\n'))-1 else '')
    for a in list(c):
     if a.tag=='{'+MC+'}AlternateContent':c.remove(a)
    modified+=1
   data=E.tostring(x,encoding='utf-8',xml_declaration=True)
  elif inf.filename.endswith('.vml'):
   x=E.fromstring(data)
   for shape in x.iter('{'+V+'}shape'):
    cd=shape.find('{'+X+'}ClientData')
    if cd is None or cd.attrib.get('ObjectType')!='Note':continue
    style=shape.attrib.get('style','');style=re.sub(r'(width|height|visibility):[^;]+;?','',style);shape.set('style',style.rstrip(';')+';width:375pt;height:315pt;visibility:hidden')
    for ch in list(cd):
     if ch.tag=='{'+X+'}Visible':cd.remove(ch)
    # Preserve the source cell and enlarge the display anchor too.
    anchor=cd.find('{'+X+'}Anchor');row=cd.find('{'+X+'}Row');col=cd.find('{'+X+'}Column')
    if anchor is not None and row is not None and col is not None:
     rr=int(row.text);cc=int(col.text);anchor.text=f'{cc+1}, 15, {rr}, 10, {cc+9}, 15, {rr+21}, 10'
    vmlcount+=1
   data=E.tostring(x,encoding='utf-8',xml_declaration=True)
  zo.writestr(inf,data)
with zipfile.ZipFile(P.with_suffix('.notes')) as z:
 for name,val in before.items():
  if name!=dictpath:assert hashlib.sha256(z.read(name)).hexdigest()==val,name
  else:
   # Dictionary contains explanatory strings only; financial source sheets are byte-identical.
   x=E.fromstring(z.read(name));assert not list(x.iter(T('f')))
os.replace(P.with_suffix('.notes'),P)
(ROOT/'entrega/Verificacao_Notas_V04.json').write_text(json.dumps({'kpi_notes_from_dictionary':modified,'hidden_classic_notes':vmlcount,'formula_values_unchanged':True,'width_pt':375,'height_pt':315},indent=2));print(modified,'KPI notes;',vmlcount,'classic hidden notes; financial formulas/caches preserved')
