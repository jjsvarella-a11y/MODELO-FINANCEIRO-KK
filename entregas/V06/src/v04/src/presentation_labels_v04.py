"""Label-only changes; byte-preserves formulas and cached numeric values."""
from engine import ROOT
import zipfile,xml.etree.ElementTree as E,json,re,html,os
P=ROOT/'entrega/KK01_Modelo_Integrado_V04.xlsx';N='{http://schemas.openxmlformats.org/spreadsheetml/2006/main}';updates={};meta=json.loads((ROOT/'src/v04_risk_metadata.json').read_text())
for item in meta:updates[item['policy']]={'A11':'Valor de saída equity M180','A16':'Ativos líquidos para equity e crédito','A17':'Crédito vendedores recuperável','A18':'Crédito vendedores sem cobertura'}
D=json.loads((ROOT/'src/v04_kpi_dictionary.json').read_text())
def clean(t):
 t=re.sub(r'([a-zà-ÿ]{2,})(\d)',r'\1 \2',str(t));t=re.sub(r'(\d)([A-Za-zÀ-ÿ]{2,})',r'\1 \2',t);t=re.sub(r'([,;])(?=[A-Za-zÀ-ÿ])',r'\1 ',t)
 for word in ['CAPEX','EBITDA','CFADS','GOP','VGV','DSCR','TIR','VPL','EAC','BAC','ETC','KK','Hotel','Board','Actual Cost','Forecast']:
  t=re.sub(re.escape(word)+r'(?=[a-zà-ÿ])',word+' ',t)
 return t
updates['V04 Dicionario KPI']={}
for i,d in enumerate(D,5):
 for col,key in [('B','title'),('C','definition'),('D','interpretation'),('E','unit'),('F','reference'),('I','method')]:d[key]=clean(d[key]);updates['V04 Dicionario KPI'][col+str(i)]=d[key]
(ROOT/'src/v04_kpi_dictionary.json').write_text(json.dumps(D,ensure_ascii=False,indent=2))
updates['EXEC KK01']={a:clean(t) for a,t in [('A6','Capital KK previsto'),('A16','EBITDA ano completo'),('E16','TIR novos efetiva15A'),('A50','Platinum TIR convencional'),('E50','Gold TIR convencional'),('I50','Silver TIR convencional'),('M50','Completion assegurada'),('A55','EAC com Actual Cost documentado'),('I55','Prontidão documental abertura')]}
with zipfile.ZipFile(P) as zi,zipfile.ZipFile(P.with_suffix('.labels'),'w',zipfile.ZIP_DEFLATED) as zo:
 wb=E.fromstring(zi.read('xl/workbook.xml'));rs={r.attrib['Id']:r.attrib['Target'] for r in E.fromstring(zi.read('xl/_rels/workbook.xml.rels'))};paths={}
 for s in wb.find(N+'sheets'):
  if s.attrib['name'] in updates:
   target=rs[s.attrib['{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id']];paths[target.lstrip('/') if target.startswith('/') else 'xl/'+target]=updates[s.attrib['name']]
 for inf in zi.infolist():
  raw=zi.read(inf.filename)
  if inf.filename in paths:
   text=raw.decode('utf-8');before=re.findall(r'<f(?:\s[^>]*)?>.*?</f>|<v>.*?</v>',text)
   for addr,value in paths[inf.filename].items():
    pattern=r'<c\b([^>]*\br="'+addr+r'"[^>]*)>.*?</c>';m=re.search(pattern,text);content='<is><t>'+html.escape(value,quote=False)+'</t></is>'
    if m:
     attr=re.sub(r'\s+t="[^"]*"','',m.group(1));text=text[:m.start()]+'<c'+attr+' t="inlineStr">'+content+'</c>'+text[m.end():]
    else:
     rr=re.search(r'\d+',addr).group();row=re.search(r'(<row\b[^>]*\br="'+rr+r'"[^>]*>)(.*?)(</row>)',text)
     if row:text=text[:row.start(2)]+'<c r="'+addr+'" t="inlineStr">'+content+'</c>'+text[row.start(2):]
   # All targeted labels contain no numeric cached <v>, except shared-string indices;
   # financial formulas themselves must remain byte-identical.
   assert re.findall(r'<f(?:\s[^>]*)?>.*?</f>',text)==re.findall(r'<f(?:\s[^>]*)?>.*?</f>',raw.decode('utf-8'))
   raw=text.encode('utf-8')
  zo.writestr(inf,raw)
os.replace(P.with_suffix('.labels'),P);print('Legendas e dicionário atualizados; fórmulas preservadas.')
