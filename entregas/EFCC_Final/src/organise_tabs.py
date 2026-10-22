"""Present only the current executive interface; preserve every worksheet part.
Legacy sheets and internal scenario/helper tabs are hidden, never deleted.
The main V06 Motor remains visible for the ongoing native calculation diagnosis.
"""
from pathlib import Path
from zipfile import ZipFile,ZIP_DEFLATED
from lxml import etree as E
import re,json,hashlib,sys
from assemble import paths,T,xml
ROOT=Path(__file__).resolve().parents[1]
DASHES=['EXEC V06','DASH 01 V06','DASH 02 V06','DASH 03 V06','DASH 04A V06','DASH 04B V06','DASH 05 V06','DASH 06 V06','DASH 07 V06','DASH 10 V06']
HELPERS={'V06 UI Anual','V06 Curva','V06 Traducoes','V06 Ref Unidades','V06 Referencia','V06 Matriz Valores'}
def organise(source):
 source=Path(source);temp=source.with_suffix('.organised.xlsx');hidden_legacy=[];hidden_technical=[];visible=[];preserved=0
 with ZipFile(source) as a:
  w,pp=paths(a);original=E.fromstring(a.read('xl/workbook.xml'));sheets=w.find(T('sheets'))
  before=[(s.get('name'),s.get('sheetId'),s.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')) for s in sheets]
  assert [s.get('name') for s in sheets[:10]]==DASHES
  for s in sheets:
   name=s.get('name');legacy=not(name.startswith('V06 ') or name.endswith(' V06'))
   technical=bool(re.fullmatch(r'V06 [SMOF]\d+',name)) or name.startswith('V06 Calend ') or name in HELPERS
   if legacy or technical:
    s.set('state','hidden');(hidden_legacy if legacy else hidden_technical).append(name)
   else:s.attrib.pop('state',None);visible.append(name)
  for v in w.find(T('bookViews')):v.set('activeTab','0');v.set('firstSheet','0')
  with ZipFile(temp,'w',ZIP_DEFLATED,compresslevel=6) as b:
   for part in a.infolist():b.writestr(part.filename,xml(w) if part.filename=='xl/workbook.xml' else a.read(part.filename))
  with ZipFile(temp) as b:
   assert a.namelist()==b.namelist();assert b.testzip() is None
   for part in a.namelist():
    if part!='xl/workbook.xml':
     assert a.read(part)==b.read(part),part;preserved+=1
   aw=E.fromstring(b.read('xl/workbook.xml'));after=[(s.get('name'),s.get('sheetId'),s.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')) for s in aw.find(T('sheets'))];assert before==after
   # The sole workbook-level modification is sheet visibility.
   for r in [original,aw]:
    for s in r.find(T('sheets')):s.attrib.pop('state',None)
   assert E.tostring(original)==E.tostring(aw)
 assert len(hidden_legacy)==313;assert 'V06 Motor' in visible;assert all(n in visible for n in DASHES)
 temp.replace(source)
 result={'status':'PASS','change':'worksheet visibility only','hidden_legacy_count':len(hidden_legacy),'hidden_legacy_sheets':hidden_legacy,'hidden_technical_count':len(hidden_technical),'hidden_technical_sheets':hidden_technical,'visible_sheets':visible,'parts_preserved_byte_for_byte':preserved,'sheet_names_order_ids_unchanged':True,'financial_formulas_and_values_unchanged':True,'native_excel365':'NOT TESTED','sha256':hashlib.sha256(source.read_bytes()).hexdigest()}
 (ROOT/'Organizacao_Folhas.json').write_text(json.dumps(result,indent=2,ensure_ascii=False));print({k:v for k,v in result.items() if not isinstance(v,list)})
 for name in ['Preservacao.json','Prova_Final_Identidade.json']:
  p=ROOT/name;r=json.loads(p.read_text());r['result_sha256' if name=='Preservacao.json' else 'sha256']=result['sha256'];r['visibility_only_followup']='Organizacao_Folhas.json: all other package parts byte-for-byte preserved';p.write_text(json.dumps(r,indent=2,ensure_ascii=False))
if __name__=='__main__':organise(sys.argv[1])
