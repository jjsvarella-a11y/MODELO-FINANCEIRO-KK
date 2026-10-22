from engine import *
import zipfile,xml.etree.ElementTree as E,re,openpyxl,hashlib
P=ROOT/'entrega/KK01_Modelo_Integrado_V04.xlsx';N='http://schemas.openxmlformats.org/spreadsheetml/2006/main';T=lambda n:'{'+N+'}'+n;R='http://schemas.openxmlformats.org/officeDocument/2006/relationships';out=[]
def check(label,condition,detail=None):out.append({'check':label,'ok':bool(condition),'detail':detail})
with zipfile.ZipFile(P) as z:
 check('ZIPintegridadeCRC',z.testzip() is None);wb=E.fromstring(z.read('xl/workbook.xml'));sheets=list(wb.find(T('sheets')));names=[s.attrib['name'] for s in sheets];rels={x.attrib['Id']:x.attrib['Target'] for x in E.fromstring(z.read('xl/_rels/workbook.xml.rels'))};paths={s.attrib['name']:(rels[s.attrib['{'+R+'}id']].lstrip('/') if rels[s.attrib['{'+R+'}id']].startswith('/') else 'xl/'+rels[s.attrib['{'+R+'}id']]) for s in sheets}
 check('313folhasunicaseoriginais',len(names)==313 and len(set(names))==313,len(names));check('InicialEXEC',names[int(wb.find(T('bookViews'))[0].attrib.get('activeTab',0))]=='EXEC KK01');check('Semmacros',not any('vbaProject' in n for n in z.namelist()));check('Semexternalworkbooks',not any(n.startswith('xl/externalLinks/') for n in z.namelist()));check('SemSilverPlusativo',not any('Silver Plus' in n for n in names))
 missing=[];broken=[];errors=[];fc=0;links=[];validations=0;filters=0
 for name,path in paths.items():
  x=E.fromstring(z.read(path));validations+=len(x.findall('.//'+T('dataValidation')));filters+=len(x.findall(T('autoFilter')))
  for c in x.iter(T('c')):
   f=c.find(T('f'));val=c.find(T('v'))
   if f is not None:
    fc+=1
    if '#REF!' in (f.text or ''):broken.append((name,c.attrib['r']))
    if val is None:missing.append((name,c.attrib['r']))
   if c.attrib.get('t')=='e':errors.append((name,c.attrib['r'],val.text if val is not None else None))
  for l in x.iter(T('hyperlink')):
   location=l.attrib.get('location','')
   if '!' in location:
    sn=location.split('!')[0].strip("'").replace("''", "'")
    if sn not in names:links.append((name,location))
 check('Formulasrefintegra',not broken,broken[:5]);check('Zerocodigoserro',not errors,errors[:5]);check('Cachescompletas',not missing,missing[:5]);check('Formulasdinamicas',fc>500000,fc);check('Hiperligacoesinternas',not links,links[:5]);check('Validacoespreservadas',validations>100,validations);check('Filtrosdetalhe',filters>=10,filters)
 ex=E.fromstring(z.read(paths['EXEC KK01']));drawing=ex.find(T('drawing'));rp=paths['EXEC KK01'].replace('worksheets/','worksheets/_rels/')+'.rels';drrels={r.attrib['Id']:r.attrib['Target'] for r in E.fromstring(z.read(rp))};target=drrels[drawing.attrib['{'+R+'}id']];import posixpath
 dp=posixpath.normpath(posixpath.join(posixpath.dirname(paths['EXEC KK01']),target));dp=dp.lstrip('/') if dp.startswith('/') else dp;dx=E.fromstring(z.read(dp));pics=dx.findall('.//{http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing}pic');clicks=dx.findall('.//{http://schemas.openxmlformats.org/drawingml/2006/main}hlinkClick');check('8iconescomlinks',len(pics)==8 and len(clicks)>=8,{'pictures':len(pics),'clicks':len(clicks)})
 notes=[];bad=[]
 for name in z.namelist():
  if name.startswith('xl/comments') and name.endswith('.xml'):
   for c in E.fromstring(z.read(name)).iter(T('comment')):
    text=''.join(x.text or '' for x in c.iter(T('t')))
    if 'Dicionário: V04 Dicionario KPI / ' in text:
     notes.append(text)
     if not all(token in text for token in ['Definição:','Interpretação:','Unidade:','AUDITORIA','Motor/fonte:','Método:']):bad.append(c.attrib['ref'])
 check('NotasKPIcontextuais',len(notes)>=90 and not bad,{'count':len(notes),'bad':bad});shapes=[]
 for name in z.namelist():
  if name.endswith('.vml'):
   for shape in E.fromstring(z.read(name)).iter('{urn:schemas-microsoft-com:vml}shape'):
    cd=shape.find('{urn:schemas-microsoft-com:office:excel}ClientData')
    if cd is not None and cd.attrib.get('ObjectType')=='Note':shapes.append(shape.attrib.get('style',''))
 check('Notasclassicasocultaselegiveis',all('visibility:hidden' in s and 'width:375pt' in s and 'height:315pt' in s for s in shapes),len(shapes));charts=[n for n in z.namelist() if re.fullmatch(r'xl/charts/chart\d+.xml',n)];check('Graficosnativoseditaveis',len(charts)>=10,len(charts));wf=E.fromstring(z.read(paths['V04 Master Workflow']));check('Workflowexpansivonativo',any(int(row.attrib.get('outlineLevel',0))>0 for row in wf.iter(T('row'))))
w=openpyxl.load_workbook(P,read_only=True,data_only=True);meta=json.loads((ROOT/'src/build_metadata.json').read_text());param={c.coordinate:c.value for row in w['Pressupostos'].iter_rows(max_col=2) for c in row if hasattr(c,'coordinate')};check('Linhaestrategica30M',param['B'+str(meta['param_rows']['limit'])]==30000000);check('94fracoesunicaseVGV',len(set(row[0] for row in w['Unidades'].iter_rows(min_row=5,max_row=98,values_only=True)))==94 and abs(sum(row[7] for row in w['Unidades'].iter_rows(min_row=5,max_row=98,values_only=True))-SOURCE['vgv_cents'])<.01)
mat=list(w['Matriz Societaria V03'].iter_rows(min_row=5,max_row=9,max_col=9,values_only=True));check('Matriz20blocose100pct',mat[0][1:5]==(1,2,0,3) and mat[1][1:5]==(0,0,5,5) and mat[2][1:5]==(1,1,10,12) and mat[4][1:7]==(2,3,15,20,1,100000))
unitdata=list(w['Unidades'].iter_rows(min_row=5,max_row=98,values_only=True));rc=list(w['Recebimentos'].iter_rows(min_row=5,max_row=98,max_col=181,values_only=True));independent=receipts(config())['units'];receipt_fail=[]
for i,(unit,row,exp) in enumerate(zip(unitdata,rc,independent)):
 if abs(sum(row[1:181])-unit[7])>.01:receipt_fail.append((row[0],'100pct'))
 for m,(actual,expected) in enumerate(zip(row[1:181],exp),1):
  if abs((actual or 0)-expected)>.01:receipt_fail.append((row[0],m))
check('16920recebimentosindividuaisreconciliados',not receipt_fail,{'comparisons':16920,'failures':receipt_fail[:6]})
rvals={c.coordinate:c.value for row in w['V04 Retornos'].iter_rows(max_row=60,max_col=11) for c in row if hasattr(c,'coordinate')};check('Metas10_6_55convencionais',all(abs(rvals['D'+str(i)]-t)<1e-8 for i,t in [(5,.1),(6,.06),(7,.055)]));check('Distribuicoes100pct',abs(rvals['B33'])<.01);check('Coberturaasseguradaindisponivel',w['V04 Completion']['C17'].value=='Indisponível');check('Dadospipelineefetivosvazios',all(row[i] is None for row in w['V04 Pipeline'].iter_rows(min_row=5,max_row=16,values_only=True) for i in [3,4,10,11,12,13,14,15,16,17]));check('TIR10Aemdashboardselecionada',abs(w['DASH Acionistas']['B8'].value-rvals['B55'])<1e-10);w.close();check('Ficheiromenos40MB',P.stat().st_size<40000000,P.stat().st_size)
check('V03originalpreservada',hashlib.sha256((ROOT/'original/KK01_Modelo_Integrado_V03.xlsx').read_bytes()).hexdigest()==json.loads((ROOT/'original/Manifesto_Base_V03.json').read_text())['sha256'])
report={'passed':sum(x['ok'] for x in out),'failed':sum(not x['ok'] for x in out),'checks':out};(ROOT/'entrega/Testes_Estruturais_V04.json').write_text(json.dumps(report,ensure_ascii=False,indent=2));print(json.dumps(report,ensure_ascii=False,indent=2));raise SystemExit(report['failed']>0)
