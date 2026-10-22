from pathlib import Path
from zipfile import ZipFile,ZIP_DEFLATED
from lxml import etree as E
import xlsxwriter,copy,json,hashlib,sys
root=Path('/workspace/MODELO-FINANCEIRO-KK/entregas/EFCC_Final');out=root/'Premium_Checkpoint';src=out/'KK01_Executive_Financial_Control_Center_CHECKPOINT_R3.xlsx';dest=out/'KK01_Executive_Financial_Control_Center_CHECKPOINT_R4.xlsx'
C='{http://schemas.openxmlformats.org/drawingml/2006/chart}';D='{http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing}';A='{http://schemas.openxmlformats.org/drawingml/2006/main}'
if len(sys.argv)>2:src,dest=Path(sys.argv[1]),Path(sys.argv[2])
if dest.exists():raise FileExistsError(dest)
b=xlsxwriter.Workbook('/tmp/Clean_Hotel_Charts.xlsx');v=b.add_worksheet('V06 Visual');s=b.add_worksheet('Hotel');seriesproof=[]
with ZipFile(src) as z:
 for i,num in enumerate([42,44,45]):
  old=E.fromstring(z.read(f'xl/charts/chart{num}.xml'));ch=b.add_chart({'type':'column','subtype':'stacked'} if i==0 else {'type':'doughnut'} if i==1 else {'type':'line'})
  for k,ser in enumerate(old.findall('.//'+C+'ser')):
   cats=ser.find(C+'cat');vals=ser.find(C+'val');cf=cats.find('.//'+C+'f').text;vf=vals.find('.//'+C+'f').text
   cat_values=[p.findtext(C+'v') for p in cats.findall('.//'+C+'pt')];values=[float(p.findtext(C+'v')) for p in vals.findall('.//'+C+'pt')];n=ser.find(C+'tx');nf=n.find('.//'+C+'f').text if n is not None and n.find('.//'+C+'f') is not None else None
   spec={'name':'='+nf if nf else n.findtext(C+'v') if n is not None and n.find(C+'v') is not None else 'Receitas / Revenue','categories':'='+cf,'values':'='+vf,'categories_data':cat_values,'values_data':values}
   if i==2 and k==1:spec['y2_axis']=True
   if i==1:spec['points']=[{'fill':{'color':x}} for x in ['142337','1765A8','C8101E','687586']]
   else:spec['line']={'color':['142337','C8101E','1765A8','687586'][k%4]}
   ch.add_series(spec);seriesproof.append({'chart':num,'categories':cf,'values':vf,'series_name_formula':nf,'points':len(values)})
  ch.set_legend({'position':'bottom','font':{'size':9}});ch.set_chartarea({'border':{'none':True},'fill':{'color':'FFFFFF'}});ch.set_size({'width':400,'height':187})
  if i!=1:ch.set_y_axis({'num_format':'0,," M€"','num_font':{'size':9}})
  if i==2:ch.set_y2_axis({'num_format':'0%','num_font':{'size':9}})
  s.insert_chart(57,4+i*7,ch)
 b.close()
 with ZipFile('/tmp/Clean_Hotel_Charts.xlsx') as assets:
  dr=E.fromstring(assets.read('xl/drawings/drawing1.xml'))
  for i,anchor in enumerate(dr):
   anchor.set('editAs','twoCell')
   for where,col,row in [('from',4+i*7,57),('to',11+i*7,89)]:
    node=anchor.find(D+where);node.find(D+'col').text=str(col);node.find(D+'row').text=str(row);node.find(D+'colOff').text='0';node.find(D+'rowOff').text='0'
   frame=anchor.find(D+'graphicFrame');xf=frame.find(D+'xfrm');xf.find(A+'off').set('x',str((4+i*7)*600000));xf.find(A+'off').set('y','6000000');xf.find(A+'ext').set('cx','3800000');xf.find(A+'ext').set('cy','1781175')
  changes={'xl/drawings/drawing7.xml':E.tostring(dr,xml_declaration=True,encoding='UTF-8',standalone=True)}
  for i,num in enumerate([42,44,45]):changes[f'xl/charts/chart{num}.xml']=assets.read(f'xl/charts/chart{i+1}.xml')
  with ZipFile(dest,'w',ZIP_DEFLATED,compresslevel=6) as target:
   for item in z.infolist():target.writestr(copy.copy(item),changes.get(item.filename,z.read(item.filename)))
with ZipFile(src) as a,ZipFile(dest) as b:
 changed=[n for n in a.namelist() if a.read(n)!=b.read(n)];assert set(changed)==set(changes)
proof={'revision':'R4','native_R3':'FAIL: user screenshot identifies drawing7.xml removed','action':'Rebuilt Hotel drawing and three chart parts using standard XlsxWriter primitives; removed bespoke doughnut label/manual-layout configuration. No engine/worksheet changes.','changed_parts':changed,'series':seriesproof,'all_other_parts_byte_identical':True,'sha256':hashlib.sha256(dest.read_bytes()).hexdigest(),'native_R4':'NOT TESTED'}
(out/'Correcao_Desenho_R4.json').write_text(json.dumps(proof,indent=2));print(proof)
