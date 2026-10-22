"""A macro-free, financial-engine-independent place to insert the original logo.
Header links open the real in-workbook insertion area. Global propagation is an
explicit Windows Excel COM operation supplied in the native validation script.
No IMAGE()/rich-data dependency or simulated upload button is introduced.
"""
from pathlib import Path
from zipfile import ZipFile,ZIP_DEFLATED
from lxml import etree as E
import sys,re
import xlsxwriter
from assemble import paths,T,xml
sys.path.insert(0,str(Path(__file__).resolve().parents[2]/'V06/src'))
from merge import merge
DASHES=['EXEC V06','DASH 01 V06','DASH 02 V06','DASH 03 V06','DASH 04A V06','DASH 04B V06','DASH 05 V06','DASH 06 V06','DASH 07 V06','DASH 10 V06']
WORK=Path('/workspace/kk01_efcc')
def add(source,dest):
 module=WORK/(Path(dest).stem+'_identity.xlsx');wb=xlsxwriter.Workbook(str(module));s=wb.add_worksheet('V06 Identidade');s.hide_gridlines(2);s.hide_row_col_headers();s.set_column('A:A',3);s.set_column('B:H',14);s.set_column('I:I',3);s.set_column('J:O',16);s.set_default_row(20)
 title=wb.add_format({'font_name':'Aptos','font_size':20,'bold':True,'font_color':'FFFFFF','bg_color':'B31F2B'});text=wb.add_format({'font_name':'Aptos','font_size':11,'text_wrap':True,'valign':'vcenter'});box=wb.add_format({'bg_color':'F2F5F8','font_color':'526A80','border':1,'border_color':'B31F2B','text_wrap':True,'align':'center','valign':'vcenter'});head=wb.add_format({'font_name':'Aptos','font_size':12,'bold':True,'font_color':'203246'});s.merge_range('B2:O3','JS-CV — IDENTIDADE / LOGÓTIPO · IDENTITY / LOGO',title);s.write_url('B5',"internal:'EXEC V06'!A1",text,string='← Dashboard Executivo / Executive Dashboard')
 s.merge_range('B7:H7','INSERIR O FICHEIRO ORIGINAL AQUI / INSERT ORIGINAL FILE HERE',head)
 s.merge_range('B8:H20','1. Inserir → Imagens → Este dispositivo\n2. Escolher o JPG/PNG original\n3. Colocar uma única imagem nesta área\n\nInsert → Pictures → This Device\nChoose the original JPG/PNG\nPlace one image inside this area',box)
 steps=[('PT','1. Abra Inserir → Imagens → Este dispositivo no Excel. Escolha o JPG/PNG original e posicione-o na área B8:H20. Pode também colar a imagem nesta folha. Não redesenhe a arte.'),('PT','2. Guarde o livro. A imagem fica incorporada no próprio XLSX. Inserir aqui não altera pressupostos nem fórmulas financeiras.'),('PT','3. Para aplicar o mesmo logótipo a todos os cabeçalhos, execute Excel365/Executar_Validacao.cmd no Windows com Excel instalado. O teste lê esta imagem, sincroniza os cabeçalhos numa cópia, recalcula e guarda essa cópia com os resultados. Este passo não foi testado neste ambiente Linux.'),('EN','1. In Excel, use Insert → Pictures → This Device. Select the original JPG/PNG and position one picture in B8:H20, or paste it into this sheet.'),('EN','2. Save the workbook. The picture is embedded in the XLSX. Inserting it here does not alter financial inputs or formulas.'),('EN','3. For every header, run Excel365/Executar_Validacao.cmd on Windows with Excel installed. It reads this picture, synchronises the headers in a copy, recalculates and saves the tested copy. This step has not been tested in Linux.')]
 for i,(language,message) in enumerate(steps):r=7+i*4;s.merge_range(r,9,r+2,14,language+' — '+message,text)
 s.merge_range('B23:H26','Área de identidade independente dos cálculos. Não contém VBA, IMAGE() ou fórmulas financeiras. A sincronização dos cabeçalhos é explícita e não uma falsa função automática. / Independent identity area; no VBA, IMAGE() or financial formulas. Header sync is explicit, not a simulated automatic action.',text);s.set_landscape();s.set_paper(9);s.fit_to_pages(1,1);s.print_area('A1:O32');wb.close()
 inter=Path(dest).with_suffix('.identity.xlsx');merge(source,module,inter)
 with ZipFile(inter) as z,ZipFile(dest,'w',ZIP_DEFLATED,compresslevel=6) as zz:
  w,pp=paths(z);sheets=w.find(T('sheets'));old=[x.get('name') for x in sheets];order=DASHES+['V06 Identidade']+[x for x in old if x not in DASHES+['V06 Identidade']];lookup={x.get('name'):x for x in sheets};sheets[:]=[lookup[x] for x in order]
  for n in w.findall(T('definedNames')+'/'+T('definedName')):
   if n.get('localSheetId') is not None:n.set('localSheetId',str(order.index(old[int(n.get('localSheetId'))])))
  for v in w.find(T('bookViews')):v.set('activeTab','0');v.set('firstSheet','0')
  changes={'xl/workbook.xml':xml(w)}
  for sn in DASHES+['V06 Brief']:
   root=E.fromstring(z.read(pp[sn]));sd=root.find(T('sheetData'));addresses=['N2'] if sn in DASHES[:9] else ['J'+str(1+50*i) for i in range(10 if sn=='DASH 10 V06' else 2)]
   for addr in addresses:
    r=int(''.join(filter(str.isdigit,addr)));row=sd.find(T('row')+'[@r="'+str(r)+'"]');cell=row.find(T('c')+'[@r="'+addr+'"]')
    if cell is None:cell=E.SubElement(row,T('c'),r=addr)
    for node in list(cell):cell.remove(node)
    cell.set('t','str');E.SubElement(cell,T('f')).text='HYPERLINK("#\'V06 Identidade\'!B8",IF(\'V06 Controlos\'!$B$5="PT","Inserir logótipo","Insert logo"))';E.SubElement(cell,T('v')).text='Inserir logótipo'
   # Keep workbook selection consistent after merging the identity sheet.
   for v in root.findall(T('sheetViews')+'/'+T('sheetView')):v.set('tabSelected','1' if sn=='EXEC V06' else '0')
   changes[pp[sn]]=xml(root)
  for part in z.infolist():
   data=changes.get(part.filename,z.read(part.filename))
   if part.filename.startswith('xl/worksheets/') and part.filename.endswith('.xml') and part.filename not in changes:
    data=re.sub(rb'(<sheetView\b[^>]*\btabSelected=")(?:1|true)(")',rb'\g<1>0\2',data)
   zz.writestr(part.filename,data)
 inter.unlink();print('Identity insertion sheet and real navigation added',dest)
if __name__=='__main__':add(Path(sys.argv[1]),Path(sys.argv[2]))
