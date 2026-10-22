"""Small native-Excel compatibility fixture: exact old formulas vs scalar revisions.
Caches are reference values, explicitly not evidence of native calculation.
"""
from core import *
from zipfile import ZipFile
from lxml import etree as E
from assemble import paths,T
ROOT=Path(__file__).resolve().parents[1]
def formulas(path,sn):
 with ZipFile(path) as z:
  _,pp=paths(z);root=E.fromstring(z.read(pp[sn]));return {c.get('r'):(c.findtext(T('f')),c.findtext(T('v'))) for c in root.iter(T('c'))}
def build():
 old=formulas('/workspace/kk01_efcc/checkpoints/C0_V06_DASH10.xlsx','V06 Motor');new=formulas('/workspace/kk01_efcc/V06_modulos.xlsx','V06 Motor');wb=xlsxwriter.Workbook(str(ROOT/'Excel365/Diagnostico_Formulas_Excel365.xlsx'));wb.set_calc_mode('auto');fm={k:wb.add_format({'num_format':'#,##0.00'} if k=='num' else {'bold':True,'bg_color':'B31F2B','font_color':'FFFFFF'} if k=='head' else {'num_format':'0'} if k=='int' else {'num_format':'0.00%'} if k=='pct' else {}) for k in ['num','head','int','pct','text']};front=wb.add_worksheet('DIAGNOSTICO');front.set_column('A:A',48);front.set_column('B:D',27);front.set_column('E:E',76);front.write_row('A1',['DIAGNÓSTICO — NÃO É O MASTER','Anterior / Old','Escalar / Scalar','Igual / Equal','Procedimento'],fm['head']);front.write('A3','Valores guardados não validam o cálculo no Excel.');front.write('E3','1. Ative edição. 2. Ctrl+Alt+F9. 3. Guarde, feche e reabra. 4. Registe que células falham.');front.set_row(2,40);front.write('E5','As fórmulas antigas são cópias exatas do V06 anterior. As novas usam a mesma curva e as mesmas 94 frações, sem IF matricial/SUMPRODUCT.');front.write('E7','Microsoft Excel 365: NOT TESTED neste ambiente. Este ficheiro permite isolar os dois blocos, sem carregar o master completo.');front.set_row(4,50);front.set_row(6,45)
 ctrl=wb.add_worksheet('V06 Controlos');params={}
 for r,(key,pt,en,v,unit,note) in enumerate(DEFAULTS,5):ctrl.write(r-1,0,pt);ctrl.write(r-1,1,v);params[key]=X(f"'V06 Controlos'!$B${r}",v)
 curve=wb.add_worksheet('V06 Curva')
 for i in range(30):curve.write(i+4,0,i+1);curve.write(i+4,1,sum(c['share']*c['weights'][i] for c in legacy.CHAPTERS))
 units=wb.add_worksheet('V06 Unidades');dataset=[]
 for r,u in enumerate(legacy.UNITS,5):
  d={}
  for c,key in [(0,'id'),(7,'price'),(8,'sale'),(11,'deed')]:units.write(r-1,c,u[key]);d[key]=X(f"'V06 Unidades'!${col(c)}${r}",u[key])
  dataset.append(d)
 model=Model(wb,fm,params,dataset,legacy.CHAPTERS);cal=model.calendar(params);motor=wb.add_worksheet('V06 Motor');motor.write_row('A4',['Month','','End','','','','Old construction weight','Scalar construction weight','','','','','','','Old sales receipts','Scalar sales receipts'],fm['head'])
 for r in range(5,185):
  motor.write(r-1,0,r-4);motor.write(r-1,2,30)
  for dest,src,table in [('G','G',old),('H','G',new),('O','O',old),('P','O',new)]:
   f,v=table[src+str(r)];motor.write_formula(dest+str(r),'='+f,fm['num'],float(v))
 for r,label,oldf,newf,expected in [(5,'Peso total construção','SUM(\'V06 Motor\'!G5:G184)','SUM(\'V06 Motor\'!H5:H184)',1),(6,'Recebimentos totais','SUM(\'V06 Motor\'!O5:O184)','SUM(\'V06 Motor\'!P5:P184)',sum(u['price'] for u in legacy.UNITS)),(9,'Peso construção M1',"'V06 Motor'!G5","'V06 Motor'!H5",float(old['G5'][1])),(10,'Recebimentos M1',"'V06 Motor'!O5","'V06 Motor'!P5",float(old['O5'][1]))]:
  front.write(r-1,0,label);front.write_formula(r-1,1,'='+oldf,fm['num'],expected);front.write_formula(r-1,2,'='+newf,fm['num'],expected);front.write_formula(r-1,3,f'=ABS(B{r}-C{r})<0.00001',None,True)
 front.write('A12','Não há IFERROR a ocultar erros neste diagnóstico.');front.write('E12','O diagnóstico não prova que o problema original tem esta causa; compara os blocos suspeitos em Excel nativo.');front.activate();front.freeze_panes(1,1);wb.close();print('Native probe built')
if __name__=='__main__':build()
