"""Accept both native numeric percent selections and literal preset labels."""
from zipfile import ZipFile,ZIP_DEFLATED
from lxml import etree as E
import copy,json
from closure_r5 import MASTER,OUT,ACT,Sheet,q
from assemble import paths,T,xml

def apply():
 with ZipFile(MASTER) as z:
  _,pp=paths(z);r=E.fromstring(z.read(pp['V06 Registo Ativo']));s=Sheet(r,{});old=s.cells['B65'];previous=old.findtext(T('f'));p=q('V06 Matriz ADR','F3');f='IF(ISNUMBER('+p+'),'+p+','+previous+')';s.put('B65',.12,f,int(old.get('s','0')))
  styles=E.fromstring(z.read('xl/styles.xml'));xfs=styles.find(T('cellXfs'));screen=E.fromstring(z.read(pp['DASH 07 V06']));cell=screen.find('.//'+T('c')+'[@r="H29"]');xf=copy.deepcopy(xfs[int(cell.get('s'))]);xf.set('numFmtId','9');xf.set('applyNumberFormat','1');cell.set('s',str(len(xfs)));xfs.append(xf);xfs.set('count',str(len(xfs)))
  changes={pp['V06 Registo Ativo']:s.finish(),pp['DASH 07 V06']:xml(screen),'xl/styles.xml':xml(styles)};temp=MASTER.with_suffix('.discount.xlsx')
  with ZipFile(temp,'w',ZIP_DEFLATED,compresslevel=6) as out:
   for item in z.infolist():out.writestr(copy.copy(item),changes.get(item.filename,z.read(item.filename)))
 temp.replace(MASTER);(OUT/'Desconto_R5.json').write_text(json.dumps({'annual_rate':'V06 Registo Ativo!B65','selector':'DASH 07 V06!H29','custom':'DASH 07 V06!N25, only when Custom is selected','native_numeric_percent':'Numeric values such as 0.08 and 0.10 are accepted directly, as well as text 8% / 10% / 12%.','monthly_conversion':'(1+annual)^(1/12)-1','cashflows':'V06 F01!B6:B185 plus month-zero B5; all matrix point cashflows use the same annual rate','IRR':'Independent of the discount rate','native_excel365':'NOT TESTED'},indent=2));print('Native numeric percent presets supported',flush=True)
if __name__=='__main__':apply()
