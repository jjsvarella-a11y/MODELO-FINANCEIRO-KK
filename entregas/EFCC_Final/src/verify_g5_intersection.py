"""Inspect the supplied Excel formula and audit the current scalar engine family."""
from pathlib import Path
from zipfile import ZipFile
from lxml import etree as E
import json,re
from assemble import paths,T
ROOT=Path(__file__).resolve().parents[1]
def verify():
 received=(ROOT/'Excel365/Formula_G5_Recebida.txt').read_text().strip();normal=received.lstrip('=').replace('@','').replace(';',',').replace('SOMARPRODUTO(','SUMPRODUCT(').replace('SE(','IF(')
 old=ROOT.parent/'V06_DASH10/KK01_V06_Executive_Master.xlsx';current=ROOT/'KK01_V06_Executive_Master.xlsx'
 with ZipFile(old) as z:
  _,pp=paths(z);r=E.fromstring(z.read(pp['V06 Motor']));original=r.find('.//'+T('c')+'[@r="G5"]').findtext(T('f'))
 assert normal==original,'Received formula differs from previous version beyond locale and implicit intersections'
 violations=[];engines=0;count=0
 with ZipFile(current) as z:
  _,pp=paths(z)
  for sn,part in pp.items():
   if not (sn in ['V06 Motor','V06 Referencia'] or re.fullmatch(r'V06 [SMO]\d+',sn)):continue
   engines+=1;r=E.fromstring(z.read(part))
   for c in r.iter(T('c')):
    f=c.findtext(T('f'))
    if f is None:continue
    count+=1
    if '@' in f or 'SUMPRODUCT(' in f or '_xlfn.SINGLE' in f:violations.append([sn,c.get('r'),f[:160]])
 assert not violations,violations[:3];assert engines==69
 result={'status':'PASS','scope':'Formula identity and removal of array-context dependencies; not a native Excel acceptance test','user_reported_cell':'V06 Motor!G5','received_implicit_intersections':received.count('@'),'received_matches_previous_g5_after_locale_and_at_normalisation':True,'identified_g5_failure':'The @IF result is scalar while SUMPRODUCT first argument is a 30-row weight vector, producing mismatched argument dimensions.','current_engines_inspected':engines,'current_engine_formulas_inspected':count,'current_engine_SUMPRODUCT_at_or_SINGLE_violations':violations,'repair':'Scalar SUM/MAX/MIN construction integration and scalar per-unit sales schedules, preserving scenario logic.','expected_baseline_g5':0.01,'native_excel365_current_candidate':'NOT TESTED'}
 (ROOT/'Diagnostico_G5.json').write_text(json.dumps(result,indent=2,ensure_ascii=False));print(result)
if __name__=='__main__':verify()
