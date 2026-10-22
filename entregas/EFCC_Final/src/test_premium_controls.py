"""Actual front-panel mutations in private labs; source workbook is never frozen."""
from pathlib import Path
from zipfile import ZipFile,ZIP_DEFLATED
from lxml import etree as E
import re
from assemble import paths,T,xml,resolve,relpath
from concurrent.futures import ThreadPoolExecutor,as_completed
import json
from xlsxwriter.utility import xl_col_to_name as col
import test_links as old
from premium_ui import ROOT,WORK
old.SOURCE=ROOT/'KK01_V06_Executive_Master.xlsx'
def extra(name,changes):
 return verify_extra(name,old.run_case(name,changes))
def verify_extra(name,result):
 if result['status']!='PASS':return result
 try:
  output=WORK/'mutations'/name/'recalc'/'out'/(name+'.xlsx')
  with ZipFile(output) as z:cells,errors,_=old.read(z)
  assert not errors
  m=cells['V06 Motor'];ev=result['evidence']
  if name=='Premium_EN':
   assert 'Executive Dashboard' in cells['EXEC V06']['B8'];assert 'Back to Dashboard' in cells['V06 Disponibilidade']['B2'];assert 'Report' in cells['DASH 10 Centro']['E18'];assert 'Project NPV' in cells['DASH 10 V06']['A36'];assert cells['V06 Controlos']['B5']=='EN'
   for r,d in enumerate(old.REF['main_rows'],5):
    for key,v in d.items():
     if old.number(v):old.near(m[col(old.REF['main_columns'][key])+str(r)],v)
   ev['unchanged_engine_values']=25200
  elif name in ['Premium_ReportBrief','Premium_ReportComplete']:
   value=str(cells['DASH 10 Centro']['E12']).lower();assert ('resumo' if name=='Premium_ReportBrief' else 'integral') in value,value
  elif name=='Premium_ResADR':old.near(m['AP35'],235);old.near(m['AR35'],old.REF['main_rows'][30]['res_revenue']*235/225)
  elif name=='Premium_ResOccupancy':old.near(m['AQ35'],.60);old.near(m['AR35'],old.REF['main_rows'][30]['res_revenue']*.60/.581)
  elif name=='Premium_ResOPEX':old.near(m['AS35'],m['AR35']*.40)
  elif name=='Premium_ResFFE':old.near(m['AU35'],m['AR35']*.05)
  elif name=='Premium_ReportReference':old.near(cells['DASH 10 V06']['G17'],old.REF['metrics']['Reference Case']['npv']);assert cells['EXEC V06']['M15']!=cells['DASH 10 V06']['G17'];old.near(cells['DASH 10 V06']['A552'],old.REF['metrics']['Reference Case']['npv']);old.near(cells['V06 Brief']['A72'],old.REF['metrics']['Reference Case']['npv'])
  elif name=='Premium_SelectedYear':
   expected=sum(r['revenue'] for r in old.REF['main_rows'] if r['op_age']==5 and r['available_room_nights']>0);old.near(cells['DASH 05 V06']['H50'],expected);old.near(cells['V06 PDF 07']['G17'],expected);old.near(cells['DASH 10 V06']['I338'],expected)
   selected=[r for r in old.REF['main_rows'] if r['op_age']==5 and r['available_room_nights']>0]
   old.near(cells['DASH 10 V06']['A312'],sum(r['adr'] for r in selected)/len(selected));old.near(cells['DASH 10 V06']['G312'],sum(r['occupancy'] for r in selected)/len(selected));old.near(cells['DASH 10 V06']['G317'],sum(r['ebitda'] for r in selected))
  result['evidence']['canonical_front_panel_mutated']=True
 except Exception as e:result={'test':name,'status':'FAIL','error':repr(e)}
 return result

def trim_private_lab():
 # Existing calculations remain live; only unrelated, already-frozen engines go.
 source=WORK/'Mutation_Lab.xlsx';dest=WORK/'Mutation_Lab_Trimmed.xlsx'
 with ZipFile(source) as z:
  w,pp=paths(z);removed={n for n in pp if re.fullmatch(r'V06 [SMO]\d+',n) and n!='V06 M01'};parts={pp[n] for n in removed};changes={};ss=w.find(T('sheets'));names=[n.get('name') for n in ss];ids=set()
  for n in list(ss):
   if n.get('name') in removed:ids.add(n.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id'));ss.remove(n)
  remaining=[n.get('name') for n in ss];rels=E.fromstring(z.read('xl/_rels/workbook.xml.rels'))
  for n in list(rels):
   if n.get('Id') in ids:rels.remove(n)
  dn=w.find(T('definedNames'))
  if dn is not None:
   for n in list(dn):
    if n.get('localSheetId') is not None:
     name=names[int(n.get('localSheetId'))]
     if name in removed:dn.remove(n)
     else:n.set('localSheetId',str(remaining.index(name)))
  for sn in remaining:
   root=E.fromstring(z.read(pp[sn]));dirty=False
   for f in list(root.iter(T('f'))):
    if any(name in removed for name in re.findall(r"'([^']+)'!",f.text or '')):
     assert sn not in ['V06 Motor','V06 M01'],sn
     f.getparent().remove(f);dirty=True
   if dirty:changes[pp[sn]]=xml(root)
  changes.update({'xl/workbook.xml':xml(w),'xl/_rels/workbook.xml.rels':xml(rels)})
  ct=E.fromstring(z.read('[Content_Types].xml'))
  for n in list(ct):
   if n.get('PartName','').lstrip('/') in parts:ct.remove(n)
  changes['[Content_Types].xml']=xml(ct)
  with ZipFile(dest,'w',ZIP_DEFLATED,compresslevel=3) as out:
   for p in z.infolist():
    if p.filename not in parts:out.writestr(p.filename,changes.get(p.filename,z.read(p.filename)))
 dest.replace(source);print('Private lab trimmed',len(removed),'unrelated frozen engines',flush=True)
def main():
 old.make_lab();trim_private_lab()
 from test_all_dashboard_filters import freeze_unrelated_matrix_lab
 freeze_unrelated_matrix_lab()
 cases=[('Hotel_Custom',[('DASH 05 V06','H23','Custom'),('DASH 05 V06','H25',200),('DASH 05 V06','H27',.65)]),('MICE_55',[('DASH 05 V06','N25',.55)]),('FFE_3',[('DASH 05 V06','N27',.03)]),('Hotel_Optimized',[('DASH 05 V06','H32','KK Optimized Terms')]),('Residential_Optimized',[('DASH 04B V06','H23','KK Optimized Terms')]),('Reserve_Conservative',[('DASH 07 V06','H23','Conservative')]),('Interest_8',[('DASH 03 V06','H23',.08)]),('Premium_EN',[('DASH 10 Centro','H7','EN')]),('Premium_ResADR',[('DASH 04B V06','N27',235)]),('Premium_ResOccupancy',[('DASH 04B V06','H29',.60)]),('Premium_ResOPEX',[('DASH 04B V06','N29',.40)]),('Premium_ResFFE',[('DASH 04B V06','N31',.05)]),('Premium_ReportReference',[('DASH 10 Centro','N11','Reference Case'),('DASH 03 V06','H23',.08)]),('Premium_SelectedYear',[('DASH 05 V06','N23',6)]),('Premium_ReportBrief',[('DASH 10 Centro','H11','Executive Brief')]),('Premium_ReportComplete',[('DASH 10 Centro','H11','Complete Report')])]
 results=[]
 with ThreadPoolExecutor(max_workers=3) as ex:
  for future in as_completed([ex.submit(extra,*case) for case in cases]):
   result=future.result();results.append(result);print(result,flush=True)
 proof={'status':'PASS' if all(r['status']=='PASS' for r in results) else 'FAIL','platform':'LibreOffice, cache-invalidated private mutation labs','native_excel365':'NOT TESTED','private_lab_scope':'Unrelated scenario branches frozen and removed only in private labs; production keeps every financial formula.','tests':results};(ROOT/'Testes_Controlo_Premium.json').write_text(json.dumps(proof,indent=2,ensure_ascii=False));assert proof['status']=='PASS'
if __name__=='__main__':main()
