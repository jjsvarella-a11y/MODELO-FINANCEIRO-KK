"""Explicit undefined-metric handling and valid input domains; no IFERROR masking."""
from zipfile import ZipFile,ZIP_DEFLATED
from lxml import etree as E
import copy,json,re
from closure_r5 import *

def apply():
 source=Source(MASTER);changes={};ranges=[]
 with ZipFile(MASTER) as z:
  _,pp=paths(z)
  r=E.fromstring(z.read(pp[IMP]));s=Sheet(r,{})
  def put(addr,f):c=s.cells[addr];s.put(addr,source.get(IMP,addr)[1],f,int(c.get('s','0')))
  put('C10','IF(ISNUMBER(\'V06 Cenarios\'!K5),\'V06 Cenarios\'!K5,"N/A")');put('C11','IF(SUM('+motor('investor_outlay')+')>0,\'V06 Cenarios\'!L5,"N/A")')
  for rr in range(5,14):
   put('D'+str(rr),'IF(AND(ISNUMBER(C'+str(rr)+'),ISNUMBER(B'+str(rr)+')),C'+str(rr)+'-B'+str(rr)+',"N/A")');put('E'+str(rr),'IF(AND(ISNUMBER(D'+str(rr)+'),ABS(B'+str(rr)+')>0.000001),D'+str(rr)+'/ABS(B'+str(rr)+'),"N/A")')
  changes[pp[IMP]]=s.finish()
  for sn,p in pp.items():
   if sn not in [n[0] for n in DASHES]+['DASH 10 V06','V06 Brief']+[f'V06 PDF {i:02}' for i in range(1,10)]:continue
   r=E.fromstring(z.read(p));dirty=False
   for f in r.iter(T('f')):
    if re.fullmatch(r"'V06 (?:Impact R5'!D10|Investor R5'![DP]29)\*100",f.text or ''):
     ref=f.text[:-4];f.text='IF(ISNUMBER('+ref+'),'+ref+'*100,"N/A")';dirty=True
   if dirty:changes[p]=xml(r)
  # Valid bounds prevent zero-price / zero-coverage divisions rather than hiding them.
  for sn,addr,kind,lo,hi in [('DASH 01 V06','H23','whole',0,120),('DASH 04A V06','N23','whole',0,120),('DASH 04A V06','N25','decimal',.01,3),('DASH 02 V06','H27','decimal',0,1),('DASH 02 V06','N27','decimal',0,10),('DASH 04B V06','H31','decimal',0,1),('DASH 03 V06','H25','whole',1,360),('DASH 03 V06','H27','whole',1,180),('DASH 03 V06','H31','decimal',0,1e9),('DASH 07 V06','N27','decimal',.01,10)]:
   p=pp[sn];r=E.fromstring(changes.get(p,z.read(p)));dvs=r.find(T('dataValidations'))
   if dvs is None:dvs=E.SubElement(r,T('dataValidations'),count='0')
   for old in list(dvs):
    if old.get('sqref')==addr:dvs.remove(old)
   dv=E.SubElement(dvs,T('dataValidation'),type=kind,operator='between',sqref=addr,allowBlank='0',showErrorMessage='1',errorStyle='stop',errorTitle='Invalid input',error=f'Use a value between {lo} and {hi}.',showInputMessage='1',promptTitle='Model input',prompt=f'Permitted range: {lo} to {hi}.');E.SubElement(dv,T('formula1')).text=str(lo);E.SubElement(dv,T('formula2')).text=str(hi);dvs.set('count',str(len(dvs)));changes[p]=Sheet(r,{}).finish();ranges.append({'sheet':sn,'cell':addr,'type':kind,'min':lo,'max':hi})
  temp=MASTER.with_suffix('.guards.xlsx')
  with ZipFile(temp,'w',ZIP_DEFLATED,compresslevel=6) as out:
   for item in z.infolist():out.writestr(copy.copy(item),changes.get(item.filename,z.read(item.filename)))
 temp.replace(MASTER);(OUT/'Dominios_Inputs_R5.json').write_text(json.dumps({'ranges':ranges,'undefined_IRR':'N/A when no investment return can be computed; differences and percentage points remain N/A, not arithmetic errors.','method':'ISNUMBER guards for undefined metrics; no IFERROR was added. Invalid zero pricing/coverage is stopped by validation.'},indent=2));print('Undefined metrics and input bounds checked',flush=True)
if __name__=='__main__':apply()
