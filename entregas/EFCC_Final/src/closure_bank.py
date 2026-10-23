"""Expose the existing draw window/facility and requested amortisation study.
Linear branch is the original formula unchanged. Bullet is an explicit alternative.
"""
from zipfile import ZipFile,ZIP_DEFLATED
from lxml import etree as E
import copy,json,re
from closure_r5 import *

def apply():
 with ZipFile(MASTER) as z:
  _,pp=paths(z);changes={};ss=E.fromstring(z.read('xl/sharedStrings.xml'));text=[''.join(n.itertext()) for n in ss];r=E.fromstring(z.read(pp[ACT]));styles={}
  for c in r.find(T('sheetData'))[0]:
   if c.get('t')=='s':
    name=text[int(c.findtext(T('v')))];
    if name in ['label','input','input_money','input_number','note']:styles[name]=int(c.get('s'))
  for sn in [REF,ACT,'V06 Controlos','DASH 03 V06']:
   changes[pp[sn]]=E.fromstring(z.read(pp[sn]))
  ref=Sheet(changes[pp[REF]],styles);active=Sheet(changes[pp[ACT]],styles);control=Sheet(changes[pp['V06 Controlos']],styles);screen=Sheet(changes[pp['DASH 03 V06']],styles)
  for n,item in enumerate(INPUTS,5):
   if (item['sheet'],item['cell']) not in [('DASH 03 V06','H27'),('DASH 03 V06','H29'),('DASH 03 V06','H31')]:continue
   value=item['value'];ref.put('A'+str(n),item['sheet']+'!'+item['cell']);ref.put('B'+str(n),value);active.put('A'+str(n),item['label_en']);active.put('B'+str(n),value,'IF(\'EXEC V06\'!H23="Reference",'+q(REF,'B'+str(n))+','+q(item['sheet'],item['cell'])+')');active.put('C'+str(n),item['sheet']+'!'+item['cell']);active.put('D'+str(n),value,q(REF,'B'+str(n)))
   rr=int(item['cell'][1:]);screen.text(f'E{rr}:G{rr}',item['label_pt'],item['label_en'],'label');screen.merge(f'H{rr}:J{rr}',value,style='input_money' if item['kind']=='money' else 'input');dvs=screen.r.find(T('dataValidations'))
   if dvs is None:dvs=E.SubElement(screen.r,T('dataValidations'),count='0')
   if not any(d.get('sqref')==item['cell'] for d in dvs):
    dv=E.SubElement(dvs,T('dataValidation'),type='list' if item.get('options') else 'decimal',sqref=item['cell'],showErrorMessage='1',allowBlank='0')
    if item.get('options'):E.SubElement(dv,T('formula1')).text='"'+','.join(item['options'])+'"'
    else:dv.set('operator','between');E.SubElement(dv,T('formula1')).text='1';E.SubElement(dv,T('formula2')).text='180' if rr==27 else '1000000000'
   dvs.set('count',str(len(dvs)))
  control.put('B40',30000000,a('DASH 03 V06','H31'));control.put('B81',30,a('DASH 03 V06','H27'))
  for area,sn,addr,pt,en in [('K27:P28','DASH 07 V06','H23','Política de reservas →','Reserve policy →'),('K29:P30','DASH 05 V06','H23','Stress operacional →','Operational stress →'),('K31:P32','DASH 04A V06','H23','Stress de vendas →','Sales stress →')]:screen.merge(area,pt,'HYPERLINK("#\''+sn+'\'!'+addr+'",'+tr(pt,en)+')','input')
  screen.text('E34:P37','Atrasar a obra não prorroga automaticamente a janela bancária. Pode testar o último mês de utilização e Linear/Bullet. Condições são simuladas, não aprovadas pelo banco.','Construction delay does not automatically extend the draw window. Test the last draw month and Linear/Bullet. Terms are simulated, not bank-approved.','note')
  for sn in [REF,ACT,'V06 Controlos','DASH 03 V06']:changes[pp[sn]]=Sheet(changes[pp[sn]],styles).finish()
  count=0
  for sn,p in pp.items():
   if sn!='V06 Motor' and not re.fullmatch(r'V06 [SMO]\d+',sn):continue
   root=E.fromstring(z.read(p));s=Sheet(root,{})
   for rr in range(5,185):
    addr=col(CC['principal'])+str(rr);cell=s.cells[addr];original=cell.findtext(T('f'))
    if '"Bullet"' in original:continue
    maturity=re.search(r'/MAX\(1,\(\(\((.*?)\)-',original)
    assert maturity is not None,(sn,rr,original)
    f='IF('+a('DASH 03 V06','H29')+'="Bullet",IF('+q(sn,'A'+str(rr))+'>=('+maturity[1]+'),'+q(sn,col(CC['debt_open'])+str(rr))+',0),'+original+')';s.put(addr,float(cell.findtext(T('v')) or 0),f,int(cell.get('s','0')));count+=1
   changes[p]=s.finish()
  # Restore native public navigation destinations that happen to be parameter cells.
  reverse={a(i['sheet'],i['cell']):q(i['sheet'],i['cell']) for i in INPUTS}
  for sn,p in pp.items():
   if sn in [REF,ACT] or sn=='V06 Motor' or re.fullmatch(r'V06 [SMO]\d+',sn):continue
   data=changes.get(p,z.read(p));root=E.fromstring(data);dirty=False
   for f in root.iter(T('f')):
    if 'HYPERLINK(' in (f.text or '') and q(ACT,'') in f.text:
     for old,new in reverse.items():f.text=f.text.replace('\"#'+old,'\"#'+new)
     dirty=True
   if dirty:changes[p]=xml(root)
  temp=MASTER.with_suffix('.bank.xlsx')
  with ZipFile(temp,'w',ZIP_DEFLATED,compresslevel=6) as out:
   for item in z.infolist():out.writestr(copy.copy(item),changes.get(item.filename,z.read(item.filename)))
 temp.replace(MASTER);spec=json.loads((OUT/'Implementation.json').read_text());spec['inputs']=INPUTS;spec['parameter_routing']={i['sheet']+'!'+i['cell']:a(i['sheet'],i['cell']) for i in INPUTS};spec['amortisation_formulas_with_preserved_linear_branch']=count;spec['additional_bank_controls']=['DASH 03 V06!H27 draw window','DASH 03 V06!H29 Linear/Bullet','DASH 03 V06!H31 facility'];(OUT/'Implementation.json').write_text(json.dumps(spec,indent=2,ensure_ascii=False));print('Bank controls and navigation repaired',count,flush=True)
if __name__=='__main__':apply()
