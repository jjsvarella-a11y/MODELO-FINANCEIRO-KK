"""Retain original report metadata and the actual global PT/EN dependency path.
The original control engine references Center!H7, H15/N15 and bank recipient H122.
These canonical cells must survive redesign of the visible reporting center.
"""
from profile_controls_v08 import *
from premium_ui import LANG
def wire(source,target,reference):
 src=Source(source);ref=Source(reference);lang=ref.get('V06 Controlos','B5')[1]
 if lang not in ['PT','EN']:raise ValueError('Reference global language must be PT or EN')
 styles=add_styles(src.styles);changes={}
 for sn,*_ in DASHES:
  s=Sheet(E.fromstring(src.z.read(src.pp[sn])),styles);s.text('U5:V5','Idioma','Language','nav')
  if sn=='EXEC V06':
   s.merge('W5:Y5',lang,style='input_text');dvs=s.r.find(T('dataValidations'))
   for old in list(dvs):
    if 'W5' in old.get('sqref','').split():dvs.remove(old)
   node=E.SubElement(dvs,T('dataValidation'),type='list',sqref='W5',showErrorMessage='1',errorStyle='stop');E.SubElement(node,T('formula1')).text='"PT,EN"'
  else:s.merge('W5:Y5',lang,'HYPERLINK("#\'EXEC V06\'!W5",'+LANG+')','nav')
  s.merge('A5:D5','PT / EN · '+lang,'HYPERLINK("#\'EXEC V06\'!W5","PT / EN · "&'+LANG+')','nav');changes[src.pp[sn]]=s.finish()
 s=Sheet(E.fromstring(changes[src.pp[CENTER]]),styles);s.put('H7',lang,q('EXEC V06','W5'))
 preserved=[]
 for addr in ['H11','N11','H15','N15']:
  f,v=ref.get(CENTER,addr);s.put(addr,v,f);preserved.append({'cell':CENTER+'!'+addr,'value':v,'formula':f})
 s.put('H122',src.get(BASE7,'D'+str(ROWS['bank_institution']))[1],active('bank_institution'))
 # Legacy metadata remains available to the existing calculation/print modules;
 # only the new compact reporting center is visible to the user.
 for rr,row in s.rows.items():
  if rr>=62:row.set('hidden','1')
 changes[src.pp[CENTER]]=s.finish();changes['xl/styles.xml']=xml(src.styles);rewrite(source,target,changes);src.z.close();ref.z.close();return {'global_language':lang,'global_input':'EXEC V06!W5','canonical_link':'DASH 10 Centro!H7 → EXEC V06!W5','preserved_metadata':preserved,'bank_recipient':'Active bank institution, no artificial zero'}
if __name__=='__main__':
 p=argparse.ArgumentParser();p.add_argument('source',type=Path);p.add_argument('target',type=Path);p.add_argument('--reference',type=Path,required=True);p.add_argument('--proof',type=Path,required=True);a=p.parse_args();a.proof.write_text(json.dumps(wire(a.source,a.target,a.reference),ensure_ascii=False,indent=2))
