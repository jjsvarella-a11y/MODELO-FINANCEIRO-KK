"""Saved status includes raw selector/manual choices, not just effective numbers."""
from scenario_v07 import *
def upgrade(file=FINAL7):
 source=Source(file);changes={}
 for mode in MAPS:
  sn=store(mode);sheet=Sheet(E.fromstring(source.z.read(source.pp[sn])),{});identifier=source.get(STATUS7,'B5' if mode=='Base' else 'B6')[1];savedfile=VERSIONS/(str(identifier)+'.json');saved=json.loads(savedfile.read_text(encoding='utf-8-sig'))['raw_inputs'] if savedfile.exists() else raw_snapshot(source,mode)
  for key,entry in saved.items():
   sheet.put('J'+str(ROWS[key]),entry['selection']);sheet.put('K'+str(ROWS[key]),entry['custom'])
  sheet.put('J5','Last saved selection');sheet.put('K5','Last saved manual value');cols=sheet.r.find(T('cols'))
  for node in list(cols):
   if int(node.get('min'))>=10 and int(node.get('max'))<=11:cols.remove(node)
  E.SubElement(cols,T('col'),min='10',max='11',width='18',hidden='1',customWidth='1');dimension=sheet.r.find(T('dimension'))
  if dimension is not None:dimension.set('ref','A1:K'+str(len(DEFS)+5))
  changes[source.pp[sn]]=sheet.finish()
 sheet=Sheet(E.fromstring(source.z.read(source.pp[STATUS7])),{})
 for row,mode in [(9,'Base'),(10,'Comparative')]:
  sn=store(mode);terms=['SUMPRODUCT(--('+q(sn,f'${a}$6:${a}${len(DEFS)+5}')+'<>'+q(sn,f'${b}$6:${b}${len(DEFS)+5}')+'))' for a,b in [('D','E'),('B','J'),('C','K')]];formula='IF('+'+'.join(terms)+'=0,"Saved","Unsaved")';sheet.put('B'+str(row),'Saved',formula)
 changes[source.pp[STATUS7]]=sheet.finish();tmp=WORK7/'Guard.xlsx';rewrite(file,tmp,changes);source.z.close();shutil.move(tmp,file)
if __name__=='__main__':upgrade()
