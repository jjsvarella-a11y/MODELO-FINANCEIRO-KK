"""Resolve all public parameter links; eliminate stale reference/discount labels."""
from bank_r6 import *
from zip_copy import rewrite
import csv

def apply():
 changes={};nav=[];fixes=[]
 with ZipFile(BM) as z:
  wb,pp=paths(z);visible={n.get('name'):n.get('state','visible')=='visible' for n in wb.find(T('sheets'))};reverse={v.split('!')[1].replace('$',''):(sn,addr) for (sn,addr),v in MAPPING.items()}
  for sn,p in pp.items():
   if sn=='V06 Motor' or re.fullmatch(r'V06 [SMO]\d+',sn):continue
   root=E.fromstring(z.read(p));s=Sheet(root,{});dirty=False
   for c in root.iter(T('c')):
    f=c.find(T('f'))
    if f is None:continue
    if 'HYPERLINK(' in f.text:
     def repair(m):
      nonlocal dirty
      target,addr=m.groups()
      if target==ACT:
       owner,cell=reverse[addr.replace('$','')];dirty=True;fixes.append([sn,c.get('r'),target+'!'+addr,owner+'!'+cell]);return 'HYPERLINK("#\''+owner+'\'!'+cell+'"'
      return m[0]
     f.text=re.sub(r'HYPERLINK\("#\'([^\']+)\'!([^\"]+)"',repair,f.text)
    # Pure preset-display aliases must show the single effective discount source.
    if sn not in ['V06 Registo Ativo',ACT,REF,DAT] and f.text.startswith('IF(') and all(x in f.text for x in ['"8%"','"10%"','"12%"',"'V06 Matriz ADR'!F3"]):f.text=q('V06 Registo Ativo','B65');dirty=True;fixes.append([sn,c.get('r'),'old text-only discount display','single effective annual discount'])
   if sn=='DASH 10 V06':
    c=s.cells['G492'];s.put('G492',-12742848.2445606,q(IMP,'B9'),int(c.get('s')));dirty=True
   if dirty:changes[p]=s.finish()
  # All visible auxiliary sheets get an explicit Home and responsible Back path.
  public=[n for n in pp if visible[n] and n not in [x[0] for x in DASHES] and n not in ['DASH 10 V06','V06 Brief']+[f'V06 PDF {i:02}' for i in range(1,10)]]
  owners={PRO:BN,SCH:BN,REP:CENTER,BRF:CENTER,'V06 Constru R5':'DASH 01 V06','V06 Capitulos':'DASH 01 V06','V06 Disponibilidade':'DASH 04A V06','V06 Matriz ADR':'DASH 05 V06','V06 Operacao':'DASH 05 V06','V06 Blocos':'DASH 06 V06','V06 Dicionario':'EXEC V06','V06 Identidade':'EXEC V06','V06 Imprimir':'EXEC V06','V06 Controlos':'EXEC V06','V06 Registo Ativo':'EXEC V06'}
  for sn in public:
   # Do not overwrite a technical table row; use unused far-right AA/AE cells.
   s=Sheet(E.fromstring(changes.get(pp[sn],z.read(pp[sn]))),{});owner=owners.get(sn,'EXEC V06')
   for cell,target,pt,en in [('AA1','EXEC V06','Início','Home'),('AE1',owner,'Voltar ao Dashboard','Back to Dashboard')]:s.put(cell,pt,'HYPERLINK("#\''+target+'\'!A1",'+tr(pt,en)+')')
   changes[pp[sn]]=s.finish()
  for sn,p in pp.items():
   root=E.fromstring(changes.get(p,z.read(p)))
   for c in root.iter(T('c')):
    for m in re.finditer(r'HYPERLINK\("#\'([^\']+)\'!([^\"]+)"',c.findtext(T('f')) or ''):
     target,addr=m.groups();assert target in pp,(sn,target);assert visible[target],(sn,c.get('r'),'target hidden',target)
     nav.append({'origin':sn+'!'+c.get('r'),'destination':target+'!'+addr,'type':'native internal hyperlink','result':'PASS (structural)','excel365_native':'NOT TESTED'})
 temp=BM.with_suffix('.ui.xlsx');rewrite(BM,temp,changes);temp.replace(BM)
 (BOUT/'Navigation_Matrix.json').write_text(json.dumps(nav,indent=2,ensure_ascii=False))
 with (BOUT/'Navigation_Matrix.csv').open('w',newline='') as f:w=csv.DictWriter(f,fieldnames=list(nav[0]));w.writeheader();w.writerows(nav)
 (BOUT/'UI_Corrections.json').write_text(json.dumps(fixes,indent=2,ensure_ascii=False));print('Navigation audited',len(nav),'fixes',len(fixes),flush=True)
if __name__=='__main__':apply()
