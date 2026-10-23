"""Include early/closing fees in liquidity, funding gaps and dividend safeguards."""
from bank_r6 import *
def apply():
 with ZipFile(BM) as z,ZipFile(MASTER) as base:
  _,pp=paths(z);_,bp=paths(base);changes={}
  for sn,p in pp.items():
   if sn!='V06 Motor' and not re.fullmatch(r'V06 [SMO]\d+',sn):continue
   s=Sheet(E.fromstring(z.read(p)),{});old=Sheet(E.fromstring(base.read(bp[sn])),{})
   for rr in range(5,185):
    def r(c):return q(sn,c+str(rr))
    fields={'DC':'MAX(0,-('+r('CF')+'-'+r('EM')+'-'+r('EN')+'))','CZ':'MAX(0,'+r('CF')+'-'+r('CJ')+'-'+r('BV')+'-'+r('EM')+'-'+r('EN')+')','CV':old.cells['CV'+str(rr)].findtext(T('f')).replace(r('CU'),'('+r('CE')+'+'+r('CA')+'+'+r('BW')+')')+'-'+r('EM')+'-'+r('EN'),'CK':'MAX(0,'+r('BP')+'+'+r('CD')+'+'+r('EL')+'-'+r('CA')+'-'+r('CJ')+')'}
    for cc,f in fields.items():
     c=s.cells[cc+str(rr)];o=old.cells[cc+str(rr)].findtext(T('f'));s.put(cc+str(rr),float(c.findtext(T('v')) or 0),'IF('+reference()+','+o+','+f+')',int(c.get('s')))
   changes[p]=s.finish()
  temp=BM.with_suffix('.cashguard.xlsx')
  with ZipFile(temp,'w',ZIP_DEFLATED,compresslevel=6) as out:
   for item in z.infolist():out.writestr(copy.copy(item),changes.get(item.filename,z.read(item.filename)))
 temp.replace(BM);print('Early and closing fees included in liquidity/dividend guards',flush=True)
if __name__=='__main__':apply()
