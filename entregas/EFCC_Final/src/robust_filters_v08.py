"""Type-aware input hardening, preserving economic values, histories and engine formulas.
Existing textual selections are retained for raw-history compatibility. New presets
use native numbers via named list ranges. No blanket VALUE replacement.
"""
from editable_base_v07 import *
from decimal import Decimal
import argparse

def native_option(value):
 if not isinstance(value,str):return value
 try:return float(Decimal(value[:-1])/100) if value.endswith('%') else float(Decimal(value))
 except Exception:return value

def harden(source,target):
 src=Source(source);changes={};proof={'numeric_formulas':[],'native_lists':[],'selected_values_changed':0,'history_changed':False,'engine_changed':False}
 roots={}
 def sh(sn):
  if sn not in roots:roots[sn]=Sheet(E.fromstring(src.z.read(src.pp[sn])),{})
  return roots[sn]
 for mode in ['Base','Comparative']:
  sn=store(mode);s=sh(sn)
  for dct in DEFS:
   rr=ROWS[dct['key']];addr='D'+str(rr);old,cached=src.get(sn,addr)
   if not isinstance(dct['initial'],(int,float)) or not old or not re.search(r'(?<![A-Z_])VALUE\(',old):continue
   # Decimal point and absence of a grouping separator are the canonical
   # representation used by the pre-existing preset text. Never parse categories.
   new=re.sub(r'(?<![A-Z_])VALUE\(SUBSTITUTE\((B\d+),"%",""\)\)',r'_xlfn.NUMBERVALUE(SUBSTITUTE(\1,"%",""),".",",")',old)
   new=re.sub(r'(?<![A-Z_])VALUE\((B\d+)\)',r'_xlfn.NUMBERVALUE(\1,".",",")',new)
   assert new!=old and 'IF(B'+str(rr)+'="Custom",C'+str(rr) in new
   s.put(addr,cached,new);proof['numeric_formulas'].append({'sheet':sn,'cell':addr,'key':dct['key'],'before':old,'after':new})
 # Keep the same 238 sheets: reusable native lists live in the existing status
 # support sheet. Named ranges permit standard cross-sheet validation in Excel.
 wb=src.w;names=wb.find(T('definedNames'))
 if names is None:names=E.Element(T('definedNames'));wb.insert(list(wb).index(wb.find(T('calcPr'))),names)
 status=sh(STATUS7);occupied=set(status.cells)
 for i,(sn,panel) in enumerate(PANELS.items()):
  sheet=sh(sn)
  for j,(key,pt,en,options) in enumerate(panel):
   letter=col(26+i*8+j);start=1000;name='KK08_'+('Bank_' if i==0 else 'Hotel_')+key
   for n,option in enumerate(options,start):
    addr=letter+str(n)
    if addr in occupied:raise ValueError('Native-list destination is occupied: '+addr)
    status.put(addr,native_option(option),style=int(sh(BASE7).cells['D'+str(ROWS[('bank_' if i==0 else '')+key])].get('s','0')))
   for prior in list(names):
    if prior.get('name')==name:names.remove(prior)
   E.SubElement(names,T('definedName'),name=name).text=q(STATUS7,'$'+letter+'$'+str(start)+':$'+letter+'$'+str(start+len(options)-1))
   dvs=sheet.r.find(T('dataValidations'))
   for mode,prefix in [('Base','H'),('Comparative','N')]:
    addr=prefix+str(23+j*2);matches=[dv for dv in dvs if addr in dv.get('sqref','').split()]
    if len(matches)!=1:raise ValueError('Expected unique preset validation: '+sn+'!'+addr)
    matches[0].find(T('formula1')).text=name
   proof['native_lists'].append({'sheet':sn,'parameter':key,'name':name,'options':[native_option(v) for v in options]})
 dim=status.r.find(T('dimension'));dim.set('ref','A1:AQ1014')
 for sn,s in roots.items():changes[src.pp[sn]]=s.finish()
 changes['xl/workbook.xml']=xml(wb)
 rewrite(source,target,changes);src.z.close()
 proof['source_sha256']=hashlib.sha256(Path(source).read_bytes()).hexdigest();proof['target_sha256']=hashlib.sha256(Path(target).read_bytes()).hexdigest()
 return proof
if __name__=='__main__':
 p=argparse.ArgumentParser();p.add_argument('source',type=Path);p.add_argument('target',type=Path);p.add_argument('--proof',type=Path,required=True);a=p.parse_args();a.proof.write_text(json.dumps(harden(a.source,a.target),ensure_ascii=False,indent=2))
