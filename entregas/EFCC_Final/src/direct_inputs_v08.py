"""Make remaining dashboard assumption fields actual editable Base inputs.
Comparative values retain their independent store; save/restore maps follow the
B-column input references rather than replacing those references with constants.
"""
from investor_v08 import *
def direct_inputs(source,target,map_target):
 src=Source(source);styles=add_styles(src.styles);changes={};roots={};bindings=[]
 def sh(sn):
  if sn not in roots:roots[sn]=Sheet(E.fromstring(src.z.read(src.pp[sn])),styles)
  return roots[sn]
 definitions=json.loads((OUT7/'Mapa_Cenarios.json').read_text());dd={d['key']:d for d in definitions['definitions']}
 for dct in DEFS:
  key=dct['key'];sn=dct.get('sheet');addr=dct.get('cell')
  if not key.startswith('I') or key in ALIASES or sn not in [x[0] for x in DASHES[:9]] or sn in PANELS:continue
  rr=ROWS[key];f,v=src.get(BASE7,'B'+str(rr));base=sh(BASE7);dash=sh(sn);dash.put(addr,v,style='input_pct' if dct.get('kind')=='pct' else 'input_text');base.put('B'+str(rr),v,q(sn,addr));bindings.append({'key':key,'mode':'Base','sheet':sn,'cell':addr})
  # Reuse the approved validation; do not invent domain ranges or options.
  candidates=[]
  for column in ['B','C']:
   for dv in base.r.findall('.//'+T('dataValidation')):
    if column+str(rr) in dv.get('sqref','').split():candidates.append(dv)
   if candidates:break
  dvs=dash.r.find(T('dataValidations'))
  if dvs is None:dvs=E.SubElement(dash.r,T('dataValidations'),count='0')
  for dv in list(dvs):
   if addr in dv.get('sqref','').split():dvs.remove(dv)
  if candidates:
   dv=deepcopy(candidates[0]);dv.set('sqref',addr);dvs.append(dv)
  dd[key]['base_panel']={'sheet':sn,'cell':addr}
 # Investor discount is also available directly, independently for each context.
 for mode,addr in [('Base','H29'),('Comparative','N25')]:
  sn='DASH 07 V06';s=sh(sn);r=ROWS['discount'];store_s=sh(store(mode));v=src.get(store(mode),'B'+str(r))[1];s.put(addr,v,style='input_pct');store_s.put('B'+str(r),v,q(sn,addr));bindings.append({'key':'discount','mode':mode,'sheet':sn,'cell':addr});dd['discount']['base_panel' if mode=='Base' else 'comparative_panel']={'sheet':sn,'cell':addr}
  dvs=s.r.find(T('dataValidations'))
  if dvs is None:dvs=E.SubElement(s.r,T('dataValidations'),count='0')
  for dv in list(dvs):
   if addr in dv.get('sqref','').split():dvs.remove(dv)
  dv=E.SubElement(dvs,T('dataValidation'),type='list',sqref=addr,showErrorMessage='1',errorStyle='stop');E.SubElement(dv,T('formula1')).text='"8%,10%,12%,Custom"'
  if mode=='Base':s.text('E29:G30','Desconto — Base','Discount — Base','label')
  else:s.text('K25:M26','Desconto — Comparativo','Discount — Comparative','label')
 for sn in {x['sheet'] for x in bindings}:
  s=sh(sn)
  if sn not in PANELS:s.text('E22:P22','Base editável · Comparativo independente nos Parâmetros','Editable Base · Independent Comparative in Assumptions','section')
 for sn,s in roots.items():changes[src.pp[sn]]=s.finish()
 changes['xl/styles.xml']=xml(src.styles);rewrite(source,target,changes);map_target.write_text(json.dumps(definitions,ensure_ascii=False,indent=2));src.z.close();return bindings
if __name__=='__main__':
 p=argparse.ArgumentParser();p.add_argument('source',type=Path);p.add_argument('target',type=Path);p.add_argument('--map',type=Path,required=True);p.add_argument('--proof',type=Path,required=True);a=p.parse_args();a.proof.write_text(json.dumps(direct_inputs(a.source,a.target,a.map),ensure_ascii=False,indent=2))
