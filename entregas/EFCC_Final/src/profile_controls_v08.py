"""Expose operational/operator profiles and bank status on the primary dashboards.
Keep categorical values categorical and implement explicit Custom operating profile.
"""
from release_v08 import *
def profiles(source,target,mapfile,proof_file):
 src=Source(source);styles=add_styles(src.styles);roots={};mapping=json.loads(mapfile.read_text());defs={d['key']:d for d in mapping['definitions']};proof=[]
 def sh(sn):
  if sn not in roots:roots[sn]=Sheet(E.fromstring(src.z.read(src.pp[sn])),styles)
  return roots[sn]
 def dv(s,addr,options):
  dvs=s.r.find(T('dataValidations'))
  if dvs is None:dvs=E.SubElement(s.r,T('dataValidations'),count='0')
  for d in list(dvs):
   if addr in d.get('sqref','').split():dvs.remove(d)
  if options:
   n=E.SubElement(dvs,T('dataValidation'),type='list',sqref=addr,showErrorMessage='1',errorStyle='stop');E.SubElement(n,T('formula1')).text='"'+','.join(options)+'"'
 def bind(sn,key,row,pt,en,options=None):
  sheet=sh(sn);sheet.text('Q'+str(row-1)+':T'+str(row-1),pt,en,'label');sheet.text('U'+str(row-1)+':Y'+str(row-1),pt+' — Comparativo',en+' — Comparative','label')
  for mode,c,e in [('Base','Q','T'),('Comparative','U','Y')]:
   rr=ROWS[key];cached=src.get(store(mode),'B'+str(rr))[1];addr=c+str(row);sheet.merge(addr+':'+e+str(row+1),cached,style='input_text');sh(store(mode)).put('B'+str(rr),cached,q(sn,addr));dv(sheet,addr,options);defs[key]['base_panel' if mode=='Base' else 'comparative_panel']={'sheet':sn,'cell':addr};proof.append({'key':key,'mode':mode,'dashboard':sn,'cell':addr})
 # String inputs are categories or free text, not numeric inputs. Custom is
 # retained literally; numeric Custom continues to use its validated C column.
 for mode in ['Base','Comparative']:
  s=sh(store(mode))
  for definition in DEFS:
   if isinstance(definition['initial'],str) and definition['key'] not in ALIASES:
    rr=ROWS[definition['key']];s.put('D'+str(rr),src.get(store(mode),'D'+str(rr))[1],'B'+str(rr))
 sn='DASH 05 V06';oper=IRO[(sn,'H23')];operator=IRO[(sn,'H32')]
 bind(sn,oper,24,'Cenário operacional — Base','Operating scenario — Base',['Conservative','Base Case','Favorable','Upside Stress','Custom'])
 bind(sn,operator,28,'Operadora — Base','Operator — Base',['Operator Proposed Terms','KK Target Terms','KK Optimized Terms','Custom Terms'])
 s=sh(sn);s.text('Q22:Y22','Perfis operacionais e da operadora','Operating and operator profiles','section')
 for row,key,pt,en,kind in [(32,'adr','ADR ativo','Active ADR','adr'),(36,'occ','Ocupação ativa','Active occupancy','pct')]:
  s.text('Q'+str(row-1)+':T'+str(row-1),pt+' — Base',en+' — Base','label');s.text('U'+str(row-1)+':Y'+str(row-1),pt+' — Comparativo',en+' — Comparative','label')
  for mode,c,e in [('Base','Q','T'),('Comparative','U','Y')]:s.merge(c+str(row)+':'+e+str(row+1),src.get(store(mode),'D'+str(ROWS[key]))[1],inp(mode,key),kind)
 for mode,c,e in [('Base','Q','T'),('Comparative','U','Y')]:s.merge(c+'39:'+e+'40',src.get(store(mode),'I9')[1],q(store(mode),'I9'),'note')
 for mode in ['Base','Comparative']:
  st=sh(store(mode))
  for key in ['adr','occ']:
   rr=ROWS[key];old,cached=src.get(store(mode),'D'+str(rr));st.put('D'+str(rr),cached,'IF(AND(B'+str(rr)+'="Profile",'+inp(mode,oper)+'="Custom"),C'+str(rr)+','+old+')')
 sn='DASH 03 V06';bind(sn,'bank_status',24,'Natureza das condições — Base','Terms status — Base',['Hypothetical','Bank proposal','Approved']);bind(sn,'bank_institution',29,'Instituição — Base','Institution — Base');bind(sn,'bank_evidence',34,'Referência documental — Base','Documentary reference — Base');s=sh(sn);s.text('Q22:Y22','Natureza e identificação das condições','Terms status and identification','section');s.text('Q38:Y40','Aprovação exige evidência documental. Hipóteses não são compromissos bancários.','Approval requires documentary evidence. Assumptions are not bank commitments.','note')
 for sn in PANELS:
  s=sh(sn);s.text('E22:J22','BASE DE REFERÊNCIA — EDITÁVEL','EDITABLE REFERENCE BASE','section');s.text('K22:P22','COMPARATIVO — EDITÁVEL','EDITABLE COMPARATIVE','section')
 changes={src.pp[sn]:s.finish() for sn,s in roots.items()};changes['xl/styles.xml']=xml(src.styles);rewrite(source,target,changes);src.z.close();mapfile.write_text(json.dumps(mapping,ensure_ascii=False,indent=2));proof_file.write_text(json.dumps(proof,ensure_ascii=False,indent=2))
if __name__=='__main__':
 p=argparse.ArgumentParser();p.add_argument('source',type=Path);p.add_argument('target',type=Path);p.add_argument('--map',type=Path,required=True);p.add_argument('--proof',type=Path,required=True);a=p.parse_args();profiles(a.source,a.target,a.map,a.proof)
