"""Preserve eight original operating aggregates displaced by the central heatmap.
Restores GOP/fees/reserve display dependencies without altering the motor.
"""
from presentation_review_v08 import *
def protect_helpers(source,target,reference):
 src=Source(source);ref=Source(reference);changes={};state=Sheet(E.fromstring(src.z.read(src.pp[STATUS7])),{});sn='DASH 05 V06';moved={}
 for addr in ['M71','M72','M74','M76','M78','M80','M82','M84']:
  old=ref.cells(sn).get(addr)
  if old is None:continue
  dest='BM'+str(2000+rc(addr)[0]+1)
  if dest in state.cells:raise ValueError('Operating aggregate archive destination occupied')
  node=deepcopy(old);node.set('r',dest);f=node.find(T('f'))
  if f is not None and f.text:
   out=[]
   for token in Tokenizer('='+f.text).items:
    value=token.value
    if token.type=='OPERAND' and token.subtype=='RANGE' and '!' not in value and re.fullmatch(r'\$?[A-Z]+\$?\d+(?::\$?[A-Z]+\$?\d+)?',value):value=q(sn,value)
    out.append(value)
   f.text=''.join(out)
  state.row(2000+rc(addr)[0]+1).append(node);state.cells[dest]=node;state.row(2000+rc(addr)[0]+1).set('hidden','1');moved[addr]=dest
 def update(f):
  def sub(match):
   addr=match[1].replace('$','');return q(STATUS7,moved[addr]) if addr in moved else match[0]
  return re.sub(r"'DASH 05 V06'!(\$?M\$?\d+)(?![0-9A-Z])",sub,f)
 proof=[]
 for name,p in src.pp.items():
  raw=src.z.read(p)
  if b"'DASH 05 V06'!" not in raw and name!=sn:continue
  root=E.fromstring(raw);dirty=False
  for c in root.iter(T('c')):
   f=c.find(T('f'))
   if f is None or not f.text:continue
   before=f.text;f.text=update(before)
   if name==sn and not (57<=rc(c.get('r'))[0]<=87 and 11<=rc(c.get('r'))[1]<=17):
    out=[]
    for t in Tokenizer('='+f.text).items:
     text=t.value
     if t.type=='OPERAND' and t.subtype=='RANGE' and text.replace('$','') in moved:text=q(STATUS7,moved[text.replace('$','')])
     out.append(text)
    f.text=''.join(out)
   if f.text!=before:dirty=True;proof.append({'sheet':name,'cell':c.get('r'),'before':before,'after':f.text})
  if dirty:changes[p]=Sheet(root,{}).finish()
 changes[src.pp[STATUS7]]=state.finish();rewrite(source,target,changes);src.z.close();ref.z.close();return {'aggregates_preserved':moved,'references_corrected':proof}
if __name__=='__main__':
 p=argparse.ArgumentParser();p.add_argument('source',type=Path);p.add_argument('target',type=Path);p.add_argument('--reference',type=Path,required=True);p.add_argument('--proof',type=Path,required=True);a=p.parse_args();a.proof.write_text(json.dumps(protect_helpers(a.source,a.target,a.reference),ensure_ascii=False,indent=2))
