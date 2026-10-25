"""Place Project Impact Summary after controls on nine existing dashboards.
Retains hidden canonical UI values in existing support sheet and rewires their
actual dependents; no financial engine calculation is replaced.
"""
from navigation_v08 import *
from openpyxl.formula.tokenizer import Tokenizer
from copy import deepcopy

def ui_archive_address(sn,addr):
 i=[x[0] for x in DASHES[:9]].index(sn);rr,cc=rc(addr);return col(cc+52)+str(rr+1101+i*100)

def apply_impact(source,target):
 src=Source(source);changes={};styles=add_styles(src.styles);roots={};mapping={};proof={'relocated_ui_cells':[],'rewired_formula_cells':[]}
 status=Sheet(E.fromstring(src.z.read(src.pp[STATUS7])),styles)
 for i,(sn,*_) in enumerate(DASHES[:9]):
  root=E.fromstring(src.z.read(src.pp[sn]));s=Sheet(root,styles);roots[sn]=s
  for addr,node in list(s.cells.items()):
   rr,cc=rc(addr)
   if 41<=rr<=52 and 4<=cc<=15:
    new=ui_archive_address(sn,addr)
    if new in status.cells:raise ValueError('UI-preservation destination occupied')
    mapping[(sn,addr)]=new
 def remap_qualified(f):
  def sub(m):
   key=(m[1],m[2].replace('$',''))
   if key in mapping:return q(STATUS7,mapping[key])
   return m[0]
  return re.sub(r"'([^']+)'!(\$?[A-Z]+\$?\d+)(?![0-9A-Z])",sub,f)
 def relocate_formula(f,oldsn):
  tokens=Tokenizer('='+f).items;out=[]
  for t in tokens:
   text=t.value
   if t.type=='OPERAND' and t.subtype=='RANGE':
    if '!' not in text:
     match=re.fullmatch(r'(\$?[A-Z]+\$?\d+)(?::(\$?[A-Z]+\$?\d+))?',text)
     if match:
      start=match[1].replace('$','');end=match[2].replace('$','') if match[2] else None
      if (oldsn,start) in mapping and (end is None or (oldsn,end) in mapping):text=q(STATUS7,mapping[(oldsn,start)]+(':'+mapping[(oldsn,end)] if end else ''))
      else:text=q(oldsn,text)
    text=remap_qualified(text)
   out.append(text)
  return ''.join(out)
 for (sn,addr),new in mapping.items():
  old=roots[sn].cells[addr];node=deepcopy(old);node.set('r',new);f=node.find(T('f'))
  if f is not None and f.text:f.text=relocate_formula(f.text,sn)
  status.row(rc(new)[0]+1).append(node);status.cells[new]=node
  status.row(rc(new)[0]+1).set('hidden','1');proof['relocated_ui_cells'].append({'source':sn+'!'+addr,'preserved_at':STATUS7+'!'+new})
 # Rewire qualified references across all parts that contain an affected pattern.
 # No source monthly engine contains such references in the audited version.
 for sn,p in src.pp.items():
  raw=src.z.read(p)
  if not any(("'"+old+"'!").encode() in raw for old,*_ in DASHES[:9]):continue
  s=roots.get(sn) or Sheet(E.fromstring(raw),styles);dirty=False
  for addr,c in s.cells.items():
   f=c.find(T('f'))
   if f is not None and f.text:
    new=remap_qualified(f.text)
    if new!=f.text:f.text=new;dirty=True;proof['rewired_formula_cells'].append(sn+'!'+addr)
  if dirty:roots[sn]=s
 for sn,pt,en in DASHES[:9]:
  s=roots[sn]
  s.merge('Q22:Y41',style='bg')
  s.text('E42:Y42','Impacto Global do Cenário no Projeto — Base | Comparativo | Δ','Project Impact Summary — Base | Comparative | Δ','section')
  s.text('E43:M43','Indicador','Metric','table_head');s.text('N43:Q43','Base atual','Current Base','table_head');s.text('R43:U43','Comparativo','Comparative','table_head');s.text('V43:Y43','Delta','Delta','table_head')
  items=[(6,'CAPEX futuro','Future CAPEX','money'),(7,'Pico de dívida','Peak debt','money'),(8,'Folga mínima sobre reservas','Minimum liquidity headroom','money'),(9,'Dividendos previstos','Planned dividends','money'),(10,'VAL novos investidores','New investor NPV','money'),(11,'TIR novos investidores','New investor IRR','pct')]
  for row,(mr,pt,en,kind) in enumerate(items,44):
   s.text('E'+str(row)+':M'+str(row),pt,en,'label')
   for col0,col1,mc in [('N','Q','B'),('R','U','C'),('V','Y','D')]:s.merge(col0+str(row)+':'+col1+str(row),src.get(MET7,mc+str(mr))[1],q(MET7,mc+str(mr)),kind)
  s.merge('E50:Y51',src.get(IMP,'B18')[1],q(IMP,'B18'),'note')
  for rr in range(42,52):s.row(rr).attrib.pop('hidden',None);s.row(rr).set('ht','19')
  s.row(52).set('hidden','1');s.row(53).set('hidden','1')
  changes[src.pp[sn]]=s.finish()
 # Include rewired print view dependencies while keeping the original print layout.
 for sn,s in roots.items():
  if sn not in [x[0] for x in DASHES[:9]]:changes[src.pp[sn]]=s.finish()
 dim=status.r.find(T('dimension'));dim.set('ref','A1:BP2000');changes[src.pp[STATUS7]]=status.finish();changes['xl/styles.xml']=xml(src.styles)
 rewrite(source,target,changes);src.z.close();return proof
if __name__=='__main__':
 p=argparse.ArgumentParser();p.add_argument('source',type=Path);p.add_argument('target',type=Path);p.add_argument('--proof',type=Path,required=True);a=p.parse_args();a.proof.write_text(json.dumps(apply_impact(a.source,a.target),ensure_ascii=False,indent=2))
