"""Refresh generated chart caches from independently recalculated support cells.
Original hand-adjusted hotel/bank chart XML is not regenerated.
"""
from view_consistency_v08 import *
def sync(source,target,proof_file):
 src=Source(source);records=json.loads(Path(proof_file).read_text());changes={};C='{http://schemas.openxmlformats.org/drawingml/2006/chart}'
 for record in records:
  sn=record['sheet'];idx=record['chart_index'];r=E.fromstring(src.z.read(src.pp[sn]));d=r.find(T('drawing'));rels=E.fromstring(src.z.read(relpath(src.pp[sn])));dp=resolve(src.pp[sn],next(x.get('Target') for x in rels if x.get('Id')==d.get('{'+NSR+'}id')));drawing=E.fromstring(src.z.read(dp));anchor=list(drawing)[idx];cn=next(x for x in anchor.iter() if E.QName(x).localname=='chart');cr=E.fromstring(src.z.read(relpath(dp)));cp=resolve(dp,next(x.get('Target') for x in cr if x.get('Id')==cn.get('{'+NSR+'}id')));chart=E.fromstring(src.z.read(cp))
  for parent in chart.iter():
   for refnode in list(parent):
    if refnode.tag not in [C+'numRef',C+'strRef']:continue
    formula=refnode.findtext(C+'f') or '';m=re.fullmatch(r"'([^']+)'!(\$?[A-Z]+\$?\d+)(?::(\$?[A-Z]+\$?\d+))?",formula)
    if not m:continue
    a=m[2].replace('$','');b=(m[3] or m[2]).replace('$','');r0,c0=rc(a);r1,c1=rc(b);values=[src.get(m[1],col(c)+str(r+1))[1] for r in range(r0,r1+1) for c in range(c0,c1+1)]
    string=refnode.tag==C+'strRef' or (parent.tag==C+'cat' and any(isinstance(x,str) for x in values))
    refnode.tag=C+('strRef' if string else 'numRef')
    for old in list(refnode):
     if old.tag in [C+'numCache',C+'strCache']:refnode.remove(old)
    cache=E.SubElement(refnode,C+('strCache' if string else 'numCache'))
    if not string:E.SubElement(cache,C+'formatCode').text='General'
    E.SubElement(cache,C+'ptCount',val=str(len(values)))
    for i,value in enumerate(values):
     if value is None or value=='':continue
     if not string and not isinstance(value,(int,float)):raise ValueError('Non-numeric chart value: '+repr(value))
     point=E.SubElement(cache,C+'pt',idx=str(i));E.SubElement(point,C+'v').text=str(value)
  count=record['support_rows'][1]
  if count>=30:
   for axis in chart.iter(C+'catAx'):
    skip=axis.find(C+'tickLblSkip')
    if skip is None:
     skip=E.Element(C+'tickLblSkip',val=str(4 if count>=60 else 2));before=next((i for i,n in enumerate(axis) if n.tag in [C+'tickMarkSkip',C+'noMultiLvlLbl',C+'extLst']),len(axis));axis.insert(before,skip)
    else:skip.set('val',str(4 if count>=60 else 2))
  changes[cp]=xml(chart)
 # Clear only the two displaced operating cells that were blank spacers in the
 # new heatmap, after their actual dependents were preserved in support cells.
 s=Sheet(E.fromstring(src.z.read(src.pp['DASH 05 V06'])),{});s.put('M78');s.put('M82');changes[src.pp['DASH 05 V06']]=s.finish()
 rewrite(source,target,changes);src.z.close()
if __name__=='__main__':
 p=argparse.ArgumentParser();p.add_argument('source',type=Path);p.add_argument('target',type=Path);p.add_argument('--charts-proof',type=Path,required=True);a=p.parse_args();sync(a.source,a.target,a.charts_proof)
