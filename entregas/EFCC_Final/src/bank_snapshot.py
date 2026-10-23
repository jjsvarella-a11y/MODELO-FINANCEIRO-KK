"""Publish actual Base/Test engine results inside the issued workbook.
Static certified comparison values are guarded against changed active assumptions.
No second engine, no manually populated live financial indicators.
"""
from bank_r6 import *
from zip_copy import rewrite

def stamp(snapshot,data,base):
 source=Source(snapshot);expected=[]
 for sn,addresses in [(DAT,['B'+str(P[x[0]]) for x in PARAMS]),(ACT,['B'+str(i) for i in range(5,5+len(INPUTS))]),('V06 Registo Ativo',['B65']),('V06 Cenarios',['M5','K5','F5','J5']),(BN,['H33','N33']),(CENTER,['H122'])]:
  for addr in addresses:
   value=base['active_profile'] if sn==BN and addr=='N33' else source.get(sn,addr)[1];ref=q(sn,addr)
   if isinstance(value,(int,float)):expr='ABS('+ref+'-('+str(value)+'))<'+str(max(1e-9,abs(value)*1e-12))
   elif isinstance(value,str):expr=ref+'="'+value.replace('"','""')+'"'
   else:continue
   expected.append(expr)
 formula='AND('+','.join(expected)+')';assert len(formula)<8192
 with ZipFile(snapshot) as z:
  wb,pp=paths(z);texts=[''.join(n.itertext()) for n in E.fromstring(z.read('xl/sharedStrings.xml'))];styles={}
  for c in E.fromstring(z.read(pp[ACT])).find(T('sheetData'))[0]:
   if c.get('t')=='s':
    name=texts[int(c.findtext(T('v')))];
    if name in ['text','label','head','section','money','pct','ratio','number','note','input']:styles[name]=int(c.get('s'))
  dn=wb.find(T('definedNames'))
  for old in list(dn):
   if old.get('name')=='KK6_Emission_Current':dn.remove(old)
  E.SubElement(dn,T('definedName'),name='KK6_Emission_Current').text=formula
  s=Sheet(E.fromstring(z.read(pp[BN])),styles);store=Sheet(E.fromstring(z.read(pp[DAT])),styles)
  s.put('N33',base['active_profile'])
  for key,*_ in PARAMS:
   addr='E'+str(P[key]);cell=store.cells[addr];store.put(addr,base['params'][key],cell.findtext(T('f')),int(cell.get('s','0')))
  s.text('E586:Y587','Comparação bancária emitida e reconciliada','Issued and reconciled bank comparison','section')
  s.merge('E588:Y590','Snapshot válido','IF(KK6_Emission_Current,'+tr('Snapshot válido · ','Valid snapshot · ')+'&"'+data['id']+'",'+tr('DESATUALIZADO — recalcular e emitir novamente','OUTDATED — recalculate and issue again')+')','note')
  for area,label in [('E592:L593','KPI'),('M592:P593',base['active_profile']),('Q592:T593',data['active_profile']),('U592:Y593','Delta')]:s.merge(area,label,style='head')
  for j,(key,value) in enumerate(data['metrics'].items()):
   rr=594+j*2;dr=60+j;v=base['metrics'][key];delta=value-v if isinstance(value,(int,float)) and isinstance(v,(int,float)) else 'N/A';kind='pct' if 'IRR' in key or 'annual financing cost' in key else 'ratio' if key in ['Minimum DSCR','MOIC'] else 'number' if 'month' in key else 'money'
   store.put('C'+str(dr),key);store.put('D'+str(dr),v);store.put('E'+str(dr),value);store.put('F'+str(dr),delta);s.merge(f'E{rr}:L{rr+1}',key,style='label')
   for area,cc in [(f'M{rr}:P{rr+1}','D'),(f'Q{rr}:T{rr+1}','E'),(f'U{rr}:Y{rr+1}','F')]:s.merge(area,{'D':v,'E':value,'F':delta}[cc],'IF(KK6_Emission_Current,'+q(DAT,cc+str(dr))+',"OUTDATED")',kind)
  rr=634;s.merge(f'E{rr}:Y{rr+3}',data['calculated_at']+' | '+data['recipient']+' | '+data['params']['status'],style='note')
  center=Sheet(E.fromstring(z.read(pp[CENTER])),styles);center.merge('E126:P127','Comparação emitida →','HYPERLINK("#\''+BN+'\'!E586",'+tr('Comparação emitida →','Issued comparison →')+')','input')
  changes={pp[BN]:s.finish(),pp[DAT]:store.finish(),pp[CENTER]:center.finish(),'xl/workbook.xml':xml(wb)}
 tmp=Path(snapshot).with_suffix('.stamped.xlsx');rewrite(snapshot,tmp,changes);tmp.replace(snapshot);return {'signature_conditions':len(expected),'signature_formula_length':len(formula),'comparison_metrics':len(data['metrics'])}
