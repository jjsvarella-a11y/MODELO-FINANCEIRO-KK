"""Formula-only extension to the built workbook; full external recalc must follow.
Generated from the same helper as future builds, not hard-coded financial results.
"""
from engine import *
from v04_formula_helpers import projected_bank_formula,terminal_assets_formula
import zipfile,xml.etree.ElementTree as E,os
from returns10 import returns10_updates,payback_formula
P=ROOT/'entrega/KK01_Modelo_Integrado_V04.xlsx';meta=json.loads((ROOT/'src/v04_risk_metadata.json').read_text());M=json.loads((ROOT/'src/build_metadata.json').read_text());pr=M['param_rows'];pv=lambda k:"'Pressupostos'!$B$"+str(pr[k]);updates={}
def add(sn,cell,f):updates.setdefault(sn,{})[cell]=f
add('V04 Completion','B8',projected_bank_formula('Fluxo 180M','CAPEX',pv))
add('V04 Completion','B21',f'IF(COUNT(C5:C8)<4,"Pendente — recursos não documentados",IF({pv("completion_cost_validated")}=0,"Pendente — custos não validados",IF({pv("completion_timing_validated")}=0,"Pendente — calendário não validado",IF(C17>=1,"Cobertura agregada documentada no cenário","Alerta — défice conclusão"))))')
for item in meta:
 fs=item['flow'];hs=item['hotel'];pol=item['policy'];wf=item['waterfall'];cap=item['capex'];extra=item['params'];end=f'{31+extra.get("construction_delay",0)}+{pv("construction_delay")}+{pv("opening_delay")}';lim=str(extra['limit']) if 'limit' in extra else pv('limit');av=f'{pv("bank_available")}*{extra.get("bank_available",1)}'
 add(pol,'B16',terminal_assets_formula(fs,hs,pv));add(pol,'B17',f'MIN(\'{fs}\'!AX184,B16)');add(pol,'B18',f'MAX(0,\'{fs}\'!AX184-B17)')
 # Existing aggregate coverage includes other cash uses; only its bank component is replaced.
 row=item['row'];oldbank=f'IF(AND({pv("bank_cp")}=1,24<{pv("draw_end")}),MAX(0,({lim})*({av})-\'{fs}\'!AC28),0)'
 # Apply after reading source XML, preserving the complete rest of the coverage formula.
 item['_old_bank']=oldbank;item['_new_bank']=projected_bank_formula(fs,cap,pv,end,lim,av)
for name in ['V04 Retornos']:
 add(name,'B41',"SUM('Fluxo 180M'!AV5:AW184)+'V04 Politica'!B17")
updates.setdefault('V04 Retornos',{}).update(returns10_updates(pv))
add('KPI','B'+str(M['kpi_rows']['investor_payback']),payback_formula())
add('DASH Acionistas','B8',"'V04 Retornos'!B55");add('DASH Acionistas','B10',"'V04 Retornos'!B56");add('DASH Acionistas','B18',"'V04 Riscos'!N5")
N='http://schemas.openxmlformats.org/spreadsheetml/2006/main';tag=lambda n:'{'+N+'}'+n
with zipfile.ZipFile(P) as zi,zipfile.ZipFile(P.with_suffix('.tmp'),'w',zipfile.ZIP_DEFLATED) as zo:
 wb=E.fromstring(zi.read('xl/workbook.xml'));rs=E.fromstring(zi.read('xl/_rels/workbook.xml.rels'));targets={r.attrib['Id']:r.attrib['Target'] for r in rs};paths={}
 for s in wb.find(tag('sheets')):
  t=targets[s.attrib['{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id']];paths[s.attrib['name']]=t.lstrip('/') if t.startswith('/') else 'xl/'+t
 changes_by_path={paths[k]:v for k,v in updates.items()}
 for inf in zi.infolist():
  data=zi.read(inf.filename)
  needed=inf.filename in changes_by_path or inf.filename==paths['V04 Riscos'] or any(inf.filename==paths[x['waterfall']] for x in meta)
  if needed:
   x=E.fromstring(data);sd=x.find(tag('sheetData'));cells={c.attrib['r']:c for c in sd.iter(tag('c'))}
   changes=changes_by_path.get(inf.filename,{}).copy()
   if inf.filename==paths['V04 Riscos']:
    for item in meta:
     addr='H'+str(item['row']);c=cells[addr];f=c.find(tag('f')).text;changes[addr]=f.replace(item['_old_bank'],item['_new_bank'])
   for item in meta:
    if inf.filename==paths[item['waterfall']]:
     for rr in range(5,185):
      addr='Y'+str(rr);f=cells[addr].find(tag('f')).text;changes[addr]=f.replace(f"IF(A{rr}=180,'{item['flow']}'!AX{rr},0)",f"IF(A{rr}=180,'{item['policy']}'!B17,0)")
   styles={key:cells.get(addr).attrib.get('s') if cells.get(addr) is not None else None for key,addr in [('pct','B27'),('mult','B28'),('eur','B29'),('text','A27')]}
   for addr,value in changes.items():
    kind,f,fmt=value if isinstance(value,tuple) else ('formula',value,None)
    c=cells.get(addr)
    if c is None:
     import re
     rn=int(re.search(r'\d+',addr).group());row=sd.find(tag('row')+'[@r="'+str(rn)+'"]')
     if row is None:
      row=E.Element(tag('row'),{'r':str(rn)});at=next((i for i,e in enumerate(sd) if int(e.attrib['r'])>rn),len(sd));sd.insert(at,row)
     c=E.SubElement(row,tag('c'),{'r':addr})
    c.attrib.pop('t',None)
    for e in list(c):c.remove(e)
    if fmt in styles and styles[fmt] is not None:c.set('s',styles[fmt])
    if kind=='text':c.set('t','inlineStr');E.SubElement(E.SubElement(c,tag('is')),tag('t')).text=f
    else:E.SubElement(c,tag('f')).text=f
   if inf.filename==paths['V04 Retornos']:
    cols=x.find(tag('cols'))
    if cols is None:
     cols=E.Element(tag('cols'));at=list(x).index(sd);x.insert(at,cols)
    # Add a specific13..26hidden range after generic inherited widths.
    E.SubElement(cols,tag('col'),{'min':'13','max':'26','width':'14','customWidth':'1','hidden':'1'})
   data=E.tostring(x,encoding='utf-8',xml_declaration=True)
  zo.writestr(inf,data)
os.replace(P.with_suffix('.tmp'),P);print('Fórmulas completadas: disponibilidade condicionada e recuperação do crédito limitada aos ativos de saída.')
