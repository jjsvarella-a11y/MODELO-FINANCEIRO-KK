"""Uniform institutional headers, signed deltas and actual report configuration."""
from kpi_layout_v08 import *
def finish_ui(source,target):
 src=Source(source);styles=add_styles(src.styles);changes={}
 xfs=src.styles.find(T('cellXfs'));fills=src.styles.find(T('fills'));header=xfs[styles['header']];fills[int(header.get('fillId'))].find('.//'+T('fgColor')).set('rgb','FFC8101E')
 fmts=src.styles.find(T('numFmts'));num=max(int(x.get('numFmtId')) for x in fmts)+1
 for kind,code in [('money','+#,##0 "€";-#,##0 "€";–'),('pct','+0.0%;-0.0%;–')]:
  E.SubElement(fmts,T('numFmt'),numFmtId=str(num),formatCode=code);xf=deepcopy(xfs[styles[kind]]);xf.set('numFmtId',str(num));styles['delta_'+kind]=len(xfs);xfs.append(xf);num+=1
 for n in [fmts,xfs]:n.set('count',str(len(n)))
 for idx,(sn,pt,en) in enumerate(DASHES,1):
  s=Sheet(E.fromstring(src.z.read(src.pp[sn])),styles)
  s.merge('E2:V3','KK-01 | HOTEL & RESIDENCES',style='header');s.text('E4:Y4',f'{idx:02d} — '+pt,f'{idx:02d} — '+en,'sub')
  if idx<10:
   for rr in range(44,50):s.cells['V'+str(rr)].set('s',str(styles['delta_pct' if rr==49 else 'delta_money']))
  changes[src.pp[sn]]=s.finish()
 s=Sheet(E.fromstring(changes[src.pp[CENTER]]),styles);report=Sheet(E.fromstring(src.z.read(src.pp[REPORT7])),styles)
 configs=[(6,23,'Formato do relatório','Report format',['Executive One Page','Bank Financing','Equity Investor','Comprehensive Project']),(7,25,'Comparação','Comparison',['Base Only','Base vs Selected','Base vs Multiple Saved']),(8,27,'Idioma','Language',['PT','EN']),(9,29,'Destinatário','Recipient',None),(10,31,'IDs dos cenários guardados','Saved scenario IDs',None),(11,33,'Confidencialidade','Confidentiality',['Confidential','Internal','Public']),(12,35,'Incluir anexos','Include appendices',['Yes','No'])]
 for rr in range(22,40):s.row(rr).attrib.pop('hidden',None);s.row(rr).set('ht','17')
 s.text('E22:P22','Configuração do relatório','Report configuration','section')
 for oldrow,newrow,pt,en,options in configs:
  addr='H'+str(newrow);cached=src.get(REPORT7,'B'+str(oldrow))[1]
  s.text('E'+str(newrow)+':G'+str(newrow+1),pt,en,'label');s.merge(addr+':P'+str(newrow+1),cached,style='input_text')
  dvs=s.r.find(T('dataValidations'))
  if dvs is None:dvs=E.SubElement(s.r,T('dataValidations'),count='0')
  for old in list(dvs):
   if addr in old.get('sqref','').split():dvs.remove(old)
  if options:
   n=E.SubElement(dvs,T('dataValidation'),type='list',sqref=addr,showErrorMessage='1',errorStyle='stop');E.SubElement(n,T('formula1')).text='"'+','.join(options)+'"'
  report.put('B'+str(oldrow),cached,q(CENTER,addr))
  rdvs=report.r.find(T('dataValidations'))
  if rdvs is not None:
   for old in list(rdvs):
    if 'B'+str(oldrow) in old.get('sqref','').split():rdvs.remove(old)
 s.text('Q22:Y22','Pré-visualização — Base guardada','Preview — saved Base','section')
 s.merge('Q23:Y25',src.get(STATUS7,'B5')[1],q(STATUS7,'B5'),'note')
 for rr,mr,pt,en,kind in [(27,6,'CAPEX futuro','Future CAPEX','money'),(29,7,'Pico de dívida','Peak debt','money'),(31,10,'VAL investidores','Investor NPV','money'),(33,11,'TIR investidores','Investor IRR','pct')]:
  s.text('Q'+str(rr)+':U'+str(rr+1),pt,en,'label');s.merge('V'+str(rr)+':Y'+str(rr+1),src.get(MET7,'B'+str(mr))[1],q(MET7,'B'+str(mr)),kind)
 s.merge('Q35:Y39',src.get(STATUS7,'B9')[1],q(STATUS7,'B9'),'note')
 for area,pt,en in [('E8:I13','One-Page Executive','One-Page Executive'),('J8:N13','Bank Financing','Bank Financing'),('O8:S13','Equity Investor','Equity Investor'),('T8:Y13','Comprehensive Project Report','Comprehensive Project Report')]:s.link(area,CENTER,'H23',pt,en)
 s.link('E19:P20',CENTER,'H23','Configurar relatório, destinatário e comparação →','Configure report, recipient and comparison →')
 s.text('E38:P40','A pré-visualização é financeira, não um PDF emitido. Validar, congelar e emitir pela ferramenta de relatórios.','This financial preview is not an issued PDF. Validate, freeze and issue using the reporting tool.','note')
 changes[src.pp[CENTER]]=s.finish();changes[src.pp[REPORT7]]=report.finish()
 # Format native preset lists as numbers/percentages, including mixed Profile
 # and Custom entries, without altering the selected values or historical raw data.
 status=Sheet(E.fromstring(src.z.read(src.pp[STATUS7])),styles)
 for i,(sn,panel) in enumerate(PANELS.items()):
  for j,(key,*_) in enumerate(panel):
   dct=next(d for d in DEFS if d['key']==('bank_' if i==0 else '')+key);kind=dct.get('kind');fmt=styles['pct' if kind=='pct' else 'money' if kind=='money' else 'count' if isinstance(dct['initial'],(int,float)) else 'text']
   for rr in range(1000,1014):
    c=status.cells.get(col(26+i*8+j)+str(rr))
    if c is not None:c.set('s',str(fmt))
 changes[src.pp[STATUS7]]=status.finish()
 # In these categorical dropdowns Custom is a genuine engine mode, not a
 # request to replace the category with the dormant manual C value.
 for mode in ['Base','Comparative']:
  sn=store(mode);typed=Sheet(E.fromstring(src.z.read(src.pp[sn])),styles)
  for dct in DEFS:
   options=dct.get('options') or (dct.get('kind') if isinstance(dct.get('kind'),list) else [])
   if isinstance(dct['initial'],str) and 'Custom' in options:
    rr=ROWS[dct['key']];typed.put('D'+str(rr),src.get(sn,'D'+str(rr))[1],'B'+str(rr))
  changes[src.pp[sn]]=typed.finish()
 changes['xl/styles.xml']=xml(src.styles);rewrite(source,target,changes);src.z.close()
if __name__=='__main__':
 p=argparse.ArgumentParser();p.add_argument('source',type=Path);p.add_argument('target',type=Path);a=p.parse_args();finish_ui(a.source,a.target)
