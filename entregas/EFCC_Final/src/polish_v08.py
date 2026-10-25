"""Uniform sidebars/frozen menus and compact four-card reporting center."""
from direct_inputs_v08 import *
def polish(source,target):
 src=Source(source);styles=add_styles(src.styles);changes={};short=[('Executivo','Executive'),('Construção','Construction'),('Fiscalidade','Taxation'),('Financiamento','Financing'),('Real Estate','Real Estate'),('Residencial','Residential'),('Hotel & MICE','Hotel & MICE'),('Acionistas','Shareholding'),('Retornos','Returns'),('Relatórios','Reports')]
 for idx,(sn,pt,en) in enumerate(DASHES):
  s=Sheet(E.fromstring(src.z.read(src.pp[sn])),styles);s.merge('A7:D150',style='side')
  for n,(dest,dpt,den) in enumerate(DASHES):
   row=22+n*4;active=n==idx;st='side_active' if active else 'side';num='side_num_active' if active else 'side_num'
   s.merge('A'+str(row)+':B'+str(row+3),f'{n+1:02d}',f'HYPERLINK("#\'{dest}\'!A1","{n+1:02d}")',num);s.merge('C'+str(row)+':D'+str(row+3),dpt,f'HYPERLINK("#\'{dest}\'!A1",'+tr(dpt,den)+')',st)
  for row in range(7,22):s.row(row).set('hidden','1')
  for row in range(22,62):s.row(row).attrib.pop('hidden',None);s.row(row).set('ht','17')
  # Hide only the obsolete duplicate PIS, not the detailed budget/investor tables.
  if sn!=CENTER and str(src.get(sn,'E150')[1] or '').startswith('Impacto Global'):
   for row in range(150,181):s.row(row).set('hidden','1')
  cols=s.r.find(T('cols'))
  for node in list(cols):
   if int(node.get('min'))<=4:
    if int(node.get('max'))>4:node.set('min','5')
    else:cols.remove(node)
  for c,w in [(1,2),(2,6),(3,32),(4,2)]:E.SubElement(cols,T('col'),min=str(c),max=str(c),width=str(w),customWidth='1')
  cols[:]=sorted(cols,key=lambda n:int(n.get('min')))
  for view in s.r.findall('.//'+T('sheetView')):
   for n in list(view):
    if E.QName(n).localname in ['pane','selection']:view.remove(n)
   view.insert(0,E.Element(T('pane'),xSplit='4',ySplit='6',topLeftCell='E7',activePane='bottomRight',state='frozen'));view.append(E.Element(T('selection'),pane='bottomRight',activeCell='E22',sqref='E22'));view.set('showGridLines','0');view.set('tabSelected','1' if idx==0 else '0')
  if sn==CENTER:
   vals={r:src.get(REPORT7,'B'+str(r))[1] for r in range(6,13)};s.merge('E7:Y140',style='bg')
   for area,pt,en in [('E22:I28','One-Page Executive','One-Page Executive'),('J22:N28','Bank Financing','Bank Financing'),('O22:S28','Equity Investor','Equity Investor'),('T22:Y28','Comprehensive Project Report','Comprehensive Project Report')]:s.link(area,CENTER,'H31',pt,en)
   s.text('E30:P30','Configuração do relatório','Report configuration','section');s.text('Q30:Y30','Pré-visualização — Base guardada','Preview — saved Base','section')
   configs=[(6,31,'Formato','Format',['Executive One Page','Bank Financing','Equity Investor','Comprehensive Project']),(7,33,'Comparação','Comparison',['Base Only','Base vs Selected','Base vs Multiple Saved']),(8,35,'Idioma','Language',['PT','EN']),(9,37,'Destinatário','Recipient',None),(10,39,'IDs comparativos guardados','Saved comparative IDs',None),(11,41,'Confidencialidade','Confidentiality',['Confidential','Internal','Public']),(12,43,'Anexos','Appendices',['Yes','No'])]
   dvs=s.r.find(T('dataValidations'))
   if dvs is not None:s.r.remove(dvs)
   dvs=E.SubElement(s.r,T('dataValidations'),count='0');rsh=Sheet(E.fromstring(src.z.read(src.pp[REPORT7])),styles)
   for old,row,pt,en,options in configs:
    addr='H'+str(row);s.text('E'+str(row)+':G'+str(row+1),pt,en,'label');s.merge(addr+':P'+str(row+1),vals[old],style='input_text');rsh.put('B'+str(old),vals[old],q(CENTER,addr))
    if options:
     n=E.SubElement(dvs,T('dataValidation'),type='list',sqref=addr,showErrorMessage='1',errorStyle='stop');E.SubElement(n,T('formula1')).text='"'+','.join(options)+'"'
   s.merge('Q31:Y33',src.get(STATUS7,'B5')[1],q(STATUS7,'B5'),'note')
   for row,mr,pt,en,kind in [(35,6,'CAPEX futuro','Future CAPEX','money'),(37,7,'Pico de dívida','Peak debt','money'),(39,10,'VAL investidores','Investor NPV','money'),(41,11,'TIR investidores','Investor IRR','pct')]:s.text('Q'+str(row)+':U'+str(row+1),pt,en,'label');s.merge('V'+str(row)+':Y'+str(row+1),src.get(MET7,'B'+str(mr))[1],q(MET7,'B'+str(mr)),kind)
   s.merge('Q43:Y44',src.get(STATUS7,'B9')[1],q(STATUS7,'B9'),'note');s.text('E46:Y47','Pré-visualização financeira. Emitir PDF requer validação, Base guardada e snapshot. Não constitui aprovação bancária.','Financial preview. PDF issuance requires validation, saved Base and a snapshot. This is not bank approval.','note');s.link('E49:O51',REPORT7,'A1','Validar · Snapshot · Pré-visualizar · Emitir PDF →','Validate · Snapshot · Preview · Issue PDF →');s.link('P49:Y51',HIST7,'A1','Histórico e versões →','History and versions →');s.text('E54:Y57','Apenas a Base guardada determina os resultados principais dos relatórios. Comparações são incluídas quando selecionadas. Relatórios impressos não incluem sidebar.','The saved Base determines the main report results. Comparisons are included when selected. Printed reports have no sidebar.','note');changes[src.pp[REPORT7]]=rsh.finish()
   for row in range(30,62):s.row(row).attrib.pop('hidden',None)
  changes[src.pp[sn]]=s.finish()
 changes['xl/styles.xml']=xml(src.styles);rewrite(source,target,changes);src.z.close()
if __name__=='__main__':
 p=argparse.ArgumentParser();p.add_argument('source',type=Path);p.add_argument('target',type=Path);a=p.parse_args();polish(a.source,a.target)
