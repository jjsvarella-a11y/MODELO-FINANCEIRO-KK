"""Apply topic charts without touching the R4 Hotel drawing or engine expressions."""
from zipfile import ZipFile,ZIP_DEFLATED
from lxml import etree as E
import copy,json,posixpath
from assemble import paths,T,xml,resolve,relpath
from closure_r5 import MASTER,OUT,CON,CASH,INV,SALES,SN,Sheet,tr,col
D='http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing';C='http://schemas.openxmlformats.org/drawingml/2006/chart';R='http://schemas.openxmlformats.org/officeDocument/2006/relationships';P='http://schemas.openxmlformats.org/package/2006/relationships';CT='http://schemas.openxmlformats.org/package/2006/content-types'

def apply():
 with ZipFile(MASTER) as z:
  _,pp=paths(z);changes={};added={};proof=[]
  def drawing(sn):
   root=E.fromstring(changes.get(pp[sn]) or z.read(pp[sn]));rel=next(r for r in E.fromstring(z.read(relpath(pp[sn]))) if r.get('Type').endswith('/drawing'));p=resolve(pp[sn],rel.get('Target'));dr=E.fromstring(changes.get(p) or z.read(p));rp=relpath(p);rels=E.fromstring(changes.get(rp) or z.read(rp));chartrels={r.get('Id'):resolve(p,r.get('Target')) for r in rels if r.get('Type').endswith('/chart')};charts=[]
   for a in dr:
    node=a.find('.//{'+C+'}chart')
    if node is not None:charts.append(chartrels[node.get('{'+R+'}id')])
   return p,dr,rp,rels,charts
  for sn,module,titles in [('DASH 01 V06',CON,[('Orçamento por capítulo','Budget by chapter'),('Desembolso trimestral','Quarterly disbursement'),('Ponte orçamental','Budget bridge')]),('DASH 03 V06',CASH,[('Caixa e reservas','Cash and reserves'),('Dívida e serviço','Debt and service'),('Fontes e usos de caixa','Cash sources and uses')]),('DASH 06 V06',INV,[('Estrutura acionista','Ownership structure'),('Alocação de blocos','Blocks allocation'),('Captação de capital','Capital raise progress')]),('DASH 04A V06',SALES,[('Vendas por ano','Annual sales'),('Vendas e recebimentos','Sales and collections'),('Inventário disponível','Available inventory')])]:
   p,dr,rp,rels,targets=drawing(sn);sp,sd,srp,srels,sources=drawing(module);assert len(targets)==3
   for target,source in zip(targets,sources):changes[target]=z.read(source);proof.append({'dashboard':sn,'target_chart':target,'source_chart':source})
   root=E.fromstring(z.read(pp[sn]));marker=root.find('.//'+T('c')+'[@r="E57"]');s=Sheet(root,{'section':int(marker.get('s'))})
   for i,(pt,en) in enumerate(titles):s.text(f'{col(4+i*7)}57:{col(10+i*7)}57',pt,en,'section')
   if sn=='DASH 03 V06':
    # Each chart has a distinct package identity and relationship; no shared owners.
    number=max(int(n.rsplit('chart',1)[1][:-4]) for n in z.namelist() if n.startswith('xl/charts/chart') and n.endswith('.xml'));newpart=f'xl/charts/chart{number+1}.xml';added[newpart]=z.read(sources[3]);rid='rId'+str(max(int(r.get('Id')[3:]) for r in rels)+1);E.SubElement(rels,'{'+P+'}Relationship',Id=rid,Type=R+'/chart',Target=posixpath.relpath(newpart,posixpath.dirname(p)))
    anchor=copy.deepcopy(list(sd)[3]);anchor.find('.//{'+C+'}chart').set('{'+R+'}id',rid);nv=anchor.find('.//{'+D+'}cNvPr');nv.set('id','100');nv.set('name','Critical Liquidity');
    for label,c,r in [('from',4,399),('to',25,431)]:
     node=anchor.find('{'+D+'}'+label);node.find('{'+D+'}col').text=str(c);node.find('{'+D+'}row').text=str(r);node.find('{'+D+'}colOff').text='0';node.find('{'+D+'}rowOff').text='0'
    anchor.set('editAs','twoCell');dr.append(anchor);changes[p]=xml(dr);changes[rp]=xml(rels);ct=E.fromstring(z.read('[Content_Types].xml'));E.SubElement(ct,'{'+CT+'}Override',PartName='/'+newpart,ContentType='application/vnd.openxmlformats-officedocument.drawingml.chart+xml');changes['[Content_Types].xml']=xml(ct);s.text('E398:Y399','Folga sobre reservas — evolução e pontos críticos','Reserve headroom — evolution and critical points','section')
   changes[pp[sn]]=s.finish()
  # Projection filters return actual matching unit rows with classic scalar INDEX/SMALL.
  sn=SALES;root=E.fromstring(z.read(pp[sn]));s=Sheet(root,{})
  for rr in range(5,99):s.put('I'+str(rr),rr-4,'IF(H'+str(rr)+'=1,ROW()-4,1000+ROW())')
  af=root.find(T('autoFilter'))
  if af is None:E.SubElement(root,T('autoFilter'),ref='A4:I98')
  changes[pp[sn]]=s.finish()
  sn='DASH 04A V06';root=E.fromstring(changes[pp[sn]]);s=Sheet(root,{})
  for j in range(94):
   rr=j+220;k=j+1;idx='MATCH(SMALL(\''+SALES+'\'!$I$5:$I$98,'+str(k)+'),\''+SALES+'\'!$I$5:$I$98,0)'
   for target,c in [('E','A'),('H','B'),('K','C'),('M','D'),('O','E'),('S','F'),('V','G'),('Y','H')]:
    addr=target+str(rr);old=s.cells[addr];sty=int(old.get('s'));f='IF('+str(k)+'<=SUM(\''+SALES+'\'!$H$5:$H$98),INDEX(\''+SALES+'\'!$'+c+'$5:$'+c+'$98,'+idx+'),"")';s.put(addr,0 if k<=94 else '',f,sty)
  changes[pp[sn]]=s.finish()
  temp=MASTER.with_suffix('.topics.xlsx')
  with ZipFile(temp,'w',ZIP_DEFLATED,compresslevel=6) as out:
   for item in z.infolist():out.writestr(copy.copy(item),changes.get(item.filename,z.read(item.filename)))
   for name,data in added.items():out.writestr(name,data)
 temp.replace(MASTER);(OUT/'Topic_Charts.json').write_text(json.dumps({'status':'APPLIED','charts':proof,'Hotel_drawing_preserved':True,'additional_cash_chart':1},indent=2));print('Topic charts applied',flush=True)
if __name__=='__main__':apply()
