"""Compact navigation and explicit versioning entry points, without financial edits."""
from scenario_v07 import *
def finish_ui(file=FINAL7):
 s=Source(file);changes={}
 def sh(sn):return Sheet(E.fromstring(s.z.read(s.pp[sn])),{})
 for sn in [BASE7,CMP7]:
  view=sh(sn)
  for addr,dest,pt,en in [('C1',BN,'Voltar ao DASH 04','Back to DASH 04'),('E1','DASH 05 V06','Voltar ao DASH 07','Back to DASH 07')]:view.put(addr,pt,'HYPERLINK("#\''+dest+'\'!A1",'+tr(pt,en)+')')
  mode='Base' if sn==BASE7 else 'Comparative';profile=inp(mode,IRO[('DASH 05 V06','H32')]);view.put('I8','Operator profile / Perfil operadora');f=profile+'&IF(OR('+','.join(q(sn,'B'+str(ROWS[k]))+'<>"Profile"' for k in ['basefee','incentive','systems'])+'),'+tr(' — personalizado',' — customised')+',"")';view.put('I9','Custom Terms',f);view.put('I12','Guardar versões: fechar o ficheiro e executar Gerir_Cenarios_Excel365.ps1; alternativa Python disponível no repositório. / Save versions: close the workbook and run the supplied Excel Desktop script; Python alternative is in the repository.');changes[s.pp[sn]]=view.finish()
 view=sh(STATUS7);view.put('A19','Versão a repor / Version to restore');view.put('B19','BASE-0001');view.put('D19','Saved version IDs');view.put('D20','BASE-0001');view.put('D21','CMP-0001');view.put('A23','Restore: select B19, save and close this workbook, then run Gerir_Cenarios_Excel365.ps1 -Action Restore. Existing versions are retained.');dvs=view.r.find(T('dataValidations'))
 if dvs is None:dvs=E.SubElement(view.r,T('dataValidations'),count='0')
 for node in list(dvs):
  if node.get('sqref')=='B19':dvs.remove(node)
 dv=E.SubElement(dvs,T('dataValidation'),type='list',sqref='B19',showErrorMessage='1',errorStyle='stop');E.SubElement(dv,T('formula1')).text='$D$20:$D$1000';dvs.set('count',str(len(dvs)));dimension=view.r.find(T('dimension'))
 if dimension is not None:dimension.set('ref','A1:D23')
 changes[s.pp[STATUS7]]=view.finish()
 view=sh(REPORT7)
 for j,label in enumerate(['Report ID','Calculation UTC','Format','Language','Recipient','Base ID','Comparative IDs','TAN','Discount','Validation','PDF SHA256','JSON SHA256']):view.put(col(j)+'24',label)
 dimension=view.r.find(T('dimension'))
 if dimension is not None:dimension.set('ref','A1:L24')
 changes[s.pp[REPORT7]]=view.finish()
 tmp=WORK7/'Finished_UI.xlsx';rewrite(file,tmp,changes);s.z.close();shutil.move(tmp,file)
if __name__=='__main__':finish_ui()
