"""Version stamp and verified detail/return destinations; no financial changes."""
from polish_v08 import *
def release_metadata(source,target):
 src=Source(source);styles=add_styles(src.styles);changes={}
 for sn,*_ in DASHES:
  s=Sheet(E.fromstring(src.z.read(src.pp[sn])),styles);s.put('E1','V08 · KK-01 · 2026-10-25 · Base editável / Independent Comparative',style='note')
  idx=[d[0] for d in DASHES].index(sn)
  if idx==3:s.text('E4:Y4','04 — Financiamento e Desempenho Financeiro','04 — Financing & Financial Performance','sub')
  if idx==9:s.text('E4:Y4','10 — Central de Relatórios Executivos','10 — Executive Reporting Center','sub')
  for n,row,pt,en in [(3,34,'Financiamento e Desempenho Financeiro','Financing & Financial Performance'),(9,58,'Central de Relatórios Executivos','Executive Reporting Center')]:
   style='side_active' if n==idx else 'side';dest=DASHES[n][0];s.merge('C'+str(row)+':D'+str(row+3),pt,f'HYPERLINK("#\'{dest}\'!A1",'+tr(pt,en)+')',style)
  changes[src.pp[sn]]=s.finish()
 sn='DASH 01 V06';s=Sheet(E.fromstring(changes[src.pp[sn]]),styles)
 for addr,area,back in [('E190','E190:S191','T190:Y191'),('E230','E230:S231','T230:Y231')]:
  f,v=src.get(sn,addr);s.merge(area,v,f,'section');s.link(back,sn,'A1','Voltar ao DASH 02','Back to DASH 02')
 changes[src.pp[sn]]=s.finish()
 sn='DASH 03 V06';s=Sheet(E.fromstring(changes[src.pp[sn]]),styles);s.link('E112:O113',sn,'E450','Condições bancárias e parâmetros secundários →','Bank terms and secondary assumptions →');s.link('P112:Y113',sn,'E546','Composição do custo do financiamento →','Financing cost breakdown →');s.text('E546:S547','Composição do custo de financiamento','Financing cost breakdown','section');s.link('T546:Y547',sn,'A1','Voltar ao DASH 04','Back to DASH 04');s.link('T450:Y451',sn,'A1','Voltar ao DASH 04','Back to DASH 04');changes[src.pp[sn]]=s.finish()
 sn='DASH 05 V06';s=Sheet(E.fromstring(changes[src.pp[sn]]),styles);s.link('E142:O144',sn,'E328','Composição da receita — detalhe →','Revenue composition — detail →');s.text('E328:S329','Composição da receita','Revenue composition','section');s.link('T328:Y329',sn,'A1','Voltar ao DASH 07','Back to DASH 07');changes[src.pp[sn]]=s.finish()
 sn='DASH 06 V06';s=Sheet(E.fromstring(changes[src.pp[sn]]),styles);s.link('E128:O130',sn,'E224','Registo dos 15 blocos e direitos →','15-block register and rights →');f,v=src.get(sn,'E224');s.merge('E224:S225',v,f,'section');s.link('T224:Y225',sn,'A1','Voltar ao DASH 08','Back to DASH 08');changes[src.pp[sn]]=s.finish()
 changes['xl/styles.xml']=xml(src.styles)
 index=[x.get('name') for x in src.w.find(T('sheets'))].index('EXEC V06')
 for v in src.w.findall('.//'+T('workbookView')):v.set('firstSheet',str(index))
 changes['xl/workbook.xml']=xml(src.w);rewrite(source,target,changes);src.z.close()
if __name__=='__main__':
 p=argparse.ArgumentParser();p.add_argument('source',type=Path);p.add_argument('target',type=Path);a=p.parse_args();release_metadata(a.source,a.target)
