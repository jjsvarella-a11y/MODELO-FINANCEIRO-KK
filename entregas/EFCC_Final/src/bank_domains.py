"""Validate dropdown domains, apply clear units and retain controlled base selection."""
from bank_r6 import *
from zip_copy import rewrite
from datetime import date
LABELS={'Maximum annual debt service':('Serviço anual máximo da dívida','Maximum annual debt service'),'Minimum DSCR':('DSCR mínimo','Minimum DSCR'),'Peak outstanding debt':('Pico de dívida','Peak outstanding debt'),'Final debt repayment month':('Mês de liquidação da dívida','Final debt repayment month'),'Total interest paid':('Total de juros pagos','Total interest paid'),'Total financing fees':('Total de comissões e encargos','Total financing fees'),'Total financing cost':('Custo total do financiamento','Total financing cost'),'Effective annual financing cost':('Custo efetivo anual do financiamento','Effective annual financing cost'),'Funding gap':('Défice de financiamento','Funding gap'),'Minimum liquidity headroom':('Folga mínima de liquidez','Minimum liquidity headroom'),'Remaining non-revolving capacity':('Capacidade remanescente da linha','Remaining non-revolving capacity'),'New investor NPV':('VAL dos novos investidores','New investor NPV'),'New investor IRR':('TIR dos novos investidores','New investor IRR'),'MOIC':('MOIC','MOIC'),'Exit equity value':('Valor dos capitais próprios na saída','Exit equity value'),'Cumulative dividends':('Dividendos acumulados','Cumulative dividends'),'Interest accrued / due':('Juros devidos','Interest accrued / due'),'Year':('Ano','Year'),'Receipts':('Recebimentos','Receipts'),'Equity':('Capital próprio','Equity'),'Draw':('Utilizações','Draw'),'Repayment':('Amortizações','Repayment'),'Debt':('Dívida','Debt'),'Interest':('Juros','Interest'),'Fees':('Comissões e encargos','Fees'),'Debt service':('Serviço da dívida','Debt service'),'Principal':('Capital programado','Principal')}
METRIC_KEYS=list(LABELS)[:17]

def apply(file=None):
 target=Path(file or BM);changes={};corrected=0
 with ZipFile(target) as z:
  wb,pp=paths(z);ss=[''.join(n.itertext()) for n in E.fromstring(z.read('xl/sharedStrings.xml'))]
  for sn,p in pp.items():
   if sn=='V06 Motor' or re.fullmatch(r'V06 [SMO]\d+',sn):continue
   root=E.fromstring(z.read(p));dirty=False;s=Sheet(root,{})
   for f in root.iter(T('formula1')):
    if f.text and 'Official KK-01 7%' in f.text:f.text=f.text.replace('Official KK-01 7%','Official KK-01 6%');dirty=True;corrected+=1
   if sn in [REP,BRF,SCH]:
    for cell in list(s.cells.values()):
     typ=cell.get('t');v=cell.findtext(T('v'));txt=ss[int(v)] if typ=='s' else ''.join(cell.find(T('is')).itertext()) if typ=='inlineStr' else v
     if txt in LABELS and cell.find(T('f')) is None:
      pt,en=LABELS[txt];s.put(cell.get('r'),pt,tr(pt,en),int(cell.get('s','0')));dirty=True
   if sn==BN:
    dvs=root.find(T('dataValidations'))
    rules={'N472':('date',36526,73415),'N456':('whole',0,180),'N464':('whole',1,360),'N470':('whole',1,180),'H476':('decimal',.01,10),'N476':('whole',0,1),'N23':('whole',0,180)}
    for addr,(kind,lo,hi) in rules.items():
     for n in list(dvs):
      if n.get('sqref')==addr:dvs.remove(n)
     n=E.SubElement(dvs,T('dataValidation'),type=kind,operator='between',sqref=addr,allowBlank='0',showErrorMessage='1',errorStyle='stop',errorTitle='Entrada / Input',error='Valor fora do domínio permitido / Value outside permitted range.');E.SubElement(n,T('formula1')).text=str(lo);E.SubElement(n,T('formula2')).text=str(hi)
    dvs.set('count',str(len(dvs)));dirty=True
   if sn==CENTER:
    dvs=root.find(T('dataValidations'))
    if dvs is not None:
     for old in list(dvs):
      if old.get('sqref')=='N11':dvs.remove(old);dirty=True
     dvs.set('count',str(len(dvs)))
   if sn==SCH:
    cell=s.cells['B230'];f=cell.findtext(T('f')).replace('SUM(C5:C184)>0,','');f=f.replace('AND(','AND('+d('interest_grace')+'<='+d('grace')+'+'+d('term')+',',1);s.put('B230',cell.findtext(T('v')) or 'BLOCKED',f,int(cell.get('s')));dirty=True
   if dirty:changes[p]=s.finish()
  # The literal dropdown and the actual scenario table must agree exactly.
  root=E.fromstring(changes.get(pp[BN],z.read(pp[BN])));s=Sheet(E.fromstring(changes.get(pp[PRO],z.read(pp[PRO]))),{});names=[]
  for rr in range(8,20):
   c=s.cells['A'+str(rr)];names.append(ss[int(c.findtext(T('v')))] if c.get('t')=='s' else ''.join(c.find(T('is')).itertext()))
  for n in root.findall(T('dataValidations')+'/'+T('dataValidation')):
   if n.get('sqref') in ['H33','N33']:assert n.findtext(T('formula1')).strip('"').split(',')==names
 tmp=target.with_suffix('.domains.xlsx');rewrite(target,tmp,changes);tmp.replace(target);print('Dropdown domains checked; obsolete bank base entries removed',corrected,flush=True)
if __name__=='__main__':apply()
