"""Bank scenarios: mutate real inputs, recalculate, assert 180-month invariants.
Only unrelated sensitivity branches are frozen in private speed-test labs.
Public file and issued reports retain all engines and use full recalculation.
"""
from bank_emit import *
from concurrent.futures import ThreadPoolExecutor,as_completed

def lab():
 target=BW/'Bank_Test_Lab.xlsx'
 with ZipFile(BM) as z,ZipFile(target,'w',ZIP_DEFLATED,compresslevel=3) as out:
  _,pp=paths(z);lookup={p:n for n,p in pp.items()}
  for item in z.infolist():
   data=z.read(item.filename);sn=lookup.get(item.filename,'');freeze=sn=='V06 Referencia' or bool(re.fullmatch(r'V06 [SMO]\d+',sn)) or bool(re.fullmatch(r'V06 F\d+',sn) and sn!='V06 F01')
   if freeze or sn=='V06 Cenarios':
    r=E.fromstring(data)
    for row in r.find(T('sheetData')):
     if sn=='V06 Cenarios' and int(row.get('r'))==5:continue
     for c in row:
      f=c.find(T('f'))
      if f is not None:c.remove(f)
    data=xml(r)
   out.writestr(copy.copy(item),data)
 return target

def case(x,source,base):
 name,changes=x;folder=BW/'tests'/name;folder.mkdir(parents=True,exist_ok=True);target=folder/'Input.xlsx'
 try:
  prior=sha(target) if target.exists() else None;mutate(source,target,[(BN,'H33','Custom')]+changes);output=folder/'recalc/out/Input.xlsx';f=output if prior==sha(target) and output.exists() else recalculate(target,folder/'recalc');data=evaluate(f);c,errors,_=read_bank(f);assert not errors
  for k in range(180):
   rr=k+5;r=data['monthly'][k]
   for sn in [v[0] for v in DASHES]:near(c[sn]['V30'],data['npv'])
   if name=='Reference_restore':
    for j in range(140):near(c['V06 Motor'][col(j)+str(rr)],base['cells']['V06 Motor'][col(j)+str(rr)])
  if name=='Opening_zero':near(sum(r['bank_fees'] for r in data['monthly']),base['metrics']['Total financing fees']-300000)
  if name=='Prepay_2pct':assert sum(r['early_fee'] for r in data['monthly'])>0
  if name=='Interest_grace':assert sum(r['capitalized_interest'] for r in data['monthly'])>0
  if name=='Variable':near(data['params']['effective_tan'],.07)
  if name.startswith('View_'):near(data['npv'],base['npv'])
  result={'test':name,'status':'PASS','changes':changes,'npv':data['npv'],'irr':data['irr'],'metrics':data['metrics'],'checks':data['checks']}
 except Exception as ex:result={'test':name,'status':'FAIL','changes':changes,'error':repr(ex)}
 print(result,flush=True);return result

def main():
 source=lab();base=evaluate(BM);base['cells']=read_bank(BM)[0]
 cases=[('TAN_5_5',[(BN,'H23',.055)]),('TAN_7',[(BN,'H23',.07)]),('Term_15years',[(BN,'H25',180)]),('Grace_48',[(BN,'N23',48)]),('Opening_zero',[(BN,'H458',0)]),('Opening_2pct',[(BN,'H458',.02)]),('Prepay_2pct',[(BN,'H466',.02)]),('No_sweep',[(BN,'N25',0)]),('Facility_35M',[(BN,'H31',35000000)]),('Variable',[(BN,'H454','Variable'),(BN,'N454',.03),(BN,'H456',.04)]),('Interest_grace',[(BN,'N456',24)]),('Commitment',[(BN,'H462',.01)]),('Management',[(BN,'H464',10000),(BN,'N464',3)]),('ACT365',[(BN,'H472','ACT/365')]),('Bullet',[(BN,'H29','Bullet'),(BN,'N25',0)]),('Custom_prepay',[(BN,'H474','Custom'),(BN,'N474',100000)]),('Include_fees_DSCR',[(BN,'N476',1)]),('Proposal_A',[(BN,'H33','Proposal Bank A')]),('Proposal_B',[(BN,'H33','Proposal Bank B')]),('Official_restore',[(BN,'H33','Official KK-01 6%')]),('Reference_restore',[(BN,'H23',.075),(BN,'N456',24),(BN,'H466',.02),('EXEC V06','H23','Reference')]),('View_quarters',[(SCH,'E320','Quarterly')])]
 results=[]
 with ThreadPoolExecutor(max_workers=3) as ex:
  for f in as_completed([ex.submit(case,x,source,base) for x in cases]):results.append(f.result())
 proof={'status':'PASS' if all(x['status']=='PASS' for x in results) else 'FAIL','tests':results,'official_base':{'tan':.06,'grace_months':42},'native_excel365':'NOT TESTED'};(BOUT/'Testes_Bancarios.json').write_text(json.dumps(proof,indent=2,ensure_ascii=False));assert proof['status']=='PASS'
if __name__=='__main__':main()
