import sys,json
sys.path.insert(0,'/workspace/MODELO-FINANCEIRO-KK/entregas/EFCC_Final/src')
import test_closure_r5 as t
from concurrent.futures import ThreadPoolExecutor,as_completed
base=t.baseline(t.WORK/'final_recalc/out'/t.MASTER.name)
labs={False:t.make_lab(False,'_extra'),True:t.make_lab(True,'_extra')}
# Use the dedicated lab without changing those used by the 39 primary tests.
C='DASH 03 V06';H='DASH 05 V06';cases=[('Bank_Draw36',[('DASH 01 V06','H23',6),(C,'H27',36)],False),('Bank_Limit35',[(C,'H31',35000000)],False),('Bank_Bullet',[(C,'H29','Bullet'),(C,'N25',0)],False),('Bank_Interest',[(C,'H23',.08)],False),('Bank_Grace',[(C,'N23',48)],False),('Bank_Term',[(C,'H25',132)],False),('Res_Fees',[('DASH 04B V06','H23','KK Optimized Terms')],False),('Capital_Delay',[('DASH 06 V06','H23',3)],False),('Capital_Fraction',[('DASH 06 V06','N23',.8)],False),('Bank_Reference',[(C,'H29','Bullet'),(C,'H27',42),(C,'H31',35000000),('EXEC V06','H23','Reference')],False),('Matrix_Chapter',[('DASH 01 V06','N194',.1)],True),('Matrix_Pricing',[('DASH 04A V06','H25','Favorable')],True)]
results=[]
with ThreadPoolExecutor(max_workers=3) as ex:
 for f in as_completed([ex.submit(t.run_case,x,base,labs) for x in cases]):results.append(f.result())
proof={'status':'PASS' if all(r['status']=='PASS' for r in results) else 'FAIL','tests':results,'native_excel365':'NOT TESTED'};(t.OUT/'Testes_Adicionais_R5.json').write_text(json.dumps(proof,indent=2,ensure_ascii=False));assert proof['status']=='PASS'
# Actual sequential save/reopen cycle retaining the stored Test inputs.
source=labs[False];sequence=[];stress=[('DASH 05 V06','H23','Custom'),('DASH 05 V06','H25',205),('DASH 07 V06','H29','8%')];stress_vector=None
for name,changes in [('Seq_Test',stress),('Seq_Reference',[('EXEC V06','H23','Reference')]),('Seq_TestResume',[('EXEC V06','H23','Test')]),('Seq_ReferenceEnd',[('EXEC V06','H23','Reference')])]:
 folder=t.WORK/name;folder.mkdir(exist_ok=True);file=folder/(name+'.xlsx');t.mutate(source,file,changes);source=t.recalculate(file,folder/'recalc');c,e,_=t.load(source);assert not e,e[:3];r=t.inspect(c,base,changes,name)
 vector=[c['V06 Motor'][t.col(k)+str(rr)] for rr in range(5,185) for k in range(140)]
 if name=='Seq_Test':stress_vector=vector
 if name=='Seq_TestResume':
  for x,y in zip(vector,stress_vector):t.near(x,y)
 assert c['DASH 05 V06']['H25']==205 and c['DASH 07 V06']['H29']=='8%'
 sequence.append(r);print(name,'PASS',flush=True)
(t.OUT/'Sequencia_R5.json').write_text(json.dumps({'status':'PASS','sequence':sequence,'native_excel365':'NOT TESTED'},indent=2))
