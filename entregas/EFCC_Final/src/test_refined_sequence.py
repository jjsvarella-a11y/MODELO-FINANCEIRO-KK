"""Actual save/reopen sequence on the main engine and refined presentation."""
import os,json,sys,shutil
from pathlib import Path
from zipfile import ZipFile
from xlsxwriter.utility import xl_col_to_name as col
from premium_refine import OUT,WORK,MASTER,H,Q
SEQ=WORK/'sequence';SEQ.mkdir(exist_ok=True)
os.environ.update(KK_TEST_OUTPUT_ROOT=str(OUT),KK_TEST_WORK=str(SEQ),KK_TEST_SOURCE=str(MASTER))
import test_premium_controls as p
import test_all_dashboard_filters as lab
from verify_all_filter_outputs import inspect as check_outputs
from scenario_reference import reference
old=p.old

def main():
 old.make_lab();p.trim_private_lab();lab.freeze_unrelated_matrix_lab();results=[];baseline=None
 for index,mode in enumerate(['Base Case','Conservative','Favorable','Custom','Base Case']):
  name=f'Sequence_{index}';changes=[(H,'H23',mode)]
  if mode=='Custom':changes.extend([(H,'H25',200),(H,'H27',.65)])
  result=old.run_case(name,changes);assert result['status']=='PASS',result;file=SEQ/'mutations'/name/'recalc/out'/(name+'.xlsx')
  with ZipFile(file) as z:cells,errors,_=old.read(z)
  assert not errors
  rows,metrics,_=reference(mode);count=0;delta=0
  for r,row in enumerate(rows,5):
   for key,value in row.items():
    if not isinstance(value,(int,float)):continue
    actual=cells['V06 Motor'][col(old.REF['main_columns'][key])+str(r)];old.near(actual,value);delta=max(delta,abs(actual-value));count+=1
  old.near(cells[Q]['M15'],metrics['npv']);views=check_outputs({'test':name,'sheet':H,'cell':'H23','choice':mode},cells)
  proof={'mode':mode,'status':'PASS','monthly_values_compared':count,'maximum_reference_difference':delta,'npv':cells[Q]['M15'],'irr':cells[Q]['E15'],'moic':cells[Q]['I15'],'exit':cells[Q]['M9'],'dividends':cells[Q]['I9'],'display_correspondences':views['bindings_checked'],'errors':0}
  if index==0:baseline=cells
  if index==4:
   diffs=[abs(value-cells['V06 Motor'][a]) for a,value in baseline['V06 Motor'].items() if isinstance(value,(int,float))];assert max(diffs)<.00005
   for a in ['E9','I9','M9','E15','I15','M15']:old.near(cells[Q][a],baseline[Q][a])
   proof['roundtrip_maximum_monthly_difference']=max(diffs);proof['restored_base_financial_kpis']=True
  results.append(proof);print(proof,flush=True);shutil.copy2(file,SEQ/'Mutation_Lab.xlsx')
 (OUT/'Testes_Sequencia.json').write_text(json.dumps({'status':'PASS','platform':'LibreOffice save/reopen recalculation','native_excel365':'NOT TESTED','sequence':results,'scope':'Main monthly engine and refined display; matrix financial propagation is separately tested with all 30 engines live'},indent=2))
if __name__=='__main__':main()
