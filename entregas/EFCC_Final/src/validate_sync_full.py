"""Corrected Base reference from the existing core and original scenario definitions."""
import ast,json,gzip,sys
from pathlib import Path
from collections import defaultdict
from zipfile import ZipFile
import scenario_reference as sr
import validate_final as vf
from reference_io import load_reference
from premium_ui import ROOT,WORK

def literal_assignment(file,name):
 for node in ast.walk(ast.parse(file.read_text())):
  if isinstance(node,ast.Assign) and any(isinstance(t,ast.Name) and t.id==name for t in node.targets):return ast.literal_eval(node.value)
 raise AssertionError(name)
def main():
 ref=load_reference(ROOT);params=ref['EFCC']['effective'].copy();params['hotel_scenario']='Base'
 model=sr.Model(sr.SinkWorkbook(),defaultdict(lambda:None),{k:sr.asx(v) for k,v in params.items()},[{k:sr.asx(v) for k,v in u.items()} for u in sr.legacy.UNITS],sr.legacy.CHAPTERS)
 model.tickets=[(code,cat,*map(sr.asx,[amount,share,month,1,1,0,0])) for code,cat,amount,share,month in sr.TICKETS]
 definitions=literal_assignment(Path(__file__).with_name('build_final.py'),'scenario_defs')
 definitions += [(f'ADR {adr} Occ {occ:.0%}',{'hotel_mode':'Custom','adr_input':adr,'occ_input':occ}) for occ in [.50,.55,.60,.65,.70,.75] for adr in [150,175,200,225,250]]
 definitions += literal_assignment(Path(__file__).with_name('enhance.py'),'extra')
 for label,overrides in definitions:
  rows=model.calculate('NumericTest',overrides)['rows'];ref['scenarios'][label]=[{k:v.v for k,v in row.items()} for row in rows];print(label,flush=True)
 ref['main_rows']=ref['scenarios']['Base'];ref['EFCC']['effective']=params
 artifact=ROOT/'Referencia_Python_V06_SYNC.json.gz'
 artifact_reference={k:ref[k] for k in ['main_rows','scenarios','main_columns']}
 artifact_reference.update(EFCC={'effective':params},scope='Independent corrected monthly numerical reference only; financial metrics are verified separately in Sincronizacao_Cenarios.json; no historical dashboard caches included')
 with gzip.open(artifact,'wt',encoding='utf8') as f:json.dump(artifact_reference,f,ensure_ascii=False)
 vf.load_reference=lambda _:ref
 result=vf.validate(WORK/'premium_final_recalc/out/KK01_V06_Executive_Master.xlsx')
 result.update(reference=artifact.name,change='Corrected Base operational profile; historical reference retained unchanged')
 (ROOT/'Validacao_Financeira.json').write_text(json.dumps(result,indent=2,ensure_ascii=False))
if __name__=='__main__':main()
