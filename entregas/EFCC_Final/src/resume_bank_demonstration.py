"""Resume the recorded R6 demonstration after full recalc, verifying input hashes.
Completed recalculations are reused ONLY after bitwise equality of the rebuilt
input ZIP and the actually recalculated private input; no stale caches accepted.
The full monthly reconciliation and every-sheet error scan run again.
"""
from bank_emit import *
from bank_snapshot import stamp
from concurrent.futures import ThreadPoolExecutor

def main():
 ident='KK01-BANK-20261023T175546Z';work=BW/'emissions'/ident;dest=BOUT/'Emissao_Demonstracao_6_vs_5_5';recipient='Instituição financeira — exemplo de negociação';sourcehash=sha(BM);proof=[]
 def complete(label,scenario,overrides=None,need_snapshot=False):
  folder=work/label;calculated=folder/'recalc/out/Input.xlsx';inputfile=folder/'Input.xlsx';changes=[('EXEC V06','H23','Test'),(BN,'H33',scenario),(CENTER,'H122',recipient)]+(overrides or [])
  if calculated.exists() and inputfile.exists():
   if label.startswith('stress_') and Source(inputfile).get(CENTER,'H122')[1]==Source(BM).get(CENTER,'H122')[1]:
    # Original private stresses did not apply the PDF recipient; this is
    # presentation-only. Rebuild that exact input and verify every other byte.
    changes=[v for v in changes if (v[0],v[1])!=(CENTER,'H122')]
   expected=folder/'ExpectedInput.xlsx';mutate(BM,expected,changes)
   if sha(expected)!=sha(inputfile):raise ValueError('Input hash changed; full recalculation required: '+label)
   expected.unlink();data=evaluate(calculated);snapshot=folder/'Snapshot.xlsx'
   if need_snapshot and not snapshot.exists():inject_caches(inputfile,calculated,snapshot)
   data.update(source_sha256=sourcehash,snapshot_sha256=sha(snapshot if need_snapshot else calculated),calculation_platform='LibreOffice full engine recalculation',native_excel365='NOT TESTED',calculated_at=datetime.fromtimestamp(calculated.stat().st_mtime,timezone.utc).isoformat());proof.append({'case':label,'input_sha256':sha(inputfile),'calculated_sha256':sha(calculated),'reused_completed_full_recalculation':True,'input_bitwise_match':True})
   return snapshot if need_snapshot else calculated,data
  file,data=compute(BM,scenario,folder,recipient,overrides,need_snapshot=need_snapshot);proof.append({'case':label,'input_sha256':sha(folder/'Input.xlsx'),'calculated_sha256':sha(folder/'recalc/out/Input.xlsx'),'reused_completed_full_recalculation':False,'input_bitwise_match':True});return file,data
 snap,data=complete('selected','Proposal Bank A',need_snapshot=True);_,base=complete('comparison','Official KK-01 6%')
 stresses=[]
 for label,change in [('Construction +6m',[('DASH 01 V06','H23',6)]),('Operational Conservative',[('DASH 05 V06','H23','Conservative')]),('Sales collections +6m',[('DASH 04A V06','N23',6)])]:
  idx=len(stresses)+1;_,r=complete('stress_'+str(idx),'Proposal Bank A',change);r['label']=label;stresses.append(r)
 for obj in [data,base]:obj.update(id=ident,recipient=recipient)
 data['comparison_guard']=stamp(snap,data,base);data['snapshot_sha256']=sha(snap);shutil.copy2(snap,dest/'Snapshot.xlsx')
 (dest/'Snapshot.json').write_text(json.dumps({'selected':data,'comparison':base,'stress':stresses,'selection_only':False},indent=2,ensure_ascii=False));(dest/'Recalculation_Provenance.json').write_text(json.dumps(proof,indent=2))
 with (dest/'Monthly_180.csv').open('w',newline='') as f:w=csv.DictWriter(f,fieldnames=list(data['monthly'][0]));w.writeheader();w.writerows(data['monthly'])
 for lang in ['PT','EN']:
  bank_pdf(data,base,dest/f'Bank_Report_10_{lang}.pdf',recipient,lang,stresses=stresses);bank_pdf(data,base,dest/f'Executive_2_{lang}.pdf',recipient,lang,brief=True)
  import closure_presentation as cp
  from export_all import export
  cp.MASTER=snap;cp.WORK=work/('consolidated_'+lang);cp.WORK.mkdir(parents=True,exist_ok=True)
  pres=cp.lab(lang);cp.chart_cache(pres);pdf=export(pres,['DASH 10 V06'],'Consolidated',work/('pdf_'+lang));shutil.copy2(pdf,dest/f'Consolidated_10_{lang}.pdf')
 if sha(BM)!=sourcehash:raise ValueError('Source changed during issue')
 files={str(p.relative_to(dest)):sha(p) for p in dest.rglob('*') if p.is_file() and p.name!='Manifest.json'};(dest/'Manifest.json').write_text(json.dumps({'status':'CALCULATED CHECKPOINT - native Excel acceptance pending','id':ident,'files':files,'source_sha256':sourcehash,'source_unchanged':True},indent=2));print('ISSUED',dest,flush=True)
if __name__=='__main__':
 try:main()
 except Exception as ex:
  dest=BOUT/'Emissao_Demonstracao_6_vs_5_5'
  for f in dest.glob('*.pdf'):f.unlink()
  (dest/'BLOCKED.json').write_text(json.dumps({'status':'BLOCKED','reason':repr(ex)},indent=2))
  raise
