"""Package only a passing, physically verified V06 release."""
from pathlib import Path
from zipfile import ZipFile,ZIP_DEFLATED
import hashlib,json,shutil
from finish import finish
OUT=Path(__file__).resolve().parents[1]
def package():
 validation=json.loads((OUT/'Validation_Final.json').read_text());selectors=json.loads((OUT/'Testes_Seletores_V06.json').read_text())
 assert validation['failed']==0 and all(x['passed'] for x in selectors)
 master=OUT/'KK01_V06_Executive_Master.xlsx';master_hash=hashlib.sha256(master.read_bytes()).hexdigest()
 for checkpoint in (OUT/'checkpoints').glob('C[1-5].json'):
  r=json.loads(checkpoint.read_text());r['workbook_sha256']=master_hash;r['final_checks_passed']=validation['passed'];r['final_checks_failed']=0;checkpoint.write_text(json.dumps(r,indent=2))
 archive=OUT/'KK01_V06_Entrega_Completa.zip';checksum=OUT/'KK01_V06_Entrega_Completa.zip.sha256'
 if checksum.exists():checksum.unlink()
 finish()
 manifest=json.loads((OUT/'SHA256SUMS.json').read_text())
 for entry in manifest:assert hashlib.sha256((OUT/entry['file']).read_bytes()).hexdigest()==entry['sha256']
 files=[p for p in sorted(OUT.rglob('*')) if p.is_file() and p!=archive and p!=checksum and '__pycache__' not in p.parts and not p.name.endswith('.pyc')]
 with ZipFile(archive,'w',ZIP_DEFLATED,compresslevel=9) as z:
  for p in files:z.write(p,str(p.relative_to(OUT)))
 with ZipFile(archive) as z:
  assert z.testzip() is None
  for entry in manifest:assert hashlib.sha256(z.read(entry['file'])).hexdigest()==entry['sha256']
 checksum.write_text(hashlib.sha256(archive.read_bytes()).hexdigest()+'  '+archive.name+'\n')
 print(json.dumps({'package':str(archive),'bytes':archive.stat().st_size,'files':len(files),'master_sha256':master_hash,'all_manifest_hashes_verified':True}))
if __name__=='__main__':package()
