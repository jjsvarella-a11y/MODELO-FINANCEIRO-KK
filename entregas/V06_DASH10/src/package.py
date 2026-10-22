from pathlib import Path
from zipfile import ZipFile,ZIP_DEFLATED
import hashlib,json
ROOT=Path(__file__).resolve().parents[1]
def package():
 validation=json.loads((ROOT/'Validacao_Entrega.json').read_text());assert validation['failed']==0
 archive=ROOT/'KK01_V06_DASH10_Entrega.zip';checksum=archive.with_suffix('.zip.sha256');manifest=ROOT/'SHA256SUMS.json'
 files=[p for p in sorted(ROOT.rglob('*')) if p.is_file() and p not in [archive,checksum,manifest] and '__pycache__' not in p.parts and not p.name.endswith('.pyc')]
 records=[{'file':str(p.relative_to(ROOT)),'bytes':p.stat().st_size,'sha256':hashlib.sha256(p.read_bytes()).hexdigest()} for p in files];manifest.write_text(json.dumps(records,ensure_ascii=False,indent=2));files.append(manifest)
 with ZipFile(archive,'w',ZIP_DEFLATED,compresslevel=9) as z:
  for p in files:z.write(p,str(p.relative_to(ROOT)))
 with ZipFile(archive) as z:
  assert z.testzip() is None
  for r in records:assert hashlib.sha256(z.read(r['file'])).hexdigest()==r['sha256']
 checksum.write_text(hashlib.sha256(archive.read_bytes()).hexdigest()+'  '+archive.name+'\n');print(json.dumps({'files':len(files),'zip_bytes':archive.stat().st_size,'checks_passed':validation['passed'],'checks_failed':validation['failed']}))
if __name__=='__main__':package()
