"""Format and protect the already calculated bank emission, preserving all engine
parts byte-for-byte. Historical source is retained by its Git commit and SHA256.
"""
from bank_emit import *
from bank_domains import apply,LABELS
from zip_copy import rewrite

def main():
 dest=BOUT/'Emissao_Demonstracao_6_vs_5_5';snap=dest/'Snapshot.xlsx';datafile=dest/'Snapshot.json';payload=json.loads(datafile.read_text());data=payload['selected'];base=payload['comparison'];commit='b900fdd5f752051096bb3ea65a097a87d7e819c9'
 import subprocess
 blob=subprocess.check_output(['git','show',commit+':entregas/EFCC_Final/Fecho_R6_Banco/KK01_Executive_Financial_Control_Center_R6_BANK_CHECKPOINT.xlsx']);assert hashlib.sha256(blob).hexdigest()==data['source_sha256'];del blob
 def finance_hash(data):
  root=E.fromstring(data);h=hashlib.sha256()
  for c in root.iter(T('c')):h.update((c.get('r')+'|'+str(c.get('t'))+'|'+str(c.findtext(T('f')))+'|'+str(c.findtext(T('v')))).encode())
  return h.hexdigest()
 original=BW/'emissions/KK01-BANK-20261023T175546Z/selected/Snapshot.xlsx'
 with ZipFile(original) as z:
  _,pp=paths(z);engine={sn:finance_hash(z.read(p)) for sn,p in pp.items() if sn=='V06 Motor' or re.fullmatch(r'V06 [SMOF]\d+',sn)}
 apply(snap)
 changes={}
 with ZipFile(snap) as z:
  w,pp=paths(z)
  for node in w.find(T('sheets')):
   sn=node.get('name')
   if node.get('state','visible')!='visible':continue
   root=E.fromstring(z.read(pp[sn]));s=Sheet(root,{})
   if root.find(T('sheetProtection')) is None:E.SubElement(root,T('sheetProtection'),sheet='1',objects='1',scenarios='1',selectLockedCells='0',selectUnlockedCells='0')
   if sn in [d[0] for d in DASHES]:
    cell=s.cells.get('E1');s.put('E1','R6 ISSUED ARCHIVE | '+data['id'],style=int(cell.get('s','0')) if cell is not None else 0)
   if sn==BN:
    for area,pt,en in [('E23:G23','Custom guardado · TAN','Stored Custom · TAN'),('K23:M23','Custom guardado · carência','Stored Custom · grace'),('E25:G25','Custom guardado · prazo','Stored Custom · term')]:
     cell=s.cells[area.split(':')[0]];s.merge(area,pt,tr(pt,en),int(cell.get('s','0')))
    for row in range(594,628,2):
     cell=s.cells.get('E'+str(row))
     if cell is not None:
      v=cell.findtext(T('v'))
      if cell.get('t')=='inlineStr':v=''.join(cell.find(T('is')).itertext())
      if v in LABELS:s.put(cell.get('r'),LABELS[v][0],tr(*LABELS[v]),int(cell.get('s','0')))
   changes[pp[sn]]=s.finish()
 tmp=snap.with_suffix('.protected.xlsx');rewrite(snap,tmp,changes);tmp.replace(snap)
 from bank_view import apply as select_home
 select_home(snap,percent=False)
 from closure_presentation import chart_cache
 from bank_chart_parts import apply as separate_charts
 chart_cache(snap);separate_charts(snap)
 with ZipFile(snap) as z:
  _,pp=paths(z)
  for sn,h in engine.items():assert finance_hash(z.read(pp[sn]))==h,(sn,'engine changed during presentation pass')
 data['source_git_commit']=commit;data['snapshot_sha256']=sha(snap);data['archive_status']='Read-only issued snapshot; use R6.1 master for new simulations';data['engine_financial_records_preserved_during_archive_formatting']=len(engine);datafile.write_text(json.dumps(payload,indent=2,ensure_ascii=False))
 for lang in ['PT','EN']:
  bank_pdf(data,base,dest/f'Bank_Report_10_{lang}.pdf',data['recipient'],lang,stresses=payload['stress']);bank_pdf(data,base,dest/f'Executive_2_{lang}.pdf',data['recipient'],lang,brief=True)
 (dest/'LEIA-ME.md').write_text('Emissão de demonstração, não enviada a qualquer instituição. Cenário hipotético TAN 5,5% vs. Base Oficial TAN 6%, ambos com carência 42 meses.\n\nOrigem financeira R6: Git '+commit+'; SHA256 em Snapshot.json. Snapshot.xlsx é um arquivo protegido contra edição acidental, com as mesmas fórmulas e resultados financeiros calculados na emissão. Usar o master R6.1 para novas simulações.\n\nPDFs A4: relatório bancário 10 páginas, consolidado 10 páginas, resumo 2 páginas, PT/EN. LibreOffice e ReportLab; Excel 365 nativo não testado.\n')
 files={str(f.relative_to(dest)):sha(f) for f in dest.rglob('*') if f.is_file() and f.name not in ['Manifest.json','BLOCKED.json']};(dest/'Manifest.json').write_text(json.dumps({'status':'CALCULATED CHECKPOINT - native Excel acceptance pending','id':data['id'],'source_git_commit':commit,'source_sha256':data['source_sha256'],'files':files,'engine_financial_records_unchanged_during_archive_formatting':len(engine)},indent=2))
 blocked=dest/'BLOCKED.json'
 if blocked.exists():blocked.unlink()
 print('Bank archive finalised; 10/2-page PDFs and engine hashes verified',flush=True)
if __name__=='__main__':main()
