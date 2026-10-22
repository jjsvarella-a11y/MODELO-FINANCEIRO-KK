from pathlib import Path
from zipfile import ZipFile,ZIP_DEFLATED
import hashlib,json,datetime
import fitz
ROOT=Path(__file__).resolve().parents[1]
def package():
 tests=[]
 for name in ['Testes_Seletores.json','Testes_Seletores_Adicionais.json']:tests+=json.loads((ROOT/name).read_text())['tests']
 assert len(tests)==17 and all(t['status']=='PASS' for t in tests)
 for file in ['Compat_Paridade_Python.json','Compat_Recalculo.json','Validacao_Financeira.json','Auditoria_OOXML.json','Prova_Final_Identidade.json','Organizacao_Folhas.json','Separacao_V04.json','Diagnostico_G5.json','Integridade_Premium.json','Validacao_Navegacao.json','Testes_Controlo_Premium.json','Testes_Todos_Filtros.json','Testes_Carencia_Prazo.json','Restauro_Mandato.json','Testes_Restauro.json','Validacao_Vinculos_Premium.json','PDF_Premium.json','Sincronizacao_Cenarios.json','Testes_Correspondencias_Dashboards.json','Preservacao_Motor_SYNC.json']:assert json.loads((ROOT/file).read_text())['status']=='PASS',file
 pdfs=[]
 for language in ['PT','EN']:
  for p in sorted((ROOT/'PDF'/language).glob('*.pdf')):
   d=fitz.open(p);expected=10 if 'Integral' in p.name else 2 if 'Brief' in p.name else 1;assert len(d)==expected,(p,len(d));paper='A4 portrait' if expected>1 else 'A3 landscape';pdfs.append({'file':str(p.relative_to(ROOT)),'pages':len(d),'paper':paper,'native_excel':'NOT TESTED'})
 assert len(pdfs)==22,len(pdfs)
 (ROOT/'PDF_Entrega_Final.json').write_text(json.dumps(pdfs,indent=2))
 files=[p for p in ROOT.rglob('*') if p.is_file() and '__pycache__' not in p.parts and p.suffix not in ['.pyc','.zip'] and p.name not in ['Manifesto.json','SHA256SUMS.txt']]
 manifest={'release':'V06.3-SYNC-20261022','native_excel365':'NOT TESTED','reported_old_version_errors':'User reports previous issue appears resolved. Premium revision has independent tests; native Excel acceptance remains NOT TESTED.','logo_entry':'V06 Identidade: user inserts original JPG/PNG inside the workbook; optional header sync provided, native test pending.','files':[{'path':str(p.relative_to(ROOT)),'bytes':p.stat().st_size,'sha256':hashlib.sha256(p.read_bytes()).hexdigest()} for p in sorted(files)]}
 manifest['test_evidence']={'current':['Sincronizacao_Cenarios.json','Testes_Todos_Filtros.json','Testes_Correspondencias_Dashboards.json','Testes_Carencia_Prazo.json','Validacao_Financeira.json','Preservacao_Motor_SYNC.json','Auditoria_OOXML.json','Validacao_Navegacao.json','Validacao_Vinculos_Premium.json','PDF_Premium.json'],'historical':['Testes_Seletores.json','Testes_Seletores_Adicionais.json','Testes_Controlo_Premium.json','Testes_Restauro.json','Restauro_Mandato.json','Compat_Paridade_Python.json','Compat_Recalculo.json'],'current_summary':'Relatorio_Sincronizacao_Cenarios.pdf','historical_reports':['Relatorio_Auditoria_EFCC.pdf','Relatorio_Alteracoes_Graficas.pdf']}
 (ROOT/'Manifesto.json').write_text(json.dumps(manifest,indent=2,ensure_ascii=False));files.append(ROOT/'Manifesto.json');(ROOT/'SHA256SUMS.txt').write_text(''.join(hashlib.sha256(p.read_bytes()).hexdigest()+'  '+str(p.relative_to(ROOT))+'\n' for p in sorted(files)));files.append(ROOT/'SHA256SUMS.txt');dest=ROOT/'KK01_EFCC_Entrega_Candidata.zip'
 with ZipFile(dest,'w',ZIP_DEFLATED,compresslevel=9) as z:
  for p in sorted(files):z.write(p,str(p.relative_to(ROOT)))
 with ZipFile(dest) as z:assert z.testzip() is None
 print('Package',dest,'files',len(files),'bytes',dest.stat().st_size)
if __name__=='__main__':package()
