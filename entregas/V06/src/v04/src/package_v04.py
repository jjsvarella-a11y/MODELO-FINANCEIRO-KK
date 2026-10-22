from engine import ROOT
from pathlib import Path
import json,hashlib,zipfile,shutil,fitz,csv
OUT=ROOT/'entrega';DEST=Path('/workspace/SIM-JSCV-KYENU/entregas/V04')
parts=['Executive_Control_Center','Strategic_Master_Workflow','Estrutura_Acionista_Waterfall','Equity_Investor_Pipeline','Financiamento_Bancabilidade','Comercializacao_94_Apartamentos','Construcao_EAC','Hotel_Residential_Operator','Cenarios_Riscos','Dicionario_KPI','Auditoria','Decisoes_Pendentes','Documentacao_Tecnica']
expected=['KK01_Modelo_Integrado_V04.xlsx']+['KK01_'+x+'_V04.pdf' for x in parts]
for name in expected:
 p=OUT/name;assert p.is_file() and 0<p.stat().st_size<40000000,name
 if p.suffix=='.pdf':
  d=fitz.open(p);assert len(d)>0 and all(len(page.get_text())>50 for page in d),name;d.close()
for name in ['Testes_Seletores','Testes_Riscos_Integrados','Testes_Retornos_Completion']:
 r=json.loads((OUT/(name+'.json')).read_text());assert r and all(x['ok'] for x in r),name
assert json.loads((OUT/'Testes_Estruturais_V04.json').read_text())['failed']==0
for name in ['Testes_Fase1','Testes_Fases23']:
 r=json.loads((OUT/(name+'.json')).read_text());assert r['tests']>0 and r['failures']==r['errors']==0
records=['Testes_Fase1.json','Testes_Fase1.txt','Testes_Fases23.json','Testes_Fases23.txt','Testes_Seletores.json','Testes_Seletores_Primeira_Passagem_V04.json','Testes_Correcao_Compat.json','Testes_Riscos_Integrados.json','Testes_Riscos_Primeira_Passagem_V04.json','Testes_Retornos_Completion.json','Testes_Estruturais_V04.json','Testes_ETC_Risco.json','Verificacao_PDFs_V04.json','Verificacao_Notas_V04.json','Verificacao_Reabertura_V04.json','Registo_Alteracoes_V04.csv','Diagnostico_Waterfall_Final.json','Mapa_Folhas_V04.csv']
for name in ['Erros_V45.csv','Inventario_V45.csv','Dependencias_V45.json']:
 shutil.copy2(ROOT.parent/'entrega'/name,OUT/name);records.append(name)
active=['engine.py','waterfall.py','selected_policy.py','completion.py','returns10.py','build_model.py','v02_panels.py','v03_panels.py','scenario_seller_formulas.py','test_phase1.py','test_phase23.py','test_selectors.py','audit_risks.py','audit_returns_v04.py','test_structure_v04.py','report_base_v04.py','reports_v04.py','finalize_notes_v04.py','export_exec_v04.py','verify_reopen_v04.py','patch_draw_dependencies_v04.py','patch_compat_v04.py','presentation_labels_v04.py','recheck_compat_v04.py','package_v04.py']
active+= [p.name for p in (ROOT/'src').glob('v04*.py')]
files=[OUT/x for x in expected+records]+list(OUT.glob('KK01_*V04.md'))+[ROOT/'README.md',ROOT/'requirements.txt']+[ROOT/'src'/x for x in dict.fromkeys(active)]+[ROOT/'src'/x for x in ['source.json','build_metadata.json','v04_metadata.json','v04_risk_metadata.json','v04_risk_results.json','v04_kpi_dictionary.json']]+list((ROOT/'src/icons').glob('*.png'))+[p for p in (ROOT/'original').rglob('*') if p.is_file() and '__pycache__' not in str(p)]
assert all(p.is_file() for p in files)
manifest={str(p.relative_to(ROOT)):{'bytes':p.stat().st_size,'sha256':hashlib.sha256(p.read_bytes()).hexdigest()} for p in files};mp=OUT/'Manifesto_Conteudo_V04.json';mp.write_text(json.dumps(manifest,ensure_ascii=False,indent=2));files.append(mp)
archive=OUT/'KK01_V04_Entrega_Completa.zip'
with zipfile.ZipFile(archive,'w',zipfile.ZIP_DEFLATED,compresslevel=6) as z:
 for p in dict.fromkeys(files):z.write(p,'kk01_v04/'+str(p.relative_to(ROOT)))
with zipfile.ZipFile(archive) as z:assert z.testzip() is None
assert archive.stat().st_size<40000000,archive.stat().st_size
pub=expected+[archive.name];manifest={n:{'bytes':(OUT/n).stat().st_size,'sha256':hashlib.sha256((OUT/n).read_bytes()).hexdigest()} for n in pub};(OUT/'Manifesto_SHA256_V04.json').write_text(json.dumps(manifest,indent=2));DEST.mkdir(parents=True,exist_ok=True)
for name in pub+['Manifesto_SHA256_V04.json']:shutil.copy2(OUT/name,DEST/name)
shutil.copy2(ROOT/'README.md',DEST/'README.md')
print('Preparados',len(pub),'ficheiros publicáveis; pacote',archive.stat().st_size,'bytes;',len(files),'ficheiros auditáveis; nenhum upload executado por este script.')
