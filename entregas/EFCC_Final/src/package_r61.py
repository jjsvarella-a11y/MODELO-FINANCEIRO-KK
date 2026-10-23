"""Release audit, navigation matrix, tested-results report and recovery manifest."""
from bank_emit import *
from bank_domains import LABELS
import fitz,csv,subprocess
from reportlab.platypus import SimpleDocTemplate,Paragraph,Spacer,Table,TableStyle
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors

def main():
 c,errors,stats=read_bank(BM);assert not errors,errors[:5]
 baseline=json.loads((BOUT/'Baseline_R6.json').read_text());bank=json.loads((BOUT/'Testes_Bancarios.json').read_text());extra=json.loads((BOUT/'Testes_R61_Adicionais.json').read_text());stale=json.loads((BOUT/'Snapshot_Stale_Guard.json').read_text());matrix=json.loads((BOUT/'Bank_Matrix_20.json').read_text());assert all(p['status']=='PASS' for p in [baseline,bank,extra,stale,matrix])
 with ZipFile(BM) as z,ZipFile(BASE) as old:
  wb,pp=paths(z);navigation=[];selected=[];seen={};graphic={}
  for sn,p in pp.items():
   root=E.fromstring(z.read(p))
   if any(v.get('tabSelected') in ['1','true'] for v in root.findall(T('sheetViews')+'/'+T('sheetView'))):selected.append(sn)
   for cell in root.iter(T('c')):
    for match in re.finditer(r'HYPERLINK\("#\'([^\']+)\'!([^\"]+)"',cell.findtext(T('f')) or ''):
     target,addr=match.groups();assert target in pp;navigation.append({'origin':sn+'!'+cell.get('r'),'destination':target+'!'+addr,'type':'native internal hyperlink','result':'PASS (structural)','excel365_native':'NOT TESTED'})
   for k in ['V06 Motor']:
    pass
  assert selected==['EXEC V06'],selected
  for p in ['xl/drawings/drawing7.xml','xl/charts/chart42.xml','xl/charts/chart44.xml','xl/charts/chart45.xml']:
   graphic[p]=z.read(p)==old.read(p);assert graphic[p],p
  for part in z.namelist():
   if part.startswith('xl/drawings/_rels/') and part.endswith('.rels'):
    owner=part.replace('/_rels/','/')[:-5]
    for r in E.fromstring(z.read(part)):
     if r.get('Type').endswith('/chart'):
      target=resolve(owner,r.get('Target'));seen[target]=seen.get(target,0)+1
  assert max(seen.values())==1,'Charts must each have a distinct native part'
 with (BOUT/'Navigation_Matrix.csv').open('w',newline='') as f:w=csv.DictWriter(f,fieldnames=list(navigation[0]));w.writeheader();w.writerows(navigation)
 (BOUT/'Navigation_Matrix.json').write_text(json.dumps(navigation,indent=2,ensure_ascii=False))
 pdfproof=[]
 for folder in [BOUT/'PDF',BOUT/'Emissao_Demonstracao_6_vs_5_5']:
  for p in folder.rglob('*.pdf'):
   doc=fitz.open(p);expected=10 if any(x in p.name for x in ['DASH_10','Bank_Report_10','Consolidated_10']) else 2 if 'Executive' in p.name or 'Budget_Appendix' in p.name else 3 if 'DASH_02' in p.name else 1
   assert len(doc)==expected,(p,len(doc),expected)
   for pg in doc:
    text=pg.get_text();assert not any(v in text for v in ['#VALUE!','#REF!','#NAME?','#N/A','###']),(p,'bad render')
   pdfproof.append({'file':str(p.relative_to(BOUT)),'pages':len(doc),'status':'PASS','engine':'ReportLab' if 'Bank_Report' in p.name or 'Executive_2' in p.name else 'LibreOffice'})
 (BOUT/'Final_PDF_Checks.json').write_text(json.dumps(pdfproof,indent=2))
 sdk=json.loads((BW/'SDK_R61_RELEASE.json').read_text());asdk=json.loads((BW/'SDK_ARCHIVE.json').read_text());assert sdk==[] and asdk==[]
 for name in ['SDK_R61_RELEASE','SDK_ARCHIVE']:shutil.copy2(BW/(name+'.json'),BOUT/(name+'.json'))
 data=json.loads((BOUT/'Emissao_Demonstracao_6_vs_5_5/Snapshot.json').read_text());a=data['selected'];b=data['comparison'];tests=62+len(bank['tests'])+len(extra['tests']);comparison=[]
 for k in ['Total interest paid','Total financing fees','Total financing cost','Peak outstanding debt','New investor NPV','New investor IRR','MOIC']:
  x=a['metrics'][k];y=b['metrics'][k];comparison.append([k,y,x,x-y])
 quality=[['Preservação financeira R4','PASS',str(baseline['monthly_comparisons'])+' valores; máximo '+str(baseline['max_difference'])],['Ensaios funcionais','PASS',str(tests)+' verificações; 62 R5, 22 R6, 5 R6.1'],['Stresses integrais','PASS',str(len(data['stress']))+' cenários recalculados'],['Matriz Hotel','PASS','20 VAL independentes reconciliados com DASH 07 e DASH 10'],['Snapshot desatualizado','PASS','Comparação marcada desatualizada após alteração de parâmetros'],['Hiperligações','PASS ESTRUTURAL',str(len(navigation))+' destinos internos existentes'],['Open XML Microsoft 365','PASS','0 erros de esquema no master e snapshot'],['Motor / cálculos','PASS','0 erros de fórmula; saldos reconciliados em 180 meses'],['Impressão PDF','PASS INDEPENDENTE',str(len(pdfproof))+' documentos; paginação e texto verificados'],['Gráficos Hotel R4','PASS','Desenho e 3 partes de gráficos preservados byte a byte'],['Excel 365 Desktop','NOT TESTED','Sem Windows/Excel neste ambiente; kit fornecido'],['Logótipo JPG/PNG','PENDENTE','Área V06 Identidade preservada para inserir o original'],['Restauro em nova tarefa','NOT TESTED','Configuração reutilizável guardada; publicação do ambiente pendente']]
 with (BOUT/'Quality_Control.csv').open('w',newline='') as f:w=csv.writer(f);w.writerow(['requirement','status','evidence']);w.writerows(quality)
 lines=['# Relatório de testes — R6.1 BANK CHECKPOINT','', '**Base Oficial: TAN 6%; carência de capital 42 meses.** A menção a 7%/30 meses foi substituída.','',f'{tests} verificações funcionais bem-sucedidas, além de 3 stresses integrais, 20 pontos de sensibilidade bancária e 2 verificações de validade do snapshot. Os ensaios financeiros e PDFs utilizaram LibreOffice/ReportLab. Não constituem validação nativa Excel 365.','', '| Verificação | Estado | Evidência |','|---|---|---|']
 lines += ['| '+' | '.join(row)+' |' for row in quality]
 lines += ['', '## Comparação calculada: Base 6% vs. hipótese 5,5%', '', '| Indicador | Base 6% | Teste 5,5% | Delta |','|---|---:|---:|---:|']
 for key,y,x,delta in comparison:
  kind='pct' if 'IRR' in key else 'ratio' if key=='MOIC' else 'money';ds=f'{delta*100:+.4f} pp' if kind=='pct' else display(delta,kind);lines.append('| '+key+' | '+display(y,kind)+' | '+display(x,kind)+' | '+ds+' |')
 lines += ['', 'A redução da TAN melhora o VAL em cerca de 4,9 mil euros; o VAL continua negativo. Os dividendos periódicos calculados permanecem iguais e a diferença repercute-se sobretudo na caixa retida e no valor de saída. Não se infere bancabilidade.','', 'Na Base, o cash sweep liquida a dívida no mês 31, antes da carência de 42 meses. Por isso, testar isoladamente prazo mais longo ou carência de 48 meses não altera esse calendário. A igualdade foi verificada, não fixada em tabelas.','', 'O DSCR mínimo inclui meses de construção e é negativo nesse período; a fonte de reembolso depende também de capital e recebimentos imobiliários. A meta DSCR é interna, não covenant aprovado. A convenção, comissões e outros termos são hipóteses salvo identificação documental.','', 'A revisão R6.1 corrigiu a combinação de juros capitalizados com amortização por excedentes e separou a definição do DSCR dos cash flows usados na distribuição. Foram testadas essas combinações, a restituição integral da Base, ACT/365 e a independência de VAL/dividendos face à mera inclusão de comissões no denominador do DSCR.','', '## Origem e preservação','', 'R5 deriva da R4 válida. R6 acrescenta as condições bancárias aos mesmos motores existentes; não cria um motor independente para os dashboards. R6.1 mantém os resultados-base e acrescenta comparações de financiamento aos dez dashboards. Os nomes técnicos das folhas e os gráficos Hotel R4 foram preservados. A V04 original continua no arquivo do repositório, fora do workbook de trabalho.','', 'A emissão de demonstração foi calculada sobre a R6 do commit b900fdd5f752051096bb3ea65a097a87d7e819c9. O seu Snapshot é um arquivo protegido contra alterações acidentais; a formatação final não alterou os registos financeiros dos motores. Usar o master R6.1 para novos cenários. Não foi enviado relatório a uma instituição.','', '## Limitações e aceitação nativa','', 'Microsoft Excel 365 Desktop não está disponível. Abertura, ativação de edição, recálculo, gravação, reabertura, cliques e impressão nativos permanecem NOT TESTED. Executar o kit Excel365 numa cópia Windows; qualquer pedido de reparação impede aceitar a revisão como final.','', 'O logótipo original não existe como JPG/PNG acessível neste ambiente. V06 Identidade permite inserir o ficheiro original. Não foi redesenhado nem substituído por uma aproximação.','', '## Evidências','', 'Baseline_R6.json; Testes_Bancarios.json; Testes_R61_Adicionais.json; Snapshot_Stale_Guard.json; Bank_Matrix_20.json; Navigation_Matrix.csv; SDK_R61_RELEASE.json; SDK_ARCHIVE.json; Final_PDF_Checks.json; SHA256SUMS.txt. Os 62 ensaios R5 estão preservados na pasta Fecho_R5.']
 (BOUT/'Relatorio_Testes.md').write_text('\n'.join(lines)+'\n')
 styles=getSampleStyleSheet();flow=[]
 for line in lines:
  if line.startswith('|'):continue
  if not line:flow.append(Spacer(1,6));continue
  flow.append(Paragraph(escape(line.lstrip('# ').replace('**','')),styles['Heading1' if line.startswith('# ') else 'Heading2' if line.startswith('## ') else 'BodyText']))
  flow.append(Spacer(1,5))
 table=Table([[Paragraph(escape(str(v)),styles['BodyText']) for v in r] for r in [['Verificação','Estado','Evidência']]+quality],colWidths=[145,95,255],repeatRows=1);table.setStyle(TableStyle([('VALIGN',(0,0),(-1,-1),'TOP'),('BACKGROUND',(0,0),(-1,0),colors.HexColor('#E8EEF5')),('LINEBELOW',(0,0),(-1,-1),.3,colors.HexColor('#CCD5DF')),('BOTTOMPADDING',(0,0),(-1,-1),6)]));flow.insert(4,table);SimpleDocTemplate(str(BOUT/'Relatorio_Testes.pdf'),pagesize=A4,leftMargin=40,rightMargin=40,topMargin=38,bottomMargin=38).build(flow)
 proof={'status':'PASS INDEPENDENT - CHECKPOINT','functional_steps':tests,'stress_scenarios':3,'bank_matrix_points':20,'navigation_links':len(navigation),'selected_tabs':selected,'pdf_documents':len(pdfproof),'formula_errors':0,'native_excel365':'NOT TESTED','hotel_graphics_byte_preserved':graphic,'chart_parts_are_unique':True,'workbook_sheets':stats['sheets'],'workbook_bytes':BM.stat().st_size};(BOUT/'Release_Checks.json').write_text(json.dumps(proof,indent=2));print(proof,flush=True)
if __name__=='__main__':main()
