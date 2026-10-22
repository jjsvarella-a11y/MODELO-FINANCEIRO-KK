"""Record native repair evidence and a strictly structural, cell-preserving fix."""
import json,hashlib
from pathlib import Path
from reportlab.platypus import SimpleDocTemplate,Paragraph,Spacer
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.pagesizes import A4
ROOT=Path(__file__).resolve().parents[1]
def main():
 repair=json.loads((ROOT/'Correcao_Reparacao_XML.json').read_text());assert repair['status']=='PASS';sha=hashlib.sha256((ROOT/'KK01_V06_Executive_Master.xlsx').read_bytes()).hexdigest();assert sha==repair['after_sha256']
 for filename,field in [('Integridade_Premium.json','sha256'),('Sincronizacao_Cenarios.json','result_sha256'),('Preservacao_Motor_SYNC.json','after_sha256')]:
  p=ROOT/filename;obj=json.loads(p.read_text());obj['tested_pre_structural_patch_sha256']=obj[field];obj[field]=sha;obj['structural_patch_evidence']='Correcao_Reparacao_XML.json';obj['carry_forward_reason']='All cell records, financial formulas/values and chart parts preserved; only three invalid empty XML containers removed';p.write_text(json.dumps(obj,indent=2,ensure_ascii=False))
 p=ROOT/'Validacao_Financeira.json';obj=json.loads(p.read_text());obj['structural_patch_evidence']='Correcao_Reparacao_XML.json';obj['recalculation_after_R1']='Not repeated: exact preservation of all financial cell records and engine/chart ZIP parts verified';p.write_text(json.dumps(obj,indent=2,ensure_ascii=False))
 text=f'''# KK-01 V06.3-SYNC-R1 — Correção do aviso de reparação

O utilizador apresentou uma captura do Microsoft Excel que confirma reparação na revisão anterior de três partes XML: sheet492.xml, sheet323.xml e sheet328.xml. A auditoria anterior verificava XML bem formado, mas não verificava a cardinalidade obrigatória dos filhos destes elementos. Essa lacuna foi corrigida.

Foram encontrados exatamente três elementos vazios nas mesmas partes: hyperlinks em V06 Identidade (sheet492.xml), hyperlinks em V06 Controlos (sheet323.xml) e dataValidations em V06 Hipoteses (sheet328.xml). Nestes contentores SpreadsheetML, o elemento filho é obrigatório quando o contentor existe. As hiperligações de navegação já são fórmulas HYPERLINK; os contentores vazios eram resíduos da migração da navegação e das validações.

Foram removidos apenas esses três contentores. Não foram removidas hiperligações funcionais nem validações com opções. A correção foi integrada no gerador para impedir a sua recriação.

Verificações: 312 partes ZIP permanecem idênticas byte a byte; nas três partes modificadas, os 1.002 registos de células permanecem idênticos. A comparação integral das árvores XML confirma que as únicas diferenças são as três remoções. Todas as fórmulas, valores guardados, formatos, imagens, gráficos, definições de impressão e resultados dos testes de cenários foram preservados. A nova auditoria verifica os contentores vazios em todas as 190 folhas. A navegação foi novamente verificada estruturalmente (617 ligações).

Os resultados financeiros e os 73 testes de filtros anteriores são mantidos por preservação exata dos dados; não foram repetidos nesta correção estritamente estrutural. Os PDFs permanecem iguais porque a apresentação e os valores não mudaram. A evidência detalhada está em Correcao_Reparacao_XML.json e Auditoria_OOXML.json.

Microsoft Excel 365: a versão anterior teve reparação confirmada pelo utilizador. A R1 ainda não foi reaberta no Excel Desktop neste ambiente, onde esse programa não está disponível. A eliminação dos três defeitos foi verificada; isso não constitui certificação nativa de ausência de outros avisos. O procedimento Windows continua disponível em Excel365/Executar_Validacao.cmd.

Descarregar novamente KK01_V06_Executive_Master.xlsx do diretório habitual. Não continuar sobre a cópia que o Excel já reparou. Identificador da revisão: V06.3-SYNC-R1. SHA-256 do Excel: {sha}.
'''
 (ROOT/'Relatorio_Reparacao_XML_R1.md').write_text(text)
 styles=getSampleStyleSheet();story=[]
 for para in text.split('\n\n'):
  story.extend([Paragraph(para.lstrip('# ').replace('\n',' '),styles['Heading1'] if para.startswith('#') else styles['BodyText']),Spacer(1,10)])
 SimpleDocTemplate(str(ROOT/'Relatorio_Reparacao_XML_R1.pdf'),pagesize=A4,leftMargin=42,rightMargin=42,topMargin=38,bottomMargin=38).build(story)
 banner='> **Revisão R1:** removidos três contentores XML vazios nas folhas indicadas pelo Excel. [Correção e testes](Relatorio_Reparacao_XML_R1.md). Cálculos e apresentação preservados; reabertura nativa da R1 ainda não testada.\n\n'
 p=ROOT/'README.md';p.write_text(banner+p.read_text())
 p=ROOT.parents[1]/'README.md';p.write_text(banner.replace('(Relatorio_Reparacao_XML_R1.md)','(entregas/EFCC_Final/Relatorio_Reparacao_XML_R1.md)')+p.read_text())
 print('R1 report and preservation metadata updated',sha)
if __name__=='__main__':main()
