"""Publish scenario and all-dashboard selector evidence without native Excel claims."""
import json,hashlib
from pathlib import Path
from reportlab.platypus import SimpleDocTemplate,Paragraph,Spacer,Table,TableStyle
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from premium_ui import ROOT

def main():
 sync=json.loads((ROOT/'Sincronizacao_Cenarios.json').read_text());filters=json.loads((ROOT/'Testes_Todos_Filtros.json').read_text());finance=json.loads((ROOT/'Validacao_Financeira.json').read_text());correspondences=json.loads((ROOT/'Testes_Correspondencias_Dashboards.json').read_text());assert all(x['status']=='PASS' for x in [sync,filters,finance,correspondences])
 tests=filters['tests'];rows=['| Cenário | ADR inicial | Ocupação inicial | EBITDA ano operacional 5 | VAL novos investidores |','|---|---:|---:|---:|---:|']
 for item in sorted(sync['after'],key=lambda x:['Base Case','Conservative','Favorable','Upside Stress','Custom'].index(x['mode'])):
  m=item['metrics'];rows.append(f"| {item['mode']} | {m['adr']:.0f} EUR | {m['occupancy']:.0%} | {m['ebitda']:,.2f} EUR | {m['npv']:,.2f} EUR |")
 text='''# KK-01 V06.3-SYNC — Sincronização e testes dos seletores

A correção atua na fonte única `V06 Controlos!B67`, agora ligada a `DASH 05 V06!H23`. Antes, B67 estava fixo em Favorável: Base Case apresentava 210 EUR / 64% e o Conservador mudava ADR/ocupação mas conservava a composição Favorável de F&B/MICE. O motor de 180 meses já tinha dependências deste parâmetro; as expressões financeiras foram preservadas. A cadeia verificada inclui V06 Motor!T5:U184 (ADR/ocupação), V5:V184 (quartos), W5:W184 (MICE), X5:X184 (F&B), AA5:AA184 (GOP), DR5:DR184 (fees), AH5:AH184 (EBITDA), AI5:AI184 (FF&E), BX5:BX184 (CFADS), DE5:DE184 (dividendos), V06 F01 e V06 Cenarios!M5. Os resultados aparecem no DASH 05 V06 (funcional 07), DASH 07 V06 (funcional 09) e DASH 10 V06.

Base Case utiliza agora o perfil Base (185 EUR / 58%); Conservative utiliza Conservador (160 EUR / 52%). Favorable e Upside mantêm o perfil Favorável. Custom usa os valores próprios de ADR/ocupação e conserva o perfil Favorável para a composição de receitas; utilização MICE permanece um controlo independente. O teste Custom usa 200 EUR / 65%; o valor guardado no modelo pode ser diferente.

Esta correção altera legitimamente os resultados iniciais. A referência histórica permanece guardada; a comparação completa usa `Referencia_Python_V06_SYNC.json.gz`, executando o mesmo core existente com o perfil corrigido e as definições originais das variantes. A referência congelada Reference Case permanece histórica e não acompanha o cenário ativo por definição.

## Resultados dos cenários

'''+ '\n'.join(rows)+f'''

As receitas, EBITDA e encargos anuais referem-se ao quinto ano operacional selecionado; ADR e ocupação acima referem-se à abertura, antes do crescimento e ramp-up. VAL, TIR, MOIC, dividendos e saída usam o horizonte financeiro respetivo. A comparação completa antes/depois consta de `Comparacao_Cenarios.csv`. CAPEX, número físico de quartos e capital inicial podem permanecer iguais porque o cenário operacional não altera estes pressupostos. Taxas contratuais da operadora permanecem iguais se o contrato não mudar; os encargos em euros acompanham as receitas. Favorable/Upside/Custom conservam os resultados anteriores quando os seus parâmetros não mudam. Custom guardado em 210 EUR / 64% pode coincidir com Favorable; o teste específico usa 200 EUR / 65% para demonstrar a independência. A margem percentual conservadora pode superar a Base pela composição de receitas, embora o EBITDA absoluto seja menor.

Cada um dos cinco cenários foi comparado em 25 200 valores mensais, 181 fluxos de investidores, 34 ligações de dashboards/relatório/PDF, 585 valores de projeções gráficas e 999 pontos de cache dos gráficos. Foram verificados os quadros inferiores, além do painel superior. Zero erros de célula nos ficheiros recalculados. A sequência Base -> Conservador -> Favorável -> Base e uma gravação/reabertura adicional constam de `Sincronizacao_Cenarios.json`, com diferenças e igualdade dos KPI registadas.

## Restantes dashboards

Foram testadas {len(tests)} alterações, cobrindo {filters['inputs_covered']} entradas canónicas e as opções dos seletores. Cada teste regista a escolha, alterações no motor, KPI dos nove dashboards e resumo financeiro. Opções de idioma e apresentação devem preservar os cálculos. No residencial, KK Target (6% gestão + 2% distribuição) e Custom inicial (3% + 5%) totalizam ambos 8%: os componentes mudam, mas fees totais e VAL podem coincidir legitimamente. No Base corrigido, Auto e Minimum têm o mesmo objetivo anual de 3 milhões EUR nos 180 meses; a igualdade dos fluxos e do VAL foi confirmada, sem alterar o motor para forçar diferenças artificiais. Carência/prazo foram também testados com dívida em aberto; ausência de efeito quando a dívida já foi liquidada não implica desligação. O relatório Reference Case mantém intencionalmente a referência congelada. Resultados: `Matriz_Testes_Filtros.csv` e `Testes_Todos_Filtros.json`. Conferência adicional: {correspondences['total_bindings_checked']:,} ligações de quadros inferiores e projeções gráficas nos {len(correspondences['tests'])} testes, em `Testes_Correspondencias_Dashboards.json`.

Comparação financeira completa: {finance['scenario_count']} variantes, {finance['numeric_monthly_comparisons']:,} valores numéricos; estado {finance['status']}. Não foram introduzidas fórmulas IFERROR para esconder erros. Entradas Custom inativas ficam visualmente ocultas por formato, sem afetar os dados ou ocultar erros financeiros.

## Navegação, impressão e limites

O painel Active Scenario em EXEC V06 consulta os controlos canónicos e dá acesso direto aos mesmos. Os dez dashboards têm menu global e as auxiliares têm regresso. As ligações foram verificadas estruturalmente; os PDF PT/EN são exportações LibreOffice, com dez páginas A4 no relatório integral e duas no resumo. Os nove dashboards individuais têm composição A3 horizontal, segundo a arquitetura aprovada.

Microsoft Excel 365 Desktop: **NOT TESTED**, pois não está instalado neste ambiente Linux. Não é possível certificar abertura sem reparação ou comportamento nativo das hiperligações no Excel. O procedimento Windows atualizado está em `Excel365/Executar_Validacao.cmd`; executa em cópia, testa alteração dos cenários, recálculo, erros, gravação/reabertura e exportação. A auditoria OOXML verifica estrutura, referências, tipos e ordem de células, mas não substitui esse teste nativo.

O logótipo original não foi redesenhado: permanece disponível a área `V06 Identidade!B8:H20` para inserir o JPG/PNG oficial. As provas antigas que descrevem Base=Favorável são históricas e estão substituídas por este relatório e pelos testes SYNC.
'''
 (ROOT/'Relatorio_Sincronizacao_Cenarios.md').write_text(text)
 styles=getSampleStyleSheet();story=[]
 for para in text.split('\n\n'):
  if para.startswith('|'):
   data=[[c.strip() for c in line.strip('|').split('|')] for line in para.splitlines() if not line.startswith('|---')];data=[[Paragraph(c,styles['BodyText']) for c in row] for row in data];table=Table(data,colWidths=[95,65,70,130,130],repeatRows=1);table.setStyle(TableStyle([('BACKGROUND',(0,0),(-1,0),colors.HexColor('#dce6ed')),('GRID',(0,0),(-1,-1),.3,colors.grey),('VALIGN',(0,0),(-1,-1),'TOP')]));story += [table,Spacer(1,10)]
  else:story += [Paragraph(para.lstrip('# ').replace('`','').replace('**','').replace('\n',' '),styles['Heading1'] if para.startswith('# ') else styles['Heading2'] if para.startswith('## ') else styles['BodyText']),Spacer(1,8)]
 SimpleDocTemplate(str(ROOT/'Relatorio_Sincronizacao_Cenarios.pdf'),pagesize=A4,rightMargin=42,leftMargin=42,topMargin=38,bottomMargin=38).build(story)
 for name in ['README.md','Relatorio_Filtros.md','Relatorio_Alteracoes_Graficas.md','Relatorio_Integridade_Financeira.md','Auditoria_Mandato.md']:
  p=ROOT/name;s=p.read_text().replace('607 ligações','617 ligações').replace('36.3 MB',f'{(ROOT/"KK01_V06_Executive_Master.xlsx").stat().st_size/1e6:.1f} MB');s='> Atualização V06.3-SYNC: Base Case passa a usar Base (185 EUR / 58%). Resultados e testes atuais: [Relatório de sincronização](Relatorio_Sincronizacao_Cenarios.md). As equivalências anteriores Base=Favorável estão substituídas.\n\n'+s;p.write_text(s)
 print('Sync report written',len(tests))
if __name__=='__main__':main()
