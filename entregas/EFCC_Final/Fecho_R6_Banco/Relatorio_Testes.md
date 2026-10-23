# Relatório de testes — R6.1 BANK CHECKPOINT

**Base Oficial: TAN 6%; carência de capital 42 meses.** A menção a 7%/30 meses foi substituída.

89 verificações funcionais bem-sucedidas, além de 3 stresses integrais, 20 pontos de sensibilidade bancária e 2 verificações de validade do snapshot. Os ensaios financeiros e PDFs utilizaram LibreOffice/ReportLab. Não constituem validação nativa Excel 365.

| Verificação | Estado | Evidência |
|---|---|---|
| Preservação financeira R4 | PASS | 1713600 valores; máximo 1.0477378964424133e-09 |
| Ensaios funcionais | PASS | 89 verificações; 62 R5, 22 R6, 5 R6.1 |
| Stresses integrais | PASS | 3 cenários recalculados |
| Matriz Hotel | PASS | 20 VAL independentes reconciliados com DASH 07 e DASH 10 |
| Snapshot desatualizado | PASS | Comparação marcada desatualizada após alteração de parâmetros |
| Hiperligações | PASS ESTRUTURAL | 972 destinos internos existentes |
| Open XML Microsoft 365 | PASS | 0 erros de esquema no master e snapshot |
| Motor / cálculos | PASS | 0 erros de fórmula; saldos reconciliados em 180 meses |
| Impressão PDF | PASS INDEPENDENTE | 30 documentos; paginação e texto verificados |
| Gráficos Hotel R4 | PASS | Desenho e 3 partes de gráficos preservados byte a byte |
| Excel 365 Desktop | NOT TESTED | Sem Windows/Excel neste ambiente; kit fornecido |
| Logótipo JPG/PNG | PENDENTE | Área V06 Identidade preservada para inserir o original |
| Restauro em nova tarefa | NOT TESTED | Configuração reutilizável guardada; publicação do ambiente pendente |

## Comparação calculada: Base 6% vs. hipótese 5,5%

| Indicador | Base 6% | Teste 5,5% | Delta |
|---|---:|---:|---:|
| Total interest paid | 783,200 EUR | 715,567 EUR | -67,633 EUR |
| Total financing fees | 1,125,000 EUR | 1,125,000 EUR | 0 EUR |
| Total financing cost | 1,908,200 EUR | 1,840,567 EUR | -67,633 EUR |
| Peak outstanding debt | 17,528,049 EUR | 17,488,924 EUR | -39,125 EUR |
| New investor NPV | -12,742,848 EUR | -12,737,906 EUR | 4,943 EUR |
| New investor IRR | 5.76% | 5.77% | +0.0038 pp |
| MOIC | 2.12x | 2.12x | 0.00x |

A redução da TAN melhora o VAL em cerca de 4,9 mil euros; o VAL continua negativo. Os dividendos periódicos calculados permanecem iguais e a diferença repercute-se sobretudo na caixa retida e no valor de saída. Não se infere bancabilidade.

Na Base, o cash sweep liquida a dívida no mês 31, antes da carência de 42 meses. Por isso, testar isoladamente prazo mais longo ou carência de 48 meses não altera esse calendário. A igualdade foi verificada, não fixada em tabelas.

O DSCR mínimo inclui meses de construção e é negativo nesse período; a fonte de reembolso depende também de capital e recebimentos imobiliários. A meta DSCR é interna, não covenant aprovado. A convenção, comissões e outros termos são hipóteses salvo identificação documental.

A revisão R6.1 corrigiu a combinação de juros capitalizados com amortização por excedentes e separou a definição do DSCR dos cash flows usados na distribuição. Foram testadas essas combinações, a restituição integral da Base, ACT/365 e a independência de VAL/dividendos face à mera inclusão de comissões no denominador do DSCR.

## Origem e preservação

R5 deriva da R4 válida. R6 acrescenta as condições bancárias aos mesmos motores existentes; não cria um motor independente para os dashboards. R6.1 mantém os resultados-base e acrescenta comparações de financiamento aos dez dashboards. Os nomes técnicos das folhas e os gráficos Hotel R4 foram preservados. A V04 original continua no arquivo do repositório, fora do workbook de trabalho.

A emissão de demonstração foi calculada sobre a R6 do commit b900fdd5f752051096bb3ea65a097a87d7e819c9. O seu Snapshot é um arquivo protegido contra alterações acidentais; a formatação final não alterou os registos financeiros dos motores. Usar o master R6.1 para novos cenários. Não foi enviado relatório a uma instituição.

## Limitações e aceitação nativa

Microsoft Excel 365 Desktop não está disponível. Abertura, ativação de edição, recálculo, gravação, reabertura, cliques e impressão nativos permanecem NOT TESTED. Executar o kit Excel365 numa cópia Windows; qualquer pedido de reparação impede aceitar a revisão como final.

O logótipo original não existe como JPG/PNG acessível neste ambiente. V06 Identidade permite inserir o ficheiro original. Não foi redesenhado nem substituído por uma aproximação.

## Evidências

Baseline_R6.json; Testes_Bancarios.json; Testes_R61_Adicionais.json; Snapshot_Stale_Guard.json; Bank_Matrix_20.json; Navigation_Matrix.csv; SDK_R61_RELEASE.json; SDK_ARCHIVE.json; Final_PDF_Checks.json; SHA256SUMS.txt. Os 62 ensaios R5 estão preservados na pasta Fecho_R5.
