# KK-01 — R6.1 BANK CHECKPOINT

Base Oficial confirmada pelo utilizador: **TAN 6% e carência de capital 42 meses**. A menção anterior a 7%/30 meses foi substituída. O prazo de 120 meses é contado **após** a carência; o último mês de utilização continua a ser 30, campo independente da carência.

[Excel de trabalho](KK01_Executive_Financial_Control_Center_R6_1_BANK_CHECKPOINT.xlsx) · [Matriz de navegação](Navigation_Matrix.csv) · [Implementação](Bank_Implementation.json) · [Procedimento Excel 365](Excel365/LEIA-ME.md).

Esta revisão integra o fecho R5 dos dez dashboards e acrescenta o simulador de financiamento e a emissão controlada de relatórios bancários. É **CHECKPOINT**: não há Microsoft Excel 365 Desktop neste ambiente. Os resultados executados e as limitações devem ser consultados em `Relatorio_Testes.md`; a validação de esquema não substitui abertura, edição, gravação e reabertura nativas.

## Base e cenários

Há dois controlos diferentes:

- **EXEC V06!H23 — Reference / Test**: Reference restitui os pressupostos financeiros R4 protegidos; Test aplica os parâmetros guardados, sem os perder ao alternar o modo.
- **DASH 04 funcional, folha técnica DASH 03 V06!H33**: escolhe Official KK-01 6%, Proposal Bank A, Proposal Bank B, Custom ou oito posições Saved. **N33** escolhe o perfil de comparação para emissão.

A Base Oficial protegida usa os mesmos pressupostos bancários R4: linha 30 M€, TAN 6%, carência 42 meses, prazo subsequente 120 meses, utilização até M30, amortização Linear, cash sweep 100%, abertura 1%, intermediação 1,5%, disponibilização 1,25%, compromisso 0%. Estes valores provêm do modelo, **não de aprovação bancária documental**.

A/B são **hipóteses de estudo**, identificadas como Hypothetical até existir proposta: A usa TAN 5,5%/carência 42; B usa TAN 6,5%/carência 48. Não são propostas recebidas de bancos identificados. Não são automaticamente classificadas como melhores.

**Custom** utiliza os campos editáveis do dashboard. Selecionar outro perfil deixa esses campos guardados, mas eles não comandam o motor: o painel mostra o perfil/TAN efetivamente ativos. Para guardar sem sobrescrever, copiar `V06 Banco Perfis!B11:AI11` e **colar valores** numa posição Saved; completar instituição, estatuto e referência documental. A linha Oficial está protegida e nunca é adotada/substituída automaticamente.

## DASH 04 — condições e resultados

Os controlos originais mantêm endereços: H23 TAN fixa, N23 carência, H25 prazo após carência, N25 fração de cash sweep, H27 último mês de utilização, H29 Linear/Bullet, H31 limite. O painel superior liga à área de condições **E450:Y489**, aos gráficos bancários **E492:Y574** e à folha de serviço da dívida.

Os campos adicionais incluem tipo fixa/variável, indexante/spread, carência de juros separada, abertura percentual/fixa, intermediação, disponibilização, compromisso e base de incidência, gestão/periodicidade, antecipação, liquidação, garantias, despesas, seguro de crédito, janela bancária, convenção 30/360 ou ACT/365, fonte de amortização antecipada, limite Custom e meta DSCR.

A TAN variável é indexante + spread **constantes no cenário**; não existe uma curva futura importada. 30/360 utiliza TAN/12; ACT/365 usa dias reais do calendário indicado. Juros incidem sobre dívida inicial + desembolso do mês, conforme convenção original de utilização no início do período. Durante a carência de juros há **capitalização**, não perdão nem ocultação do encargo. Esse saldo entra na dívida e no cálculo dos juros seguintes.

Abertura, intermediação e disponibilização são cobradas uma vez, na primeira necessidade de crédito, com incidência no limite autorizado conforme a R4. Abertura fixa soma ao componente percentual. Compromisso usa a base selecionada e não é duplicado na abertura. Gestão é cobrada na janela de disponibilidade pela periodicidade indicada. Garantias/despesas/seguro adicional são cobrados no primeiro mês permitido; o seguro é adicional aos seguros imobiliários já existentes. Antecipação incide na amortização antecipada efetiva e liquidação fixa apenas no mês em que se extingue dívida utilizada.

O cronograma `V06 Banco Servico` mostra os 180 meses e uma vista agrupada **Monthly / Quarterly / Annual** em E320. Há distinção entre juros devidos/pagos/capitalizados, principal programado/antecipado, comissões, serviço contratual e pagamentos bancários totais. DSCR é N/A quando não há serviço.

**Definição de estudo do DSCR:** CFADS do motor, antes de imposto e caixa residencial, dividido por principal programado + juros pagos + compromisso. As demais comissões entram apenas se o seletor respetivo for 1. A meta configurável é interna, não covenant aprovado. A amortização antecipada não é confundida com serviço contratual programado; está incluída nos pagamentos bancários totais. Validar a definição num contrato real antes de a tratar como covenant.

As comissões de antecipação/liquidação afetam caixa, lucro tributável, folga para dividendos e défice de financiamento. O limite de antecipação considera dívida, caixa, reservas e custo da antecipação. Não se inventa financiamento para cobrir um défice: ele permanece visível.

## DASH 10 — emissão controlada

No centro do relatório, **E90:Y123**, estão cenário/comparação, formato, apresentação e destinatário. Cenário e comparação são os seletores canónicos do DASH 04; os atalhos abrem-nos sem criar entradas concorrentes.

O emissor é externo ao Excel, para manter o livro `.xlsx` sem dependência obrigatória de VBA. Hiperligações nativas abrem folhas; não executam programas nem acionam a impressão. No ambiente Linux:

```bash
python entregas/EFCC_Final/src/bank_emit.py \
  --recipient 'Instituição destinatária' \
  --scenario 'Proposal Bank A' \
  --comparison 'Official KK-01 6%' \
  --stress
```

Requer Python, lxml, XlsxWriter, ReportLab, Matplotlib, PyMuPDF e LibreOffice. Sem os argumentos scenario/comparison, são usados os seletores gravados no Excel; sem recipient, é usado DASH 10 Centro!H122, que deve estar identificado. `--selected-only` omite a comparação no PDF. O modo Base vs Test é o padrão. O formato gravado no DASH 10 é respeitado. `--format bank`, `brief`, `consolidated` ou `all` permite escolher os documentos. A demonstração entregue inclui os três formatos.

Cada perfil é recalculado integralmente no **mesmo motor de 180 meses**. O emissor procura erros em todas as folhas, verifica saldos de caixa/dívida/reservas/inventário, limite de utilização, pagamentos, comissões e VAL a partir dos 181 cash flows. Só depois produz PDFs A4 PT/EN. Com `--stress`, recalcula obra +6 meses, operação Conservador e recebimentos +6 meses. Os stresses são condições explícitas de estudo; não uma certificação exaustiva de risco.

A emissão guarda Snapshot.xlsx, pressupostos/resultados JSON, CSV de 180 meses, data UTC, identificador e manifesto SHA256. Não sobrescreve emissões nem o original. Uma falha deixa BLOCKED.json e remove os PDFs dessa emissão. O exportador não aceita caches como prova de recálculo. A proteção aplica-se a este percurso de emissão; o comando Ficheiro/Imprimir do próprio Excel continua sob controlo do utilizador.

O custo efetivo anual é IRR dos fluxos líquidos de financiamento, anualizada, incluindo encargos modelados; não é declarado como TAEG contratual/regulamentar. Se ainda houver dívida no fim de 180 meses, é N/A para evitar apresentar uma taxa de vida completa sem fluxos posteriores. Os encargos devidos/capitalizados e a dívida remanescente continuam visíveis. A taxa de desconto dos acionistas permanece independente da TAN.

## Fecho funcional incluído

Os dez dashboards mantêm sidebar, Home, parâmetros e impressão individual. Os auxiliares acessíveis têm caminhos de regresso. O idioma usa o seletor global. Os nomes técnicos não foram alterados. O relatório consolidado tem dez páginas A4 e o resumo duas; os painéis individuais usam A3 horizontal conforme a arquitetura anterior.

Estão incluídos Impacto Global Reference/Active/Delta nos dez painéis, filtros ligados ao motor, capítulos/ETC e curva ajustada de construção, vista de vendas por tipologia/área/preço/estado previsto, condições residenciais, políticas de reservas, desconto único, comparação de investidores e a matriz Hotel de 20 VAL reais. Os gráficos Hotel R4 preservados não foram reconstruídos nesta extensão bancária.

Os direitos económicos adquiridos JSCV/JALS não são transformados em investimento histórico fictício. Existem 25 M€ de subscrições previstas; 0€ documentados nos campos de recebimento do ficheiro não demonstra ausência de pagamentos reais. A distribuição segue as regras do modelo e não é presumida pro rata.

O logótipo original continua sem bytes JPG/PNG acessíveis; permanece a área **V06 Identidade** para inserir o original diretamente no Excel. Não foi redesenhado. A eventual cópia para cabeçalhos e o acabamento após inserção precisam de verificação nativa.

## Entrega e evidências

[Relatório de testes PDF](Relatorio_Testes.pdf) · [Relatório detalhado](Relatorio_Testes.md) · [Quality Control](Quality_Control.csv) · [PDFs PT](PDF/PT) · [PDFs EN](PDF/EN) · [Emissão de demonstração 6% vs. 5,5%](Emissao_Demonstracao_6_vs_5_5/LEIA-ME.md) · [Recuperação](KK01_R61_Recovery.zip).

A R6.1 inclui comparação de financiamento nos dez dashboards e distingue claramente campos Custom guardados dos parâmetros ativos. O arquivo de emissão preserva o cenário calculado na R6; usar o master R6.1 para simulações novas. O custo efetivo não é TAEG regulamentar. O DSCR de estudo inclui construção e não é covenant aprovado.

Durante carência de juros, os encargos são capitalizados e deduzidos na base fiscal por competência, hipótese contabilística/fiscal a confirmar antes de adotar condições contratuais. A data indicativa M1 é uma hipótese editável usada somente em ACT/365.

A configuração do ambiente foi guardada (`install_script` e `start_skill`); requer revisão, gravação e publicação nas definições do ambiente. A instância atual está validada; o restauro numa tarefa nova ainda não foi testado.
