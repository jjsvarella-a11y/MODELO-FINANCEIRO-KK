# KK-01 — EXECUTIVE FINANCIAL CONTROL CENTER
## MASTER PROMPT DE CONTINUIDADE, CORREÇÃO E CONSOLIDAÇÃO FINAL
### Excel 365 | Integridade Financeira | Arquitetura Aprovada | Cenários | Reporting

**DATA DE REFERÊNCIA:** 25/10/2026  
**PRIORIDADE:** CRÍTICA  
**MODALIDADE:** Execução integral, sem aprovações intermédias  
**FICHEIRO DE TRABALHO:** `KK01_Executive_Financial_Control_Center.xlsx`  
**GABARITO VISUAL OBRIGATÓRIO:** `Markdown colado.md`

---

# 1. MISSÃO

Dar continuidade ao desenvolvimento do KK-01, aproveitando integralmente o trabalho já realizado e corrigindo os problemas identificados nas últimas revisões.

**Não iniciar um novo projeto. Não reconstruir o workbook. Não reinterpretar o desenho aprovado.**

O objetivo é entregar um único modelo consolidado que cumpra simultaneamente:

1. Integridade dos cálculos financeiros.
2. Eliminação dos erros `#VALOR!` cuja causa seja identificada.
3. Reprodução fiel da organização dos dez dashboards aprovada pelo utilizador.
4. Navegação direta para quadros detalhados e regresso funcional.
5. Base de Referência editável.
6. Cenários Comparativos independentes.
7. Filtros financeiros e operacionais predefinidos, com opção Custom.
8. Project Impact Summary atualizado.
9. Central de Relatórios Executivos dinâmica.
10. Impressão sempre ancorada na Base de Referência ativa e guardada.

O Codex é responsável pela implementação técnica, preservação da arquitetura e verificação interna.

---

# 2. HIERARQUIA OBRIGATÓRIA DAS ESPECIFICAÇÕES

Em caso de conflito entre documentos ou instruções anteriores, seguir esta ordem:

**NÍVEL 1 — Arquitetura visual**

O documento `Markdown colado.md` é vinculativo para:

- Layout dos dashboards.
- Hierarquia dos conteúdos.
- Sidebar.
- Cabeçalhos.
- Posição dos filtros.
- Project Impact Summary.
- KPI.
- Gráficos.
- Quadros financeiros.
- Navegação para detalhes.

**NÍVEL 2 — Funcionalidades posteriormente aprovadas**

Prevalecem as decisões recentes relativas a:

- Base de Referência editável.
- Cenários Comparativos editáveis.
- TAN predefinida de 2% a 8%.
- Filtros Hotel & MICE.
- Cenários guardados.
- Relatórios e impressão a partir da Base ativa.

**NÍVEL 3 — Motor financeiro existente**

Preservar os módulos financeiros e respetivas dependências, corrigindo apenas erros identificados ou incoerências demonstradas.

**NÍVEL 4 — Instruções técnicas anteriores**

Utilizar apenas quando não contrariem os níveis anteriores.

Não utilizar decisões antigas para substituir opções posteriormente aprovadas.

---

# 3. ESTABILIZAÇÃO DO FICHEIRO E CORREÇÃO DOS ERROS #VALOR!

A análise anterior identificou:

- 238 folhas.
- Aproximadamente 2,12 milhões de fórmulas.
- Ausência de erros armazenados no XML das células na cópia inspecionada.
- Recálculo automático e completo na abertura.
- 282 fórmulas que utilizam `VALUE()`, distribuídas entre `V07 Base` e `V07 Comparativo`.

Foi identificada uma hipótese concreta de falha na interpretação dos filtros percentuais armazenados como texto, particularmente quando apresentam ponto decimal.

Exemplo:

`V07 Base!D121`

```excel
=IF(B121="Custom",C121,IF(ISNUMBER(B121),B121,IF(RIGHT(B121,1)="%",VALUE(SUBSTITUTE(B121,"%",""))/100,VALUE(B121))))
```

Existe fórmula equivalente em:

`V07 Comparativo!D121`

Estas fórmulas estão ligadas à interpretação dos filtros que alimentam os módulos financeiros.

## 3.1 Diagnóstico

Confirmar a primeira célula efetivamente afetada quando o erro puder ser reproduzido.

Rastrear as dependências até ao input original.

Distinguir:

- Falha comprovada.
- Conversão potencialmente vulnerável.
- Erro propagado.
- Resultado armazenado desatualizado.

Não assumir automaticamente que todas as fórmulas `VALUE()` estão erradas.

## 3.2 Correção

Privilegiar valores numéricos nativos nos filtros.

Percentagens devem ser armazenadas como frações numéricas, por exemplo:

- 2% = 0,02.
- 5% = 0,05.
- 6% = 0,06.
- 8% = 0,08.

Aplicar formatação percentual sem transformar a taxa em texto.

Quando a conversão textual for inevitável, utilizar interpretação explícita e independente da configuração regional, respeitando o tipo de parâmetro.

Distinguir:

- Percentuais.
- Valores monetários.
- Números inteiros.
- Categorias.
- Perfis.
- Custom.

Preservar as fórmulas legítimas de seleção de categorias e perfis.

## 3.3 Restrições

Não utilizar `IFERROR` para ocultar falhas.

Não substituir resultados financeiros por zero.

Não alterar o valor económico dos pressupostos.

Não eliminar cenários.

Não modificar artificialmente o histórico para simular alterações efetuadas pelo utilizador.

Não reconstruir o motor financeiro.

## 3.4 Verificação

Após a correção, verificar os resultados dependentes de:

- TAN.
- Prazo.
- Carência.
- Fees bancários.
- ADR.
- Ocupação.
- Fees da operadora.
- Serviço da dívida.
- DSCR.
- Tesouraria.
- Dividendos.
- VAL.
- TIR.
- MOIC.
- Relatórios.

Distinguir verificações programáticas de testes efetivamente realizados no Microsoft Excel 365.

---

# 4. REPRODUÇÃO EXATA DO GABARITO APROVADO

A arquitetura dos dez dashboards está fechada.

**Não produzir uma nova proposta visual.**

Reproduzir a estrutura aprovada no documento `Markdown colado.md`, respeitando a ordem e organização dos elementos.

## Layout comum

- Sidebar à esquerda.
- Cabeçalho JS-CV.
- Logótipo institucional.
- Seletor PT/EN.
- Filtros específicos.
- Project Impact Summary.
- KPI principais.
- Gráficos financeiros.
- Quadros de análise.
- Alertas e conclusões.
- Acesso a informação detalhada.

## Identidade gráfica

- Azul-marinho: `#142337`.
- Grafite: `#252932`.
- Vermelho JS-CV: `#C8101E`.
- Fundo claro: `#F5F7FA`.
- Inputs editáveis: `#DCEAF8`.
- Tipografia Aptos.

Preservar os gráficos já ajustados manualmente sempre que possível.

Não recriar gráficos funcionais desnecessariamente.

## Layout principal

A página principal deve ser limpa, legível e adequada à visualização no Excel 365.

Informação extensa pode ser apresentada em folhas complementares, desde que exista acesso imediato a partir da página principal.

---

# 5. SIDEBAR PERMANENTE — NAVEGAÇÃO DASH 01 A 10

Manter a mesma estrutura em todos os dashboards:

01 — Executivo  
02 — Construção  
03 — Fiscalidade  
04 — Financiamento  
05 — Real Estate  
06 — Residencial  
07 — Hotel & MICE  
08 — Acionistas  
09 — Retornos  
10 — Relatórios

A sidebar deve:

- Manter dimensões consistentes.
- Destacar a página ativa.
- Ter hiperligações funcionais.
- Ser legível em PT/EN.
- Evitar títulos cortados.
- Permanecer visualmente fixa sempre que possível.

Não alterar a posição da sidebar entre dashboards.

---

# 6. NAVEGAÇÃO DIRETA PARA QUADROS DETALHADOS

Implementar a lógica:

**DASH Principal → Quadro Detalhado → Voltar ao DASH Principal**

## DASH 02 — Orçamento Sintético

O comando «Ver Orçamento Sintético» deve abrir diretamente o quadro do orçamento.

No ficheiro anteriormente analisado, o quadro estava aproximadamente em:

`'DASH 01 V06'!E190`

Confirmar o endereço na versão atual.

Não abrir simplesmente uma folha auxiliar no início.

## DASH 02 — Quarterly Disbursement

O comando «Ver Quarterly Disbursement» deve abrir diretamente o quadro trimestral.

Referência anterior:

`'DASH 01 V06'!E230`

Confirmar o endereço na versão atual.

## Comando de regresso

Cada vista detalhada deve incluir:

**VOLTAR AO DASH XX**

O comando deve regressar ao início visual do dashboard correspondente.

Aplicar este sistema também aos restantes dashboards com informação extensa.

---

# 7. ARQUITETURA DOS DEZ DASHBOARDS

A organização visual abaixo resume o gabarito, sem o substituir.

O documento `Markdown colado.md` determina a apresentação exata.

## DASH 01 — Dashboard Executivo

Apresentar:

- CAPEX.
- VGV.
- Equity.
- Pico de dívida.
- VAL.
- TIR/MOIC.

Gráficos:

- Fontes e aplicações.
- Liquidez e dívida.
- Desempenho financeiro global.

Alertas:

- Restrições financeiras.
- Meses críticos.
- Decisões pendentes.

## DASH 02 — Construção

Apresentar:

- Orçamento por capítulo.
- Incidência percentual.
- Ajustamentos de cenário.
- Orçamento Base/Teste.
- Desembolso trimestral.
- IVA.
- Acumulados.

Gráficos:

- Orçamento por capítulo.
- Quarterly Disbursement.
- Curva S / Budget Bridge.

Garantir acesso imediato ao orçamento integral e ao desembolso trimestral.

## DASH 03 — Fiscalidade

Apresentar:

- IVA suportado.
- IVA recuperável.
- Crédito por recuperar.
- IRPC.
- Isenções.
- Seguros.

Gráficos:

- Recuperação do IVA.
- IRPC anual.
- Impacto fiscal na liquidez.

## DASH 04 — Financiamento

Apresentar:

- Condições bancárias editáveis.
- Base e Comparativo.
- Serviço da dívida.
- CFADS.
- DSCR.
- Liquidez.
- Meses críticos.
- Custos de financiamento.

Gráficos:

- Drawdown & Repayment.
- Debt Service.
- CFADS vs Debt Service.
- Critical Liquidity Months.

## DASH 05 — Real Estate

Apresentar:

- Inventário das 94 unidades.
- Áreas.
- Tipologias.
- Pricing.
- VGV.
- Disponibilidade.
- Recebimentos.

Gráficos:

- Disponibilidade.
- Preço por tipologia.
- Recebimentos trimestrais.

## DASH 06 — Residencial

Apresentar:

- Receitas.
- OPEX.
- Fees.
- Yield dos proprietários.
- Reservas.
- Distribuições.

Gráficos:

- Owner Yield.
- Receita vs OPEX.
- Distribuições e reservas.

## DASH 07 — Hotel & MICE

Apresentar:

- ADR.
- Ocupação.
- RevPAR.
- Receita.
- EBITDA.
- Fees.
- FF&E.

Gráficos:

- Evolução de receitas.
- ADR × Ocupação / VAL.
- EBITDA e margem.

## DASH 08 — Acionistas

Apresentar:

- JSCV 40%.
- JALS 10%.
- Investidores 50%.
- Blocos de investimento.
- Captação.
- Direitos e privilégios.

Gráficos:

- Ownership Structure.
- Capital Raise Progress.
- Composição dos blocos.

## DASH 09 — Retornos

Apresentar:

- TIR.
- VAL.
- MOIC.
- Dividendos.
- Capital recuperado.
- Exit.
- Investor Impact Summary.

Gráficos:

- Dividendos anuais.
- Equity Cash Flow / Exit.
- Waterfall.

## DASH 10 — Relatórios

Apresentar:

- One-Page Executive Report.
- Bank Financing Report.
- Equity Investor Report.
- Comprehensive Project Report.

Incluir seleção de destinatário, cenário, idioma, comparação, anexos, confidencialidade e pré-visualização.

---

# 8. BASE DE REFERÊNCIA EDITÁVEL

A Base não é imutável.

Representa as condições financeiras e operacionais atualmente adotadas para estudo e negociação.

O utilizador pode modificar a Base através de filtros predefinidos ou Custom.

As alterações guardadas atualizam os resultados-base do projeto.

Preservar o histórico das versões anteriores.

Identificar sempre:

**BASE DE REFERÊNCIA — EDITÁVEL**

---

# 9. CENÁRIOS COMPARATIVOS EDITÁVEIS

Permitir criar, escolher e modificar cenários comparativos independentes.

Cada cenário deve preservar:

- Identificação.
- Condições financeiras.
- Condições operacionais.
- Outras premissas editáveis.
- Data.
- Resultados calculados.

As alterações ao Comparativo não podem modificar silenciosamente a Base.

As alterações à Base não podem modificar silenciosamente o Comparativo.

Todos os Project Impact Summaries devem mostrar:

**Base Atual | Cenário Comparativo | Delta**

---

# 10. DASH 04 — FILTROS BANCÁRIOS PREDEFINIDOS

Os filtros da Base e do Comparativo devem ter opções idênticas.

## TAN

**2% a 8%, em incrementos de 0,5 pontos percentuais.**

Opções:

2%; 2,5%; 3%; 3,5%; 4%; 4,5%; 5%; 5,5%; 6%; 6,5%; 7%; 7,5%; 8%; Custom.

O limite inferior de 2% não é a taxa-base inicial.

Preservar a TAN atualmente guardada.

## Financiamento

20 M€; 25 M€; 30 M€; 35 M€; 40 M€; Custom.

## Prazo

5; 10; 12; 15; 20 anos; Custom.

## Carência

0; 12; 24; 30; 36; 42 meses; Custom.

## Comissão de abertura

0%; 0,5%; 1%; 1,5%; 2%; Custom.

## Comissão de compromisso

0%; 0,25%; 0,5%; 0,75%; 1%; Custom.

## Amortização antecipada

0%; 0,5%; 1%; 1,5%; 2%; Custom.

## Estratégia de amortização

- Contratual.
- Cash Sweep.
- Recebimentos imobiliários.
- Custom.

Os filtros devem alimentar realmente:

- Juros.
- Serviço da dívida.
- Comissões.
- DSCR.
- Liquidez.
- Dividendos.
- VAL.
- TIR.
- MOIC.

---

# 11. DASH 07 — FILTROS HOTEL & MICE

Base e Comparativo devem possuir as mesmas opções predefinidas.

## ADR

120; 140; 160; 180; 200; 225; 250 €; Custom.

## Ocupação

45%; 50%; 55%; 60%; 65%; 70%; 75%; 78%; Custom.

## Crescimento ADR

0%; 1%; 2%; 3%; 4%; 5%; 6%; Custom.

## Utilização MICE

20%; 30%; 40%; 50%; 60%; 70%; 80%; Custom.

## Base Management Fee

2%; 2,5%; 3%; 3,5%; 4%; Custom.

## Incentive Fee

5%; 6%; 7%; 8%; 9%; 10%; Custom.

## Marketing, Distribution & Systems

1,5%; 2%; 2,5%; 3%; 3,5%; 4%; Custom.

Preservar bases de incidência distintas quando contratualmente aplicável.

## FF&E

2%; 3%; 4%; 5%; 6%; Custom.

## Perfis da Operadora

- Operator Proposed.
- KK Target.
- KK Optimized.
- Custom.

Os valores dos perfis devem poder ser ajustados individualmente.

Uma alteração manual deve ser identificada como personalização do perfil, sem sobrescrever silenciosamente a configuração original.

---

# 12. PROJECT IMPACT SUMMARY

Manter o Project Impact Summary em todos os dashboards.

Exibir:

- Future CAPEX.
- Peak Debt.
- Minimum Liquidity Headroom.
- Planned Dividends.
- Investor NPV.
- Investor IRR.

Incluir MOIC, Exit e outras métricas específicas quando relevante.

Sempre com:

**Base | Comparativo | Delta**

Não utilizar valores fixos ou caches desatualizadas.

Os resultados devem estar associados ao cenário efetivamente calculado.

---

# 13. DASH 10 — REPORTING E IMPRESSÃO

A Base de Referência guardada e ativa determina os resultados principais dos relatórios.

Implementar os quatro formatos aprovados.

## 13.1 One-Page Executive

Uma página A4 com síntese de investimento, financiamento, principais KPI, riscos e conclusões.

## 13.2 Bank Financing

Relatório dirigido a bancos contendo:

- Financiamento solicitado.
- Construção.
- CAPEX.
- Desembolsos trimestrais.
- IVA.
- Fontes e aplicações.
- Equity.
- Recebimentos imobiliários.
- Serviço da dívida.
- CFADS.
- DSCR.
- Liquidez.
- Garantias e covenants.
- Stress tests.
- Capacidade de reembolso.

## 13.3 Equity Investor

Relatório para investidores contendo:

- Investment Thesis.
- Estrutura acionista.
- Tickets.
- Captação.
- Dividendos.
- VAL.
- TIR.
- MOIC.
- Cash flows.
- Exit.
- Sensibilidades.
- Riscos.

## 13.4 Comprehensive Project Report

Relatório geral integrado, com todas as componentes financeiras, operacionais, societárias e comerciais relevantes.

## Comparações

Permitir:

- Base Only.
- Base vs Selected Scenario.
- Base vs Multiple Saved Scenarios.

A Base deve aparecer sempre como principal referência.

As comparações só devem ser impressas quando selecionadas.

## Formatação

- A4.
- Identidade institucional JS-CV.
- Gráficos adequados ao destinatário.
- Quadros financeiros legíveis.
- Paginação flexível.
- Sumário nos documentos extensos.
- Cenário e versão identificados.
- Ausência de sidebar nos relatórios impressos.

## Snapshot

Cada relatório emitido deve conservar:

- Versão da Base.
- Cenários comparativos.
- Pressupostos.
- Resultados.
- Data.
- Destinatário.
- Identificador do relatório.

Os relatórios anteriores não devem ser alterados por modificações posteriores da Base.

---

# 14. PRESERVAÇÃO DA ESTRUTURA ACIONISTA

Manter:

**Promotores**

JSCV: 40%, dois blocos Platinum.

JALS: 10%, cinco blocos Silver.

**Novos investidores**

1 Platinum: 20% — 10 M€.

2 Gold: 10% cada — 5 M€ cada.

5 Silver: 2% cada — 1 M€ cada.

Total captação prevista: 25 M€.

Não confundir direitos económicos existentes dos promotores com entradas de novo capital.

Não atribuir tickets fictícios aos promotores.

---

# 15. INTEGRIDADE DOS RESULTADOS E INDICADORES

Preservar a coerência entre os 180 meses do modelo e os indicadores apresentados.

Verificar:

- CAPEX e respetivos perímetros.
- Desembolsos de construção.
- IVA pago e recuperado.
- Vendas e recebimentos.
- Operação residencial.
- Hotel & MICE.
- Financiamento.
- Serviço da dívida.
- Liquidez.
- Reservas.
- Dividendos.
- VAL.
- TIR.
- MOIC.
- Exit.

A TAN não pode ser confundida com a taxa de desconto do investidor.

A alteração da taxa de desconto deve afetar o VAL, mas não alterar mecanicamente a TIR.

---

# 16. MÉTODO DE EXECUÇÃO — ENTREGA ÚNICA

Executar internamente pela sequência mais eficiente:

1. Inspeção da versão atual.
2. Correção das falhas financeiras confirmadas.
3. Reconciliação dos inputs Base e Comparativo.
4. Correção das hiperligações e navegação.
5. Reorganização visual conforme o gabarito.
6. Integração dos filtros e cenários.
7. Integração do reporting.
8. Verificação final.
9. Entrega consolidada.

Estas são etapas internas de execução, não pontos de aprovação do utilizador.

Não produzir múltiplos ficheiros provisórios.

Não solicitar autorização intermédia para aspetos já aprovados.

Não reformular a arquitetura.

---

# 17. REQUISITOS DE QUALIDADE

O trabalho apenas poderá ser apresentado como concluído quando:

- Os dez dashboards respeitarem o gabarito.
- A navegação estiver funcional.
- Os quadros detalhados abrirem na localização correta.
- A Base e o Comparativo forem independentes.
- Os filtros predefinidos e Custom funcionarem.
- Os resultados financeiros forem rastreáveis.
- Os erros identificados tiverem correção fundamentada.
- Os gráficos estiverem ligados aos cálculos.
- Os relatórios respeitarem a Base ativa.
- A estrutura financeira permanecer coerente.

Não declarar que o Excel foi validado no Microsoft Excel 365 Desktop se essa validação não tiver sido realizada.

As limitações reais devem ser identificadas com precisão, sem impedir a execução das restantes correções possíveis.

---

# 18. ENTREGA FINAL

Entregar um único pacote consolidado com:

1. Workbook KK-01 atualizado.
2. Dez dashboards com organização fiel ao gabarito.
3. Navegação funcional para quadros detalhados.
4. Erros financeiros corrigidos na origem, quando confirmados.
5. Filtros Base e Comparativo.
6. Cenários guardados e histórico de referência.
7. Project Impact Summary.
8. Central de Relatórios Executivos.
9. Quatro relatórios de referência em PDF.
10. Resumo técnico de correções e exceções.

O resumo técnico deve ser conciso e indicar apenas:

- O que foi corrigido.
- Quais as falhas efetivamente identificadas.
- Quais as validações realizadas.
- Quais as limitações ainda existentes.

---

# ORDEM FINAL DE EXECUÇÃO

O projeto KK-01 já possui arquitetura, identidade gráfica, organização financeira e funcionalidades aprovadas.

**Não é necessário voltar a conceber o sistema.**

É necessário executar fielmente as decisões existentes.

O documento `Markdown colado.md` constitui o gabarito visual obrigatório.

O motor financeiro existente constitui a base técnica, sujeita à correção das falhas comprovadas.

As decisões posteriores sobre Base editável, cenários comparativos, filtros e reporting são igualmente obrigatórias.

O resultado pretendido é um **Executive Financial Control Center profissional, coerente, funcional e financeiramente auditável**, adequado à gestão do projeto KK, à negociação bancária e à captação de investidores.

**Executar a revisão integral e apresentar uma única entrega consolidada, sem aprovações intermédias e sem alterações arbitrárias ao desenho aprovado.**
