# KK-01 V06 — Auditoria Financeira e Técnica

22/10/2026 · V06 CANDIDATE

## Fontes e preservação

SHA-256 da V04 Fixed: be8288ced2494ef31c2dbe1d31d77b8277a0753a1209a8df0a83c3a5d86dcdc0. Fontes preservadas read-only no diretório de trabalho. Registos de células V04 preservados; apenas a seleção visual da antiga página inicial foi desativada para evitar folhas agrupadas; as novas folhas são V06. Estilos, strings, registo de folhas e relações globais foram combinados com remapeamento de índices.

## Arquitetura

As folhas V04 preservadas são referência histórica e mantêm os seus consumidores próprios. Os nove dashboards novos consultam exclusivamente V06 Motor e os seus módulos V06. Não se devem misturar KPI históricos V04 com KPI V06.

94 apartamentos; 15 blocos; 8 tickets novos; 180 meses; 60 trimestres; 15 anos; 25 cenários completos. Os dados de referência e geradores V04 acompanham a entrega.

## Métodos financeiros

IVA: incorrido em construção e fiscalização; fração elegível e prazo editáveis; crédito inicial + elegível - reembolso = crédito final. Sem compensação automática entre entidades. IVA não recuperável é deduzido como encargo de estudo; classificação contabilística final requer validação.

IRPC: resultado da promoção por área, operação Hotel e residual KK, menos depreciação, juros e encargos; perdas fiscais acumuladas e anos fiscais do projeto. Taxa incide sobre lucro tributável, não receita ou EBITDA. Pagamento três meses após fecho anual, por hipótese. As isenções de exploração só excluem o resultado dos meses de operação elegíveis; não apagam retroativamente lucros de promoção anteriores à abertura. Benefícios de exploração são simulações; estado da Convenção nunca concede isenção.

Financiamento não renovável; a janela de draw termina M30 e não é prolongada automaticamente pelo atraso de obra. Juros sobre abertura + draw, comissões na primeira utilização, serviço obrigatório e cash sweep limitado por reservas/necessidades futuras. Último mês de draw protege imposto reconhecido pela proxy sem juros atuais, evitando circularidade.

Dividendos somente M54 a M174, onze datas; cobertura dos doze meses ANTERIORES e liquidez acumulada independentes. Em modo Proxy, lucros não significam aprovação societária. Modo Documented limita ao montante documental introduzido. Rateio por metas das categorias; saída proporcional 40/10/50 por hipótese.

Seguro obra: 0,30% + 0,05% sobre construção, aplicados uma vez. Seguro operacional: 40,1 M x 0,35% + receita x 0,15% + receita x 0,15% + 15 mil/ano. Frações já incluídas nos contratos/forecast evitam dupla contagem.

DSCR de Hotel é antes de alocação do IRPC consolidado, exibido separadamente; a cobertura de distribuição deduz o imposto consolidado e serviço obrigatório. Comparar o DSCR Hotel preservado com o contexto consolidado, não tratá-lo como covenant contratado.

M120 e M180 são alternativas exclusivas. EV usa EBITDA últimos 12 meses; deduz custos, dívida e IRPC pendente; soma caixa uma vez e protege HOTUSA. Custos históricos individuais dos promotores indisponíveis: TIR e MOIC individuais não inventados.

## Testes executados

V04: 34 testes Python e 9.750 comparações de referência previamente executados nesta sessão, todos aprovados.

V06: 27 verificações aprovadas; 0 falhas. 517500 comparações mensais entre Python e fórmulas recalculadas externamente após invalidar as caches.

Seletores: 82892 comparações em quatro mutações reais: idioma EN, IVA 18 meses, juro 8% e capital +6 meses. Resultado: PASS

Casos de TIR com raízes múltiplas: [{"scenario": "Construction +6M", "annual_roots": [0.1919988533072422, 283.586121931309], "selected_root": 0.1919988533072422, "criterion": "Root nearest Excel monthly IRR guess 1%; prefer NPV when roots are multiple"}]. A tabela escolhe a raiz mais próxima do guess mensal de 1% do Excel; não representa taxa única. Privilegiar o VPL.

Testes abrangem identidades de caixa/dívida, limite não renovável, reservas, datas/limites dos dividendos, rendas limitadas à caixa, rateio dos 15 blocos, taxas de retorno, cenários e erros armazenados. LibreOffice é somente evidência adicional de recálculo.

## Diferenças não ocultadas

CAPEX e IVA de referência conciliam. Resultados de caixa, dívida, IRPC, EBITDA e retornos diferem da V05 publicada, devido a inputs reconstruídos e métodos identificados. A ausência do modelo V05 de precisão integral impede certificar paridade exata. Não foram introduzidos valores fixos para forçar os benchmarks.
