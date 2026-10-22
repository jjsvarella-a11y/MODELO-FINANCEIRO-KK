# KK-01 V06.3-SYNC — Sincronização e testes dos seletores

A correção atua na fonte única `V06 Controlos!B67`, agora ligada a `DASH 05 V06!H23`. Antes, B67 estava fixo em Favorável: Base Case apresentava 210 EUR / 64% e o Conservador mudava ADR/ocupação mas conservava a composição Favorável de F&B/MICE. O motor de 180 meses já tinha dependências deste parâmetro; as expressões financeiras foram preservadas. A cadeia verificada inclui V06 Motor!T5:U184 (ADR/ocupação), V5:V184 (quartos), W5:W184 (MICE), X5:X184 (F&B), AA5:AA184 (GOP), DR5:DR184 (fees), AH5:AH184 (EBITDA), AI5:AI184 (FF&E), BX5:BX184 (CFADS), DE5:DE184 (dividendos), V06 F01 e V06 Cenarios!M5. Os resultados aparecem no DASH 05 V06 (funcional 07), DASH 07 V06 (funcional 09) e DASH 10 V06.

Base Case utiliza agora o perfil Base (185 EUR / 58%); Conservative utiliza Conservador (160 EUR / 52%). Favorable e Upside mantêm o perfil Favorável. Custom usa os valores próprios de ADR/ocupação e conserva o perfil Favorável para a composição de receitas; utilização MICE permanece um controlo independente. O teste Custom usa 200 EUR / 65%; o valor guardado no modelo pode ser diferente.

Esta correção altera legitimamente os resultados iniciais. A referência histórica permanece guardada; a comparação completa usa `Referencia_Python_V06_SYNC.json.gz`, executando o mesmo core existente com o perfil corrigido e as definições originais das variantes. A referência congelada Reference Case permanece histórica e não acompanha o cenário ativo por definição.

## Resultados dos cenários

| Cenário | ADR inicial | Ocupação inicial | EBITDA ano operacional 5 | VAL novos investidores |
|---|---:|---:|---:|---:|
| Base Case | 185 EUR | 58% | 4,093,567.40 EUR | -12,742,848.24 EUR |
| Conservative | 160 EUR | 52% | 2,929,030.42 EUR | -15,676,780.22 EUR |
| Favorable | 210 EUR | 64% | 5,560,671.09 EUR | -9,604,777.39 EUR |
| Upside Stress | 250 EUR | 75% | 7,262,659.09 EUR | -6,441,596.43 EUR |
| Custom | 200 EUR | 65% | 5,367,605.04 EUR | -10,203,553.66 EUR |

As receitas, EBITDA e encargos anuais referem-se ao quinto ano operacional selecionado; ADR e ocupação acima referem-se à abertura, antes do crescimento e ramp-up. VAL, TIR, MOIC, dividendos e saída usam o horizonte financeiro respetivo. A comparação completa antes/depois consta de `Comparacao_Cenarios.csv`. CAPEX, número físico de quartos e capital inicial podem permanecer iguais porque o cenário operacional não altera estes pressupostos. Taxas contratuais da operadora permanecem iguais se o contrato não mudar; os encargos em euros acompanham as receitas. Favorable/Upside/Custom conservam os resultados anteriores quando os seus parâmetros não mudam. Custom guardado em 210 EUR / 64% pode coincidir com Favorable; o teste específico usa 200 EUR / 65% para demonstrar a independência. A margem percentual conservadora pode superar a Base pela composição de receitas, embora o EBITDA absoluto seja menor.

Cada um dos cinco cenários foi comparado em 25 200 valores mensais, 181 fluxos de investidores, 34 ligações de dashboards/relatório/PDF, 585 valores de projeções gráficas e 999 pontos de cache dos gráficos. Foram verificados os quadros inferiores, além do painel superior. Zero erros de célula nos ficheiros recalculados. A sequência Base -> Conservador -> Favorável -> Base e uma gravação/reabertura adicional constam de `Sincronizacao_Cenarios.json`, com diferenças e igualdade dos KPI registadas.

## Restantes dashboards

Foram testadas 73 alterações, cobrindo 50 entradas canónicas e as opções dos seletores. Cada teste regista a escolha, alterações no motor, KPI dos nove dashboards e resumo financeiro. Opções de idioma e apresentação devem preservar os cálculos. No residencial, KK Target (6% gestão + 2% distribuição) e Custom inicial (3% + 5%) totalizam ambos 8%: os componentes mudam, mas fees totais e VAL podem coincidir legitimamente. No Base corrigido, Auto e Minimum têm o mesmo objetivo anual de 3 milhões EUR nos 180 meses; a igualdade dos fluxos e do VAL foi confirmada, sem alterar o motor para forçar diferenças artificiais. Carência/prazo foram também testados com dívida em aberto; ausência de efeito quando a dívida já foi liquidada não implica desligação. O relatório Reference Case mantém intencionalmente a referência congelada. Resultados: `Matriz_Testes_Filtros.csv` e `Testes_Todos_Filtros.json`. Conferência adicional: 25,914 ligações de quadros inferiores e projeções gráficas nos 73 testes, em `Testes_Correspondencias_Dashboards.json`.

Comparação financeira completa: 69 variantes, 1,738,800 valores numéricos; estado PASS. Não foram introduzidas fórmulas IFERROR para esconder erros. Entradas Custom inativas ficam visualmente ocultas por formato, sem afetar os dados ou ocultar erros financeiros.

## Navegação, impressão e limites

O painel Active Scenario em EXEC V06 consulta os controlos canónicos e dá acesso direto aos mesmos. Os dez dashboards têm menu global e as auxiliares têm regresso. As ligações foram verificadas estruturalmente; os PDF PT/EN são exportações LibreOffice, com dez páginas A4 no relatório integral e duas no resumo. Os nove dashboards individuais têm composição A3 horizontal, segundo a arquitetura aprovada.

Microsoft Excel 365 Desktop: **NOT TESTED**, pois não está instalado neste ambiente Linux. Não é possível certificar abertura sem reparação ou comportamento nativo das hiperligações no Excel. O procedimento Windows atualizado está em `Excel365/Executar_Validacao.cmd`; executa em cópia, testa alteração dos cenários, recálculo, erros, gravação/reabertura e exportação. A auditoria OOXML verifica estrutura, referências, tipos e ordem de células, mas não substitui esse teste nativo.

O logótipo original não foi redesenhado: permanece disponível a área `V06 Identidade!B8:H20` para inserir o JPG/PNG oficial. As provas antigas que descrevem Base=Favorável são históricas e estão substituídas por este relatório e pelos testes SYNC.
