> Atualização V06.3-SYNC: Base Case passa a usar Base (185 EUR / 58%). Resultados e testes atuais: [Relatório de sincronização](Relatorio_Sincronizacao_Cenarios.md). As equivalências anteriores Base=Favorável estão substituídas.

# Aderência ao mandato — revisão de navegação

A revisão anterior não cumpria integralmente os mandatos: o Menu abria capítulos, o conjunto de cartões Hotel omitia RevPAR e a decisão final não distinguia os cinco intervenientes. Estes três pontos foram corrigidos. Não se afirma que houve uma regressão financeira: a comparação direta com o ficheiro publicado confirma fórmulas e valores idênticos nos 69 motores.

| Requisito | Resultado e alcance |
|---|---|
| Dashboard 01 inicial | PASS estrutural: EXEC V06 é a folha inicial selecionada. Abertura nativa Excel pendente. |
| Dez dashboards na sidebar | PASS estrutural; dez acessos adicionais na linha 6 imobilizada. |
| Menu funcional | PASS: abre a faixa de acesso aos dez dashboards, em vez dos capítulos do relatório. |
| Home e Back nos auxiliares | PASS estrutural; destinos definidos por responsabilidade e verificados na matriz. |
| Filtros nos dashboards | PASS: 50 entradas, valores, endereços e regras preservados; motor único. |
| PT/EN | Teste de idioma e conservação de resultados; números 01–10 são neutros. |
| Ausência de #VALOR! | PASS no recálculo independente e ensaios indicados em Testes_Restauro.json; não certifica Excel. |
| Hiperligações no Excel 365 | NOT TESTED: sem Excel Desktop ou conector neste ambiente. |
| Impressão sem navegação | PDFs usam composições próprias; faixas de navegação ficam fora das áreas documentais. |
| Reconciliação financeira | PASS: 1.738.800 comparações mensais; fórmulas e caches dos 69 motores idênticos à revisão publicada. |
| Hotel — oito indicadores | RevPAR reposto; GOP permanece na tabela operacional. |
| Relatório — cinco intervenientes | Promotor, investidores, banca, compradores e operadora com indicador e condição/decisão próprios. |
| Logótipo original | PENDENTE: área V06 Identidade preservada para inserir a imagem original; não foi redesenhado. |

## Imobilização

O cabeçalho e as colunas da sidebar estão imobilizados. Os títulos completos da sidebar acompanham a deslocação vertical da folha; a faixa compacta 01–10 na linha 6 permanece visível, mantendo todos os destinos acessíveis. Isto não é apresentado como uma sidebar inteira com deslocação independente do corpo, função que Freeze Panes não fornece.

## Evidência e aceitação

`Matriz_Navegacao.csv` indica origem, destino, tipo, resultado estrutural e teste de clique nativo. `Restauro_Mandato.json` compara diretamente o ficheiro publicado antes da correção com o novo. Os 73 testes dos filtros anteriores são preservados como evidência, apoiados pela identidade dos controlos e motores; não foram apresentados como 73 execuções novas. Os testes novos estão separados em `Testes_Restauro.json`.

A aceitação integral no Microsoft Excel 365 permanece aberta. Extraia a entrega num Windows com Excel 365 e execute `Excel365/Executar_Validacao.cmd` para obter os registos reais de abertura, edição, recálculo, cliques disponíveis à API e guardar/reabrir. A inspeção humana continua necessária para mensagens de reparação e apresentação nativa.
