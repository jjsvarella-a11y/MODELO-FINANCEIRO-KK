# KK-01 V06 — DASH 10

Esta revisão acrescenta **10 — Relatório Executivo Consolidado de Viabilidade e Investimento / Consolidated Investment & Feasibility Report** ao modelo V06 existente. O motor financeiro foi preservado. As dez secções foram propostas para cobrir os temas da mensagem, cuja lista terminou durante a secção 02.

## Ficheiros

- KK01_V06_Executive_Master.xlsx: revisão autónoma, com o DASH 10 depois dos nove dashboards existentes.
- KK01_DASH10_Relatorio_PT.pdf e KK01_DASH10_Relatorio_EN.pdf: exportações integrais de dez páginas A4 cada, para o cenário de referência ativo na entrega.
- Fontes_DASH10.json: 83 referências financeiras, fontes e mapa de secções.
- Testes_DASH10.json: testes de ligações, idioma e alteração real do juro.
- KK01_V06_Executive_Master.preservation.json: comparação de 2.113.521 registos existentes, sem alterações financeiras.
- Validacao_Entrega.json: verificações estruturais, funcionais e de paginação.

## Utilização

1. Nos dashboards existentes, selecionar o novo acesso **10 — Relatório de Viabilidade** na sidebar.
2. Alterar o idioma no seletor global `V06 Controlos!B5`. O relatório acompanha PT/EN.
3. Alterar os pressupostos nos controlos existentes. O relatório lê os mesmos resultados dos dashboards; não existem inputs financeiros manuais no DASH 10.
4. Para exportar no Excel, selecionar apenas a folha `DASH 10 V06`, imprimir/exportar todas as suas páginas e respeitar a área de impressão A4. Não selecionar o livro inteiro, que contém folhas técnicas e dashboards com outras áreas de impressão.
5. Os PDFs entregues são snapshots. Após mudar um pressuposto, recalcular e exportar novamente para obter um PDF atualizado.

As células financeiras do relatório são referências aos mapas existentes. As únicas agregações novas são somas de apresentação de colunas do motor existente, sem novo cálculo de dívida, imposto, dividendo ou retorno.

## Preservação e limites

A base V06 anterior continua inalterada em `../V06/`. As únicas alterações nos nove dashboards existentes são os novos acessos de navegação B43. Nenhuma fórmula financeira, cenário, valor de input ou nome de folha existente foi substituído.

O trabalho desta revisão cobre o pedido expresso DASH 10. O anexo inclui um mandato mais amplo de novos simuladores, otimização, logótipo e revisão dos nove dashboards; esse âmbito não foi integrado nesta revisão de reporting. Mantêm-se as limitações financeiras já documentadas na V06. Não foram inventados calendários de eventos MICE nem saldos de fundos FF&E sem razão autónoma.

Microsoft Excel 365 Desktop: **NOT TESTED**. Os testes de recálculo e PDF foram feitos adicionalmente em LibreOffice; não certificam compatibilidade nativa Microsoft. O pedido de erros #VALUE! no anexo não foi dado como resolvido: não foi recebido um livro com esses erros reproduzíveis nesta revisão.

## Reconstrução

Dependências Python: xlsxwriter, openpyxl, lxml e PyMuPDF; LibreOffice para testes suplementares e exportação. Definir `KK01_REPORT_WORKDIR` para uma pasta temporária gravável ou usar o diretório de trabalho documentado nos scripts.

1. Executar `python src/build_report.py /caminho/para/V06_anterior.xlsx`.
2. Executar `python src/extend.py /caminho/para/V06_anterior.xlsx /workspace/kk01_dash10/DASH10_modulo.xlsx KK01_V06_Executive_Master.xlsx`.
3. Os testes de recálculo usam uma cópia reduzida dos módulos V06 existentes, com exatamente o mesmo motor e dados, para evitar recalcular folhas históricas que não foram alteradas. O relatório no livro completo é comparado com os mesmos vínculos e fontes.
4. `src/recalculate.py` invalida caches numa cópia privada antes do recálculo suplementar.
5. `src/export_pdf.py` exporta todo o relatório, sem cortar páginas. Usa uma cópia privada, retira áreas de impressão das outras folhas e remove a imobilização apenas nessa cópia para evitar artefactos de impressão do LibreOffice. O livro entregue mantém a imobilização para navegação.
6. `src/validate_delivery.py` verifica o artefacto completo e os PDFs entregues.

A fonte anterior exata está no commit GitHub `37ef42a66543aac967cd31ddd6f35d36bdc8fd43`, em `entregas/V06/KK01_V06_Executive_Master.xlsx`; o seu SHA-256 é registado no relatório de preservação. Não usar a revisão já estendida como base para inserir novamente uma folha de mesmo nome.
