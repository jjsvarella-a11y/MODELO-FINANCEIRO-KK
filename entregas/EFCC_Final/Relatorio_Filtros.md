> Atualização V06.3-SYNC: Base Case passa a usar Base (185 EUR / 58%). Resultados e testes atuais: [Relatório de sincronização](Relatorio_Sincronizacao_Cenarios.md). As equivalências anteriores Base=Favorável estão substituídas.

# Verificação dos filtros — todos os dashboards

Resultado: PASS. 50 controlos, 73 mutações/revisões de escolhas, com caches invalidados antes do recálculo. A matriz detalhada está em [Matriz_Testes_Filtros.csv](Matriz_Testes_Filtros.csv); os valores e indicadores alterados constam de Testes_Todos_Filtros.json.

- Dashboard 01: atalhos para as fontes canónicas, sem segunda cópia editável.
- Dashboard 02: atraso, sobrecusto e ETC adicional.
- Dashboard 03: IRPC, IVA, prazo de reembolso e isenção simulada.
- Dashboard 04: juros, carência, prazo e cash sweep.
- Dashboard 05: atraso de vendas.
- Dashboard 06: condições residenciais, yield, ADR, ocupação, OPEX e FF&E.
- Dashboard 07: cenários operacionais, condições Hotel, MICE, FF&E, crescimento ADR e prioridade.
- Dashboard 08: atraso e fração de capital recebido.
- Dashboard 09: três reservas, seis políticas de dividendos, desconto, múltiplo de saída, distribuição parcial e cobertura.
- Dashboard 10: idioma, cenário do relatório, integral/resumo, data e identificador.

Correção de leitura: os valores ativos estão ligados às saídas do motor e ao Registo Ativo. ADR/ocupação Custom, fees Custom e desconto Custom permanecem guardados e editáveis; ficam a cinzento quando um preset não os utiliza. Não são apresentados como se fossem parâmetros ativos.

Equivalências verificadas: Base Case conserva a calibração Hotel original Favorável; condições residenciais Target e Custom podem ter o mesmo total de fees com repartições distintas. Estes casos podem manter o VAL. A carência e o prazo podem não produzir alterações quando o cash sweep já liquidou a dívida; ambos foram também testados com dívida em aberto, em Testes_Carencia_Prazo.json.

A escolha de condições da operadora altera fees; ADR e ocupação pertencem ao cenário operacional separado. O ano de consulta e os controlos documentais alteram a apresentação, preservando os fluxos mensais do motor.

Os laboratórios privados congelam ramos não usados pelo teste; o master conserva todos os ramos dinâmicos. Microsoft Excel 365: NOT TESTED. Os testes são Python/LibreOffice, sem equivalência a abertura, edição e guardar/reabrir no Excel.
