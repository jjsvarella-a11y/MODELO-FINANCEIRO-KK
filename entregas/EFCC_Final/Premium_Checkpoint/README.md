# KK-01 — PREMIUM R2 CHECKPOINT

Esta é a versão de trabalho mais recente. O Excel R1 na pasta superior permanece como referência histórica intacta.

- [Excel editável](KK01_Executive_Financial_Control_Center_CHECKPOINT.xlsx)
- [Testes e limites](Relatorio_Testes.pdf) e [matriz de qualidade](Qualidade.csv)
- [Relatório consolidado — 10 páginas A4](PDF/PT/DASH_10_PT.pdf)
- [Resumo — 2 páginas A4](PDF/PT/Executive_Brief_PT.pdf)
- [Dashboards individuais PT](PDF/PT/) e [EN](PDF/EN/)
- [Matriz de navegação](Matriz_Navegacao.csv)
- [Validação nativa no Windows](Excel365/Executar_Validacao.cmd)

A matriz de sensibilidade está no DASH 07 funcional (`DASH 05 V06`), abaixo das tabelas, e na página 9 do DASH 10. Usa 20 VAL reais do motor de 180 meses. O limiar de proximidade de zero é editável em W116. A seleção atual é marcada pela combinação mais próxima; o VAL exato aparece no bloco de equilíbrio.

74 alterações de filtros e 86.875 correspondências passaram; 12 casos específicos da matriz passaram. No Base, todas as combinações são negativas. Não são atribuídas cores verdes artificiais por posição relativa.

**Estado: CHECKPOINT.** Excel 365 Desktop não está disponível neste ambiente. Abertura sem reparação, recálculo/gravação/reabertura e interação nativa ainda não foram certificados. O ficheiro original do logótipo continua indisponível; utilizar a área de inserção em `V06 Identidade`. O motor permanece com fórmulas ativas no Excel entregue.

Os PDFs foram gerados com LibreOffice. Os dashboards individuais mantêm a folha de impressão A3 existente; o relatório integral é A4 (10 páginas) e o resumo A4 (2 páginas). Para imprimir sem VBA, seguir o acesso Imprimir/Exportar PDF para a folha dedicada e usar Ctrl+P; uma hiperligação XLSX não abre por si só a caixa de impressão.

`Recuperacao.zip` contém a base R1, a nova versão, o código e as provas. `SHA256SUMS.txt` identifica os artefactos entregues.
