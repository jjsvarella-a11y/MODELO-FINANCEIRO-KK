# Diagnóstico de V06 Motor!G5

A fórmula enviada pelo utilizador corresponde à fórmula anterior, após normalização do idioma e remoção dos oito operadores @. O @SE reduz o resultado a um escalar, enquanto SOMARPRODUTO recebe um vetor de 30 pesos: as dimensões incompatíveis explicam o #VALOR! nesta fórmula. As folhas dependentes podem propagar o erro do motor; isso não exclui outras causas independentes.

O master atual usa integração escalar SUM/MAX/MIN e calendários escalares por fração. A inspeção das 1.738.800 fórmulas dos 69 motores não encontrou SUMPRODUCT, @ nem SINGLE. Base Case: G5 esperado 0,01 e H5 esperado 564.000 EUR. Os seletores e as ligações financeiras foram preservados.

Use KK01_V06_Executive_Master.xlsx desta entrega (179 folhas, 36,1 MB). O pequeno Diagnostico_Formulas_Excel365.xlsx contém deliberadamente a fórmula antiga para comparação; não é o modelo de produção. Não remova globalmente @ em livros arbitrários.

Os testes independentes passaram; a abertura, ativação de edição, recálculo completo, guardar e reabrir da revisão atual no Microsoft Excel 365 ainda não foram executados neste ambiente. Não se declara eliminado todo o #VALOR! sem essa prova.
