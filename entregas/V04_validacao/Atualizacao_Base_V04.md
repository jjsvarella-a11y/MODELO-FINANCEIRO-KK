# Atualização — Base V04 recebida

A V04 Excel Fix recebida coincide byte a byte, pelo SHA-256, com a base identificada na auditoria anterior: `be8288ced2494ef31c2dbe1d31d77b8277a0753a1209a8df0a83c3a5d86dcdc0`.

## Verificações desta execução
- 313 folhas, das quais 249 ocultas.
- 1.485.103 elementos de fórmula; 347 nomes definidos; 17 gráficos.
- 886 validações e 553 hiperligações.
- Integridade CRC ZIP e leitura dos XML concluídas sem erro.
- Sem alvos internos ausentes ou IDs de relações duplicados nas verificações executadas.
- Nenhuma célula com tipo de erro em cache.

Estas verificações são estáticas: não houve recálculo financeiro ou teste de abertura em Microsoft Excel.

O ZIP V04 foi verificado e extraído numa pasta de referência separada. Contém código, documentação, resultados históricos e uma V04 diferente da Excel Fix. Essa V04 tem 1.469.149 elementos de fórmula; a diferença de contagem não demonstra, por si, perda de cálculos, pois representações de fórmulas partilhadas podem diferir. As versões não foram substituídas nem normalizadas.

## Correção anterior documentada
O script histórico patch_compat_v04.py altera o prefixo de MINIFS para _xlfn.MINIFS, controlos de desvios de retorno e uma expressão de ETC. Os resultados arquivados relatam comparações em LibreOffice; Verificacao_Reabertura_V04.json declara explicitamente excel_desktop_tested=false. Esses resultados não foram reexecutados nem apresentados como certificação Microsoft.

## Limite atual
A base V04 permite uma referência autenticada para comparações futuras. Não substitui a V05.1: a V05 alterou módulos fiscais, calendários, dividendos e saída. Não é seguro gerar uma V05.2 financeira a partir da V04 e presumir paridade com a V05.1 inacessível.

É necessário receber a V05.1 Excel Fix num ZIP inferior a 32 MiB ou em partes menores. O relatório de reparação error*.xml, se disponível, fornecerá evidência direta. Nenhum ficheiro financeiro foi modificado; a causa da reparação da V05.1 continua por determinar.
