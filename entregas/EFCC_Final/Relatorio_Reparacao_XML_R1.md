# KK-01 V06.3-SYNC-R1 — Correção do aviso de reparação

O utilizador apresentou uma captura do Microsoft Excel que confirma reparação na revisão anterior de três partes XML: sheet492.xml, sheet323.xml e sheet328.xml. A auditoria anterior verificava XML bem formado, mas não verificava a cardinalidade obrigatória dos filhos destes elementos. Essa lacuna foi corrigida.

Foram encontrados exatamente três elementos vazios nas mesmas partes: hyperlinks em V06 Identidade (sheet492.xml), hyperlinks em V06 Controlos (sheet323.xml) e dataValidations em V06 Hipoteses (sheet328.xml). Nestes contentores SpreadsheetML, o elemento filho é obrigatório quando o contentor existe. As hiperligações de navegação já são fórmulas HYPERLINK; os contentores vazios eram resíduos da migração da navegação e das validações.

Foram removidos apenas esses três contentores. Não foram removidas hiperligações funcionais nem validações com opções. A correção foi integrada no gerador para impedir a sua recriação.

Verificações: 312 partes ZIP permanecem idênticas byte a byte; nas três partes modificadas, os 1.002 registos de células permanecem idênticos. A comparação integral das árvores XML confirma que as únicas diferenças são as três remoções. Todas as fórmulas, valores guardados, formatos, imagens, gráficos, definições de impressão e resultados dos testes de cenários foram preservados. A nova auditoria verifica os contentores vazios em todas as 190 folhas. A navegação foi novamente verificada estruturalmente (617 ligações).

Os resultados financeiros e os 73 testes de filtros anteriores são mantidos por preservação exata dos dados; não foram repetidos nesta correção estritamente estrutural. Os PDFs permanecem iguais porque a apresentação e os valores não mudaram. A evidência detalhada está em Correcao_Reparacao_XML.json e Auditoria_OOXML.json.

Microsoft Excel 365: a versão anterior teve reparação confirmada pelo utilizador. A R1 ainda não foi reaberta no Excel Desktop neste ambiente, onde esse programa não está disponível. A eliminação dos três defeitos foi verificada; isso não constitui certificação nativa de ausência de outros avisos. O procedimento Windows continua disponível em Excel365/Executar_Validacao.cmd.

Descarregar novamente KK01_V06_Executive_Master.xlsx do diretório habitual. Não continuar sobre a cópia que o Excel já reparou. Identificador da revisão: V06.3-SYNC-R1. SHA-256 do Excel: 4ef37eca1fe8d2bd249900b4e6e74c407d64b7f69dc978f48b656a80b19ae6c0.
