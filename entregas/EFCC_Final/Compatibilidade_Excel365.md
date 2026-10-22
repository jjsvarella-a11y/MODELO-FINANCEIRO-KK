# Microsoft Excel 365 — estado de compatibilidade

Estado atual: **NOT TESTED**. A tentativa e a ausência de Excel/conector estão registadas em `Tentativa_Excel365.json`. Plataforma disponível: Linux, Python e LibreOffice. Não se declara compatibilidade nativa, ausência de reparação, cliques nativos ou sucesso de guardar/reabrir sem execução real no Excel.

A fórmula G5 anterior enviada pelo utilizador continha oito operadores @, incompatíveis com o vetor de 30 pesos de SOMARPRODUTO. A revisão usa fórmulas escalares. O utilizador indicou que o erro parecia resolvido na base; os testes atuais independentes não substituem o ensaio nativo desta apresentação e dos novos controlos.

No Windows com Excel instalado, extraia a entrega e execute `Excel365/Executar_Validacao.cmd`. O script usa uma cópia e regista edição/versão, ativação de edição quando suportada, recálculo completo, erros por célula, mutações, PDFs e guardar/reabrir. Tenta seguir as hiperligações da matriz; quando a API não permite seguir uma ligação de fórmula, mantém esse clique NOT TESTED para revisão humana. Inspecione também avisos de reparação, títulos, dropdowns, filtros, logótipo e paginação 10/2.

O logótipo continua dependente de inserção do JPG/PNG original em V06 Identidade. A cópia para cabeçalhos via COM é opcional, explícita e ainda não foi testada neste ambiente.
