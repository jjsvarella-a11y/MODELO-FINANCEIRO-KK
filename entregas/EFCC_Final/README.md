> Atualização V06.3-SYNC: Base Case passa a usar Base (185 EUR / 58%). Resultados e testes atuais: [Relatório de sincronização](Relatorio_Sincronizacao_Cenarios.md). As equivalências anteriores Base=Favorável estão substituídas.

# KK-01 V06 — Executive Financial Control Center

Revisão premium de apresentação, navegação e impressão, sobre a base financeira estável `4111b2c`. O ficheiro `KK01_V06_Executive_Master (2).xlsx` é mencionado no mandato, mas não foi fornecido como anexo acessível; foi usada a revisão estável do repositório que o utilizador indicou parecer ter resolvido o erro.

- [Abrir o modelo](KK01_V06_Executive_Master.xlsx) — 36.3 MB, 190 folhas: dez painéis executivos, auxiliares e motor atual. As 313 folhas V04 continuam fisicamente removidas; a V04 original permanece intacta em `../Arquivo/V04`.
- [Entrega completa](KK01_EFCC_Entrega_Candidata.zip).
- [Testes de todos os filtros](Matriz_Testes_Filtros.csv).
- [Matriz de navegação](Matriz_Navegacao.csv) — 617 ligações verificadas estruturalmente.
- [Alterações gráficas](Relatorio_Alteracoes_Graficas.md).
- [Integridade financeira](Relatorio_Integridade_Financeira.md).
- [Compatibilidade Microsoft Excel 365](Compatibilidade_Excel365.md).

## Utilização

Abra **EXEC V06**, que corresponde ao Dashboard 01. A sidebar contém os dez painéis numerados. Os nomes técnicos existentes foram conservados: por exemplo, **DASH 05 V06** é o painel funcional 07 — Hotel & MICE; **DASH 07 V06** é o painel funcional 09 — Equity Returns.

O **DASH 10 Centro** reúne idioma global, formato integral/resumo, cenário ativo/referência, data e identificador. A escolha PT/EN é única; os atalhos de idioma abrem essa célula para evitar cópias editáveis concorrentes. **DASH 10 V06** continua a ser a folha técnica do relatório integral, com dez páginas; **V06 Brief** contém duas páginas.

Nos dashboards, as células azuis dos controlos são entradas editáveis; os valores ativos são apresentados em branco e os Custom não aplicados ficam a cinzento. Os controlos Hotel permanecem no painel 07. O residencial tem condições de gestão, yield, ADR, ocupação inicial, OPEX e FF&E. O painel 09 concentra as três políticas de reservas, dividendos, taxa de desconto e múltiplo de saída. Os pressupostos avançados continuam consultáveis por links, sem exigir busca pelas folhas do motor.

Cada auxiliar acessível tem Início e Voltar para o painel responsável. Nas folhas técnicas ocultas há uma ligação discreta ao executivo quando a primeira célula estava livre. As folhas não foram renomeadas. Freeze Panes conserva o cabeçalho e as colunas da sidebar; ao deslocar verticalmente a folha, os comandos compactos do cabeçalho mantêm acesso ao menu e aos parâmetros.

## Imprimir

Nos painéis, **Imprimir / PDF** abre a composição documental **V06 PDF 01–09**, com regresso ao respetivo dashboard. Prima **Ctrl+P**, escolha folhas ativas e use a área configurada. As nove composições são A3 horizontal, uma página cada, sem sidebar nem filtros de trabalho.

No centro do DASH 10, abra o relatório integral ou o resumo. Cada um possui área própria A4 vertical, com quebras de página e margens. Os comandos de navegação destas folhas ficam fora da área impressa. Antes de exportar, recalcule e confirme o cenário. Hiperligações Excel não abrem por si só a caixa de impressão: o procedimento nativo Ctrl+P está explicitamente descrito em **V06 Imprimir**.

Os PDFs PT/EN foram gerados com LibreOffice; não são provas de impressão nativa Excel. Os gráficos são objetos nativos e usam saídas do motor existente. As projeções anuais e os rácios de apresentação não constituem outro motor financeiro.

## Integridade e limites

O recálculo independente dos 69 cenários comparou 1,738,800 valores mensais com a referência Python: 0 erros e 0 divergências; diferença absoluta máxima 6e-08. A bateria preservada de controlos apresenta 16/16 PASS. A verificação global abrange 50 controlos em 73 testes de escolhas/entradas, com resultados em `Matriz_Testes_Filtros.csv`.

Trinta entradas foram transferidas para células canónicas nos painéis; os endereços antigos passaram a referenciar essas entradas. Quatro coeficientes residenciais antes fixos foram expostos como parâmetros, conservando exatamente a calibração original: ADR 225 EUR, ocupação inicial 58,1%, OPEX 35% e FF&E 4%. Esta ligação altera 48.960 expressões dos 68 motores ativos; o caso de referência permanece congelado. Todas as outras expressões dos motores e todos os valores iniciais foram preservados. Não se afirma que nenhuma fórmula financeira mudou: as quatro exposições são a alteração técnica necessária aos filtros pedidos.

**Microsoft Excel 365: NOT TESTED nesta revisão.** Não existe Excel Desktop/Windows neste ambiente. O utilizador observou que o problema anterior parecia resolvido; isso não certifica a nova revisão. Execute `Excel365/Executar_Validacao.cmd` numa cópia Windows para registar abertura, ativar edição, recálculo, navegação, PDF, guardar e reabrir. Os cliques não acessíveis à API permanecem NOT TESTED até observação humana. Não se apresenta estado Recalculated com base apenas em cache ou data.

**Logótipo original:** as imagens da conversa não chegaram como JPG/PNG acessível. Foi mantida **V06 Identidade**, para inserir o original diretamente no livro, e reservados espaços nos cabeçalhos. O script Windows pode copiar essa imagem incorporada para os cabeçalhos, preservando proporções; essa rotina continua NOT TESTED. Não se redesenhou nem substituiu o logótipo oficial por uma aproximação.

SHA256 do master: `94d0e1c511588a9a045e54f9259204a8c19ad9f3cb648357588c5d98a9f1fb40`.

## Aderência ao mandato de navegação

A revisão corrige o Menu, acrescenta dez acessos numerados na linha 6 imobilizada de cada dashboard, inclui RevPAR nos oito cartões Hotel e diferencia os cinco intervenientes na decisão final. A sidebar mantém títulos completos; em deslocação vertical, o menu compacto superior conserva acesso direto aos dez dashboards.

`Restauro_Mandato.json` compara esta revisão com a publicada: fórmulas e caches dos 69 motores, bem como as 50 entradas e validações, são idênticos. Os 73 testes anteriores são evidência preservada; os ensaios executados nesta correção constam de `Testes_Restauro.json`. Consulte `Auditoria_Mandato.md`. A aceitação nativa Excel 365 e o logótipo original incorporado permanecem pendentes.
