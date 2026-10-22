# KK-01 — Executive Financial Control Center (candidato)

Revisão integrada do V06 existente; não é uma certificação Microsoft Excel 365.

## Ficheiros principais

- `KK01_V06_Executive_Master.xlsx`: dez dashboards dinâmicos, 69 cenários de 180 meses, matriz ADR × ocupação, fees, reservas, disponibilidade, relatório integral e brief.
- `PDF/PT` e `PDF/EN`: nove dashboards individuais, relatório DASH 10 com dez páginas A4 e brief com duas páginas A4.
- `Excel365/Executar_Validacao.cmd`: teste nativo para Windows com Excel instalado, numa cópia do livro. Gera JSON, erros por célula, ambiente, mutações, exportações e teste guardar/reabrir. Nunca foi executado neste ambiente.
- `Parametros_Ativos_EFCC.xlsx`: snapshot identificado dos parâmetros efetivos e seletores. O registo dinâmico permanece dentro do master, em `V06 Registo Ativo`.
- `Relatorio_Auditoria_EFCC.pdf`, `Guia_Utilizacao.pdf`, registo de erros, evidências JSON e código-fonte.

## Resultado financeiro de referência

TIR investidores: **7,70%**. MOIC: **2.6867×**. VAL a 12%: **-9.604.777,39 €**. Equity Value de saída: **103.625.929,31 €**.

O cenário Base Case reproduz a referência V06: ADR inicial 210 €, ocupação 64%, crescimento ADR 3%, FF&E 5%, fees Hotel 3%/3%/6%, fees residenciais 3%+5%, reserva geral flexível 2 M€ + HOTUSA 1 M€. Os números da imagem de formatação não foram usados como pressupostos.

A 12%, 0 de 30 casos ADR 150–250 € / ocupação 50–75% alcançam VAL não negativo. Não se extrapolou equilíbrio fora da grelha.

## Evidência e limites

- PASS: paridade de 517.500 valores mensais dos 25 cenários anteriores após remoção das fórmulas matriciais extensas.
- PASS: 1,738,800 comparações numéricas em 69 cenários; sem erros ou falhas de identidades/restrições nesse recálculo independente.
- 17/17 testes de alteração real de seletores passaram em LibreOffice. Os testes da taxa de desconto comparam os 30 VAL da matriz com os fluxos de caixa originais.
- OOXML: PASS, verificações estruturais descritas em `Auditoria_OOXML.json`. Não é validação pelo Open XML SDK nem prova de abertura sem reparações no Excel.
- Microsoft Excel 365: **NOT TESTED**. Não há Excel Desktop/Windows nem conector Microsoft Excel neste ambiente Linux. A fórmula G5 enviada pelo utilizador contém oito operadores @: o resultado escalar de @SE é incompatível com o vetor de 30 pesos de SOMARPRODUTO. A revisão usa fórmulas escalares nos 69 motores; a eliminação de todos os erros no Excel permanece por confirmar.
- Logótipo: área de inserção disponível em `V06 Identidade`, independente das fórmulas. O utilizador pode incorporar o JPG/PNG original no XLSX. A sincronização explícita de cabeçalhos pelo ensaio Windows ainda é NOT TESTED nativamente.
- Contratos, recebimentos reais, calendários e capacidade/dias/eventos/participantes MICE carecem de documentação. O índice MICE normaliza 50% ao modelo original; 1.500 pessoas não representa presença média.
- Dados documentais em falta aparecem como “Not documented” ou “Unknown”. Previsões não equivalem a vendas, recebimentos ou gastos comprovados.

## Utilização

1. Abra `EXEC V06`. A barra PT/EN abaixo do cabeçalho abre o seletor global `V06 Controlos!B5`.
2. Escolha pressupostos nos painéis dos dashboards. No Hotel, `Custom` ativa ADR/ocupação editáveis; `Base Case` conserva o cenário documentado. As células azuis são inputs e as restantes apresentam fórmulas.
3. Fees Hotel: escolha o perfil; em Custom, marketing/distribuição são inputs no painel e base/sistemas/incentivo abrem os controlos existentes. Os 4%/3%/1,5% dos perfis são agregados de receita, não três cobranças repetidas. Incentivo usa GOP após base e agregado. Prioridade do proprietário desativada por defeito.
4. A reserva geral é independente de HOTUSA, FF&E, impostos a pagar e outros compromissos. Reduzi-la não elimina os testes de dívida, lucros ou cobertura histórica 1,20×.
5. `V06 Disponibilidade` filtra as 94 frações e separa estado documental e previsão. `V06 Cobrancas` guarda recebimentos documentais; o motor continua identificado como previsão. Meses relativos não são datas contratuais comprovadas.
6. `V06 Hipoteses` contém cenário/data/ID do relatório e formato. `Reference Case` usa parâmetros e inventário congelados; `Active Scenario` acompanha o motor. O seletor Complete Report / Executive Brief abre a folha de dez / duas páginas.
7. Para imprimir, use o controlo Imprimir/PDF, conclua o recálculo, prima Ctrl+P e escolha folhas ativas / PDF. O botão conduz a esse fluxo suportado; não executa uma falsa exportação automática.
8. Para testar nativamente, extraia o pacote no Windows e execute `Excel365/Executar_Validacao.cmd`. Os resultados são guardados numa pasta nova e o original não é modificado. A aceitação visual, logótipo, eventuais avisos de reparação e paginação ainda exigem observação humana.

## Reprodutibilidade

O código estende `entregas/V06/src/core.py` e reutiliza a fonte V04 no repositório. As fórmulas do master são autónomas e não dependem de Python para funcionar no Excel. `src/REBUILD.md` descreve a geração e testes neste ambiente. Os checkpoints anteriores V06 e V06_DASH10 permanecem no repositório. O master atual contém apenas 179 folhas V06. As 313 folhas legadas foram fisicamente removidas e a V04 original ficou intacta em `entregas/Arquivo/V04`, separada do livro operacional. Os 115 campos originais do motor V06 conservam as posições e recebem apenas campos adicionais no fim.

Estado de entrega: **CANDIDATE — aceitação nativa pendente; área de inserção do logótipo disponível**.

## Logótipo no próprio modelo

Abra **V06 Identidade** pelo link **Inserir logótipo** dos cabeçalhos. Use Inserir → Imagens → Este dispositivo, coloque um único JPG/PNG original em B8:H20 e guarde o livro. A imagem fica incorporada no XLSX, sem afetar o motor financeiro. A sincronização dos cabeçalhos é uma ação explícita fornecida no teste Windows/Excel, numa cópia; ainda NOT TESTED nativamente. Consulte `Excel365/LOGOTIPO.md`.

## Erro ao ativar edição

O utilizador confirmou que os valores guardados aparecem antes de editar, mas os #VALOR! surgem depois de ativar edição e reabrir. O ensaio nativo tenta Vista Protegida → Ativar Edição, regista os valores antes/depois e segue com recálculo completo, guardar e reabrir. Se a API de Vista Protegida não funcionar, esse passo fica NOT TESTED. Ver valores antes de editar não comprova o cálculo no Excel.

## Livro operacional sem V04 incorporada

As 313 folhas legadas foram removidas fisicamente. O master passou de 61,5 MB / 492 folhas para 36,1 MB / 179 folhas, uma redução de 41,3%. Foram retiradas 1.485.103 fórmulas antigas e os respetivos gráficos/partes sem uso. Os 179 XML de folhas V06 foram preservados byte a byte; não foram encontradas dependências para as folhas retiradas. A V04 original está no arquivo separado `entregas/Arquivo/V04`, fora do master e do ZIP operacional. Os dez dashboards atuais e os controlos permanecem acessíveis; os motores auxiliares atuais ficam ocultos. Ver `Separacao_V04.json` e `Reducao_Ficheiro.json`.

A primeira célula de erro reportada pelo utilizador é `V06 Motor!G5`. A ausência de #VALOR! no seu Excel ainda precisa de confirmação; esta redução de volume não é uma correção nativa certificada.
