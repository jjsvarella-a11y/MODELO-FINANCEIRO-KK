# KK-01 V08 — correções e verificação

25/10/2026. Fonte: workbook V07 publicado no GitHub, SHA-256 `0c937badf07a04de702e0ffbd1047721f71cbf905c8a21d9504c66e50659f638`. O commit posterior `9d361a7` apenas acrescentou um ficheiro vazio denominado Recebidos. A cópia alterada no Excel pelo utilizador não foi obtida pelo OneDrive e **não está incorporada nesta entrega**.

## Correções

- Conversões numéricas tipadas com NUMBERVALUE e separadores explícitos; novos presets bancários/hoteleiros são números nativos. Categorias permanecem texto. Seleções já guardadas mantêm a sua representação histórica e o mesmo valor económico. Nenhuma fórmula VALUE permanece nos dois gestores de inputs.
- Corrigido o tratamento de Custom nos cenários de preços, residencial e amortização: anteriormente devolvia Base/Existing da coluna manual. Perfis operacionais e da operadora ficam acessíveis diretamente no DASH 07, com overrides independentes.
- Corrigidas referências residuais a R4/contexto ativo nos indicadores do investidor selecionado. Por exemplo, os proveitos de saída do Comparativo passam de 40,35 M€ para 43,79 M€, correspondendo ao seu próprio motor. Não se modificaram os fluxos financeiros para obter esse resultado.
- Filtros editáveis nos dashboards, Base/Comparativo independentes, painel Active Scenario, sidebar comum 01–10 e menu superior imobilizado. Removida seleção simultânea de folhas: apenas EXEC V06 abre selecionada.
- Project Impact Summary depois dos filtros; KPI e gráficos específicos; quatro gráficos bancários promovidos à área principal; matriz real ADR × ocupação ao centro do DASH 07. Preservados os agregados hoteleiros usados pelas tabelas quando a matriz ocupa a sua antiga posição.
- Reparadas coordenadas desatualizadas dos indicadores de abertura, incluindo DASH 05 V06!P96/P98, agora ligados aos parâmetros ativos em vez de células vazias H24/H26.
- Orçamento integral com Base atual, Comparativo, delta e ajustes. Ligações diretas para E190 e E230, escolha de capítulo e retorno. Detalhe adicional de IVA/desembolso bruto a partir de E300. Os trimestres agregam os motores de 180 meses, incluindo fiscalização, contingência e ETC no CAPEX.
- DASH 10 com os quatro formatos, configuração direta e acesso aos PDFs/snapshots. Preservadas as dependências antigas de idioma e metadados; o seletor global efetivo está em EXEC V06!W5.

## Verificações executadas

- Recálculo integral independente das 238 folhas; sem erros de fórmula armazenados nem referências #REF! explícitas. Plataforma: LibreOfficeDev 26.8; não Microsoft Excel.
- 53.280 comparações numéricas mensais entre versões, sem diferenças superiores a 0,005 €, e preservação de todos os pressupostos efetivos adotados. Fórmulas e inputs de 222 folhas não destinadas à intervenção mantidos.
- Testes de TAN Base/Comparativo, ADR/ocupação, fees, Custom, desconto, perfis, IRPC, OPEX residencial, preços Custom, filtros T1, promotores e idioma. Alterar o desconto não altera a TIR; alterar um cenário não substitui o outro.
- Conservador → Favorável → Base regressa aos resultados iniciais. No teste integral de sensibilidade, alterando TAN, ADR, ocupação, MICE, fees, FF&E e desconto, matriz e motor produzem o mesmo VAL: −8.330.657,11 €.
- 1.077 destinos de hiperligações internas verificados estruturalmente. A navegação por clique no Excel Desktop não foi executada.
- Open XML SDK para Microsoft 365: zero erros de schema. ZIP íntegro. Inspeção visual de dez páginas renderizadas; matriz com vinte valores legíveis, sem marcadores ###.
- Quatro relatórios A4: executivo 1 página, bancário 11, investidores 5 e integrado 15, respeitando a paginação flexível do último mandato. Letra mínima 9 pt e conteúdo dentro das margens. Hashes conferem com snapshots.

## Causas confirmadas e limites

A dependência regional de VALUE foi eliminada, mas **não se reproduziu nem confirmou a causa do erro regional no Excel 365 Desktop**. Os defeitos de Custom, as referências residuais de cenários e a seleção simultânea de folhas foram confirmados. Não se acrescentou IFERROR para esconder resultados financeiros nem se substituíram erros por zeros.

Microsoft Excel 365 Desktop não está disponível neste ambiente: abertura, recálculo, gravação/reabertura, cliques e ausência de avisos de reparação nessa aplicação continuam por validar. Schema e recálculo independente não constituem essa certificação.

A imagem oficial JS-CV não existe como ficheiro binário na fonte disponível. A área de inserção/identidade foi preservada; não se redesenhou o logótipo. A ABC certificada do Aparthotel continua por confirmar e não foi substituída pelos 47.000 m² globais.

A análise de equilíbrio utiliza a grelha técnica discreta, sem extrapolação: os rótulos indicam mínimo **testado**, não uma raiz contínua. Os eixos da sensibilidade são ADR/ocupação iniciais; os KPI operacionais respeitam o ano de leitura escolhido. Filtros de consulta não alteram, por si, os cash flows globais do projeto. Os gráficos distinguem distribuições do projeto, investidor selecionado e novos investidores globais.

Não se conclui automaticamente que o projeto é bancável: o VAL da Base permanece negativo e o DSCR deve ser interpretado no perímetro temporal do motor.
