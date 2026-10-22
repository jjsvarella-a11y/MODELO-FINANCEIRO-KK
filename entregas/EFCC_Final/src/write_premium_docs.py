"""Release documentation for premium presentation and native print architecture."""
from pathlib import Path
import json,hashlib,csv
from write_docs import pdf
from premium_ui import ROOT

def write():
 integrity=json.loads((ROOT/'Integridade_Premium.json').read_text());nav=json.loads((ROOT/'Validacao_Navegacao.json').read_text());fin=json.loads((ROOT/'Validacao_Financeira.json').read_text());tests=json.loads((ROOT/'Testes_Controlo_Premium.json').read_text());filters=json.loads((ROOT/'Testes_Todos_Filtros.json').read_text());master=ROOT/'KK01_V06_Executive_Master.xlsx';sha=hashlib.sha256(master.read_bytes()).hexdigest();size=master.stat().st_size/1e6
 readme=f'''# KK-01 V06 — Executive Financial Control Center

Revisão premium de apresentação, navegação e impressão, sobre a base financeira estável `4111b2c`. O ficheiro `KK01_V06_Executive_Master (2).xlsx` é mencionado no mandato, mas não foi fornecido como anexo acessível; foi usada a revisão estável do repositório que o utilizador indicou parecer ter resolvido o erro.

- [Abrir o modelo](KK01_V06_Executive_Master.xlsx) — {size:.1f} MB, 190 folhas: dez painéis executivos, auxiliares e motor atual. As 313 folhas V04 continuam fisicamente removidas; a V04 original permanece intacta em `../Arquivo/V04`.
- [Entrega completa](KK01_EFCC_Entrega_Candidata.zip).
- [Testes de todos os filtros](Matriz_Testes_Filtros.csv).
- [Matriz de navegação](Matriz_Navegacao.csv) — {nav['route_count']} ligações verificadas estruturalmente.
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

O recálculo independente dos 69 cenários comparou {fin['numeric_monthly_comparisons']:,} valores mensais com a referência Python: {len(fin['errors'])} erros e {len(fin['differences'])} divergências; diferença absoluta máxima {fin['maximum_absolute_numeric_difference']:.2g}. Os testes atuais de controlos são {sum(x['status']=='PASS' for x in tests['tests'])}/{len(tests['tests'])} PASS. A verificação global abrange {filters['inputs_covered']} controlos em {len(filters['tests'])} testes de escolhas/entradas, com resultados em `Matriz_Testes_Filtros.csv`.

Trinta entradas foram transferidas para células canónicas nos painéis; os endereços antigos passaram a referenciar essas entradas. Quatro coeficientes residenciais antes fixos foram expostos como parâmetros, conservando exatamente a calibração original: ADR 225 EUR, ocupação inicial 58,1%, OPEX 35% e FF&E 4%. Esta ligação altera 48.960 expressões dos 68 motores ativos; o caso de referência permanece congelado. Todas as outras expressões dos motores e todos os valores iniciais foram preservados. Não se afirma que nenhuma fórmula financeira mudou: as quatro exposições são a alteração técnica necessária aos filtros pedidos.

**Microsoft Excel 365: NOT TESTED nesta revisão.** Não existe Excel Desktop/Windows neste ambiente. O utilizador observou que o problema anterior parecia resolvido; isso não certifica a nova revisão. Execute `Excel365/Executar_Validacao.cmd` numa cópia Windows para registar abertura, ativar edição, recálculo, navegação, PDF, guardar e reabrir. Os cliques não acessíveis à API permanecem NOT TESTED até observação humana. Não se apresenta estado Recalculated com base apenas em cache ou data.

**Logótipo original:** as imagens da conversa não chegaram como JPG/PNG acessível. Foi mantida **V06 Identidade**, para inserir o original diretamente no livro, e reservados espaços nos cabeçalhos. O script Windows pode copiar essa imagem incorporada para os cabeçalhos, preservando proporções; essa rotina continua NOT TESTED. Não se redesenhou nem substituiu o logótipo oficial por uma aproximação.

SHA256 do master: `{sha}`.
'''
 (ROOT/'README.md').write_text(readme)
 graphic='''# Alterações gráficas e navegação

- Paleta institucional: azul #142337, grafite #252932, vermelho #C8101E, branco, fundo #F5F7FA, azul #1765A8 e cinza #687586; Aptos.
- Dez painéis com sidebar numerada, painel ativo a vermelho, cabeçalho institucional e acesso permanente a Início, parâmetros, menu e impressão.
- Auxiliares com título e cabeçalho compacto Início | Voltar ao Dashboard | Parâmetros | Ajuda; destino Back definido por responsabilidade funcional.
- Hotel: pressupostos próximos dos indicadores, fees personalizados diretamente acessíveis, indicadores anuais selecionados, receitas por componente, donut de mix e EBITDA/margem. Sem substituir os números reais pelos da referência visual.
- Leituras ativas de ADR, ocupação e fees Hotel, fees residenciais e desconto; valores Custom guardados a cinzento quando não aplicáveis. Base Case preserva a calibração Hotel de referência.
- Residencial: condições separadas do Hotel, yield, ADR, ocupação, OPEX e FF&E com fontes canónicas.
- Equity Returns: reservas, dividendos, desconto e múltiplo de saída no painel responsável.
- DASH 10 Centro: controlos, índice e ligação que acompanha a escolha integral/resumo. Relatório integral e resumo são folhas documentais separadas, sem duplicar o motor financeiro.
- Nove folhas V06 PDF contêm composições próprias A3 horizontal, sem sidebar. Relatório: dez páginas A4; resumo: duas páginas A4. Navegação fora das áreas impressas.
- Duas matrizes VAL (10% e 12%) calculadas a partir dos 30 vetores de caixa existentes; escala em M€, ponto neutro zero.
- Cartões continuam numéricos. Percentagens, euros e rácios têm formatos distintos; valores negativos permanecem visíveis.
- Nenhum nome técnico existente foi alterado. As folhas de apresentação acrescentadas não são motores financeiros antigos.

A inspeção e os PDFs usam LibreOffice. A validação visual e os cliques no Microsoft Excel 365 permanecem pendentes. A marca tipográfica não substitui o logótipo oficial; o original deverá ser inserido na área própria.
'''
 (ROOT/'Relatorio_Alteracoes_Graficas.md').write_text(graphic)
 financial=f'''# Integridade financeira

A base estável é identificada por SHA256 `{integrity['source_sha256']}`. O master atual tem SHA256 `{sha}`.

Preservação: {integrity['original_engine_formulas_preserved']:,} fórmulas anteriores dos motores permaneceram idênticas. As únicas alterações de expressões do motor são {integrity['residential_coefficient_exposures']:,} exposições de quatro coeficientes residenciais em 68 motores, com valores iniciais exatos. O motor de referência continua congelado. Todos os caches financeiros anteriores foram preservados na montagem e depois confrontados com recálculo independente.

O recálculo completo dos 69 cenários passou {fin['numeric_monthly_comparisons']:,} comparações, sem erros nem divergências acima da tolerância. Caixa, dívida, reservas, FF&E, inventário e restrições de dividendos continuam reconciliados. As entradas antigas apontam agora para 30 células canónicas de painel; não existem duas cópias editáveis concorrentes.

Os testes atuais de mutação constam de `Testes_Controlo_Premium.json`; a matriz de navegação consta de `Matriz_Navegacao.csv`. Os laboratórios privados congelam ramos independentes não testados; essa simplificação nunca é aplicada ao master entregue. Os 69 motores do master permanecem dinâmicos.

A referência Python completa e os testes históricos permanecem no pacote como evidência da base. O ensaio financeiro atual e os testes premium identificam o alcance efetivamente verificado. Não foi realizado teste nativo Microsoft Excel 365.
'''
 filter_report=f"""# Verificação dos filtros — todos os dashboards

Resultado: {filters['status']}. {filters['inputs_covered']} controlos, {len(filters['tests'])} mutações/revisões de escolhas, com caches invalidados antes do recálculo. A matriz detalhada está em [Matriz_Testes_Filtros.csv](Matriz_Testes_Filtros.csv); os valores e indicadores alterados constam de Testes_Todos_Filtros.json.

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
"""
 (ROOT/'Relatorio_Filtros.md').write_text(filter_report)
 (ROOT/'Relatorio_Integridade_Financeira.md').write_text(financial)
 native='''# Microsoft Excel 365 — estado de compatibilidade

Estado atual: **NOT TESTED**. A tentativa e a ausência de Excel/conector estão registadas em `Tentativa_Excel365.json`. Plataforma disponível: Linux, Python e LibreOffice. Não se declara compatibilidade nativa, ausência de reparação, cliques nativos ou sucesso de guardar/reabrir sem execução real no Excel.

A fórmula G5 anterior enviada pelo utilizador continha oito operadores @, incompatíveis com o vetor de 30 pesos de SOMARPRODUTO. A revisão usa fórmulas escalares. O utilizador indicou que o erro parecia resolvido na base; os testes atuais independentes não substituem o ensaio nativo desta apresentação e dos novos controlos.

No Windows com Excel instalado, extraia a entrega e execute `Excel365/Executar_Validacao.cmd`. O script usa uma cópia e regista edição/versão, ativação de edição quando suportada, recálculo completo, erros por célula, mutações, PDFs e guardar/reabrir. Tenta seguir as hiperligações da matriz; quando a API não permite seguir uma ligação de fórmula, mantém esse clique NOT TESTED para revisão humana. Inspecione também avisos de reparação, títulos, dropdowns, filtros, logótipo e paginação 10/2.

O logótipo continua dependente de inserção do JPG/PNG original em V06 Identidade. A cópia para cabeçalhos via COM é opcional, explícita e ainda não foi testada neste ambiente.
'''
 (ROOT/'Compatibilidade_Excel365.md').write_text(native)
 pdf(ROOT/'Relatorio_Auditoria_EFCC.pdf','KK-01 — Integridade e compatibilidade', [('Integridade financeira',[financial]),('Compatibilidade nativa',[native])])
 pdf(ROOT/'Guia_Utilizacao.pdf','KK-01 — Navegação e impressão',[('Utilização dos controlos',[readme.split('## Utilização')[1].split('## Imprimir')[0]]),('Impressão',[readme.split('## Imprimir')[1].split('## Integridade')[0]]),('Logótipo e validação',[native])])
 pdf(ROOT/'Relatorio_Alteracoes_Graficas.pdf','KK-01 — Revisão gráfica e navegação',[('Alterações',[graphic])])
 with (ROOT/'Aceitacao.csv').open('w',newline='',encoding='utf-8-sig') as f:
  wr=csv.writer(f,delimiter=';',lineterminator='\n');wr.writerow(['Teste','Estado','Plataforma']);wr.writerows([['Navegação estrutural',nav['status'],'OOXML'],['69 cenários e reconciliação',fin['status'],'Python + LibreOffice'],['Controlos premium',tests['status'],'LibreOffice'],['Todos os filtros e presets',filters['status'],'LibreOffice'],['Abertura, edição e recálculo Excel 365','NOT TESTED','Excel Desktop indisponível'],['Cliques nativos, impressão e guardar/reabrir','NOT TESTED','Exige Excel Windows'],['Logótipo original incorporado','PENDENTE','Inserção pelo utilizador em V06 Identidade']])
 print('Premium documentation generated',sha)
if __name__=='__main__':write()
