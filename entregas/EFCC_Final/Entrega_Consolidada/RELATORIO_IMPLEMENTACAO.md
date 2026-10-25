# Relatório de implementação e exceções — V07

A revisão implementa uma Base global editável, Comparativo independente, os dois painéis de financiamento e Hotel & MICE, comparação financeira e emissão controlada dos quatro formatos. A origem é a R6.1, commit `f5409903c326fb65fa353ce7b011f9b3f51c732f`. R4, R6.1 e os relatórios anteriores permanecem preservados.

## Configuração entregue

- **Base:** TAN 6%, carência de capital 42 meses, amortização 120 meses após a carência, ADR 185 €, ocupação inicial 58%.
- **Comparativo demonstrativo:** TAN 5%, ADR 200 €, restantes condições inicialmente iguais. A configuração comparativa anterior está conservada em `CMP-0001`.
- Os 2% não são uma nova Base: são o limite inferior das opções, que avançam de 0,5 pontos percentuais até 8%.
- Cada contexto tem a sua própria configuração completa. O seletor do Dashboard 01 escolhe o contexto dos resultados operacionais. O Project Impact Summary compara sempre Base atual com Comparativo.
- Os controlos antigos que constituíam segundas entradas foram convertidos em acessos à entrada efetiva. Os valores Custom e perfis estão nas folhas de pressupostos, com caminhos de regresso aos painéis 04 e 07.

## Resultados da configuração entregue

| Indicador | Base atual | Comparativo |
|---|---:|---:|
| CAPEX futuro | 63,168,000.00 € | 63,168,000.00 € |
| Pico de dívida | 17,528,049.06 € | 17,449,911.76 € |
| Dividendos previstos | 26,960,052.94 € | 28,571,060.31 € |
| VAL novos investidores | -12,742,848.24 € | -11,832,458.30 € |
| TIR novos investidores | 5.7646% | 6.3752% |
| MOIC | 2.1174x | 2.2851x |
| Equity na saída | 80,709,171.59 € | 87,589,642.30 € |
| Custo de financiamento | 1,908,200.04 € | 1,773,371.82 € |

Base ID: `BASE-20261025T074219179008Z`. Comparativo ID: `CMP-20261025T074221097287Z`.

O VAL permanece negativo nos dois casos. Estes resultados não constituem aprovação bancária ou garantia de rendibilidade. A taxa de desconto dos investidores é 12%; a TAN é 6% na Base; o custo efetivo anual de financiamento calculado para a Base é 19.9377%. As três medidas não são intercambiáveis. O DSCR apresentado abrange meses com serviço da dívida, incluindo construção; valores negativos nesta fase exigem análise das fontes de reembolso e das condições bancárias, sem presumir dispensa contratual de covenants.

## Alterações técnicas relevantes

O motor de 180 meses não foi reconstruído. Dois contextos utilizam as mesmas expressões mensais e dependências escalares da R6.1, para calcular Base e Comparativo em simultâneo. As expressões de ADR e ocupação passam a consultar os parâmetros efetivos; mantêm crescimento, ramp-up e teto do motor. A parametrização de comissões continua a alimentar as respetivas bases de incidência existentes. Não foram introduzidas aproximações do VAL nem IFERROR para ocultar problemas novos.

O ano de leitura do painel Hotel foi transferido de N23 para N41, com atualização das referências locais e externas; N23 é agora o ADR do Comparativo. A comparação dos investidores selecionados usa fluxos atuais de cada contexto, em vez de valores congelados R4. As mensagens de restrição usam o contexto Comparativo, mesmo quando os gráficos mostram Base.

O ficheiro tem 238 folhas e 66.96 MB. A adição dos contextos de comparação aumenta o tamanho face à R6.1; não se afirma uma redução de tamanho nesta revisão. As folhas V04 anteriormente removidas não foram reintroduzidas.

## Ensaios executados

| Ensaio | Resultado |
|---|---|
| Reconciliação de 79 380 células numéricas entre R6.1, motor ativo, Base e Comparativo iniciais | PASS; diferença máxima zero |
| Dez casos de TAN, ADR, Custom, perfis, desconto, contexto ativo e retorno à Base | PASS |
| Recálculo de todas as folhas e sensibilidades na configuração de entrega | PASS; zero erros de fórmula |
| Stress conjunto TAN/ocupação/MICE/fees/FF&E/desconto, com matriz completa ativa | PASS; VAL da célula ativa = −9 491 808,68905679 €, igual ao motor |
| Saldos de caixa, dívida, FF&E e inventário dos 180 meses, por contexto | PASS; sem fluxos irreconciliados |
| Alteração apenas da taxa de desconto | VAL alterado; TIR e dividendos preservados |
| Idioma EN, cabeçalhos, navegação e pressupostos novos | PASS |
| Guardar/restaurar Base, conservar versão anterior e bloquear emissão por alterações não guardadas | PASS |
| Troca para Custom com TAN efetiva inalterada em 6% | Detetada como não guardada; escolha original e valor manual são versionados |
| Restrição comparativa com linha de 10 M€ e Base sem alteração | Folga comparativa −6 789 077,00787202 €; mensagem correta de reserva não coberta |
| 1299 destinos de navegação | PASS estrutural; cliques nativos não executados |
| 96 partes gráficas originais | Preservadas byte a byte, incluindo os gráficos ajustados manualmente |
| Open XML SDK, alvo Microsoft 365 | Zero erros estruturais |
| Nove PDFs A4, texto dentro das páginas, registo e hashes | PASS; corpo/tabelas ≥9 pt, rodapé 7 pt |
| Reprodução do relatório integral a partir do JSON congelado, sem ler o XLSX atual | PASS; 15 páginas com texto e valores idênticos |

Os PDFs incluem quatro formatos em PT/EN e um exemplo bancário Base vs Comparativo. Os relatórios bancário, de investidores e integral não têm limite fixo de páginas. As duas sínteses executivas têm exatamente uma página. Cada JSON conserva resultados, fluxos mensais e tabelas de orçamento, 94 unidades, matriz real e break-even, para que alterações futuras do livro não alterem o documento emitido.

Os testes foram executados em Linux, com recálculo independente LibreOffice, validação Open XML SDK e geração PDF ReportLab. Os JSON e a matriz CSV identificam os métodos e os limites; não se apresentam testes de versões anteriores como execução nativa desta revisão.

## Limites explícitos

1. **Microsoft Excel 365 Desktop indisponível:** não foi possível executar abertura, edição, recálculo, gravação, reabertura e PDF nativos. Zero erros de esquema e de recálculo independente não equivalem a essa aceitação. O auxiliar PowerShell fornecido usa Excel Desktop no computador do utilizador e está identificado como não executado neste ambiente.
2. **Versionamento sem VBA:** Ctrl+S guarda o livro, mas não substitui a operação de guardar a versão financeira. O auxiliar nativo ou o comando Python cria o histórico, verifica o recálculo e confirma os valores guardados. O emissor bloqueia a Base não guardada.
3. **Emissão automatizada:** os quatro formatos PDF são gerados pelo emissor incluído no repositório. Hiperligações do Excel dão acesso à configuração; não executam programas nem abrem automaticamente a impressão. As composições individuais e a impressão nativa existentes foram preservadas.
4. **Logótipo:** o binário original não consta da origem acessível. `V06 Identidade` permanece disponível para inserir a arte oficial; não foi inventada ou redesenhada.
5. **Documentação financeira:** condições sem evidência mantêm a natureza de hipótese; capital, vendas e recebimentos previstos não são tratados como documentos contratuais. Não foi identificada área bruta certificada exclusiva do Aparthotel; 47 000 m² permanece proxy global, não uma medição certificada dessa componente.

Consultar `LEIA-ME.md` para utilização e comandos, `Testes_Financeiros.json`, `Teste_Integral_Matriz.json`, os ensaios de histórico/Custom/restrições/reprodução, `Testes_PDF.json`, `Auditoria_Estrutural.json` e `Matriz_Navegacao.csv` para evidência.
