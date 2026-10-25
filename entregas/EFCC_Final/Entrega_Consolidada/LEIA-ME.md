# KK-01 — Base de Referência editável e Comparativos independentes

O ficheiro desta pasta é a revisão consolidada V07. A Base inicial preserva **TAN 6%, carência de capital 42 meses, prazo de amortização 120 meses após a carência, ADR inicial 185 € e ocupação inicial 58%**. Os 2% são apenas o limite inferior das opções de TAN. O Comparativo selecionado na entrega é uma demonstração independente com TAN 5% e ADR 200 €; a configuração anterior 6% / 185 € permanece em `CMP-0001`.

## Utilização no Excel

1. Abrir `KK01_Executive_Financial_Control_Center.xlsx`. A página inicial é o Dashboard 01.
2. No menu funcional **04 — Financiamento**, editar separadamente os painéis **Base de Referência** e **Comparativo**. As TAN predefinidas vão de 2% a 8%, de 0,5 em 0,5 pontos percentuais.
3. No menu funcional **07 — Hotel & MICE**, editar os dois painéis de ADR, ocupação, crescimento, MICE, fees e FF&E. `Profile` usa o perfil selecionado nos detalhes; uma percentagem ou um ADR explícito sobrepõe apenas esse parâmetro. O ano de leitura está agora em **N41**, separado dos pressupostos operacionais.
4. Abrir **Custom e detalhes** para valores manuais (coluna C), perfis e restantes condições. Os valores efetivos estão na coluna D; a coluna E conserva os últimos valores guardados. As linhas antigas que remetem para novos parâmetros são calculadas, para evitar duas entradas concorrentes.
5. No Dashboard 01, **Resultados ativos: Base / Comparative** escolhe o contexto dos gráficos e quadros operacionais. O painel global compara sempre a **Base atual** com o **Comparativo**, independentemente desta escolha.
6. Usar o seletor global PT/EN existente. Os nomes técnicos das folhas mantêm-se; o Dashboard funcional 04 corresponde à folha `DASH 03 V06`, e o funcional 07 à `DASH 05 V06`.

Os perfis operacionais conservam o ramp-up e crescimento anual do motor. `Custom` permite alterar um parâmetro sem obrigar a criar um cenário de teste para a Base. TAN, custo efetivo do crédito e taxa de desconto dos investidores são medidas diferentes.

## Guardar e restaurar versões no Windows / Excel Desktop

**Guardar o XLSX com Ctrl+S não cria, por si só, um novo registo de versão de pressupostos.** A indicação “Unsaved” compara os pressupostos efetivos com a última versão financeira guardada. A alteração da escolha para Custom também é detetada, mesmo que o número efetivo não mude.

Para criar uma versão rastreável: guardar e fechar o Excel, manter esta pasta e a subpasta `Historico` juntas, e executar `Gerir_Cenarios.cmd`. Escolher B para guardar Base, C para Comparativo ou T para validar. O comando chama o Excel Desktop, recalcula, verifica erros de fórmula e os 180 meses, conserva os JSON anteriores e atualiza o histórico no próprio livro. Requer Windows com Microsoft Excel Desktop; não usa VBA nem Python.

Para restaurar: selecionar o ID em `V07 Estado!B19`, guardar e fechar o livro e escolher R no comando. Também é possível executar:

```powershell
.\Gerir_Cenarios_Excel365.ps1 -Action SaveBase
.\Gerir_Cenarios_Excel365.ps1 -Action SaveComparative
.\Gerir_Cenarios_Excel365.ps1 -Action Restore -VersionID BASE-...
```

O auxiliar PowerShell foi fornecido para execução no computador do utilizador. **Não foi executado neste ambiente Linux.** Os testes de guardar/restaurar incluídos na entrega executaram a alternativa Python com recálculo independente.

## Relatórios — DASH 10

O acesso principal abre `V07 Relatorios`. Estão disponíveis quatro formatos: síntese executiva de uma página A4, relatório bancário, relatório de investidores e relatório integral. Os últimos três não têm limite fixo de páginas. A configuração inicial é **Base Only**: Base atual guardada, sem comparações obrigatórias.

A emissão controlada recalcula o livro inteiro, reconcilia os fluxos e **bloqueia alterações da Base ainda não guardadas**. Cada emissão conserva PDF e JSON com versão, destinatário, pressupostos, resultados, fluxos mensais e hashes. Relatórios anteriores não são substituídos.

Os exemplos PDF estão em `Relatorios`: quatro formatos PT/EN e um exemplo bancário Base vs Comparativo. As tabelas de orçamento, unidades e sensibilidades ficam congeladas no JSON, além dos pressupostos e fluxos. A geração automatizada faz-se no ambiente de execução do repositório, **não por uma hiperligaçāo que execute código no Excel**:

```bash
python entregas/EFCC_Final/src/scenario_v07.py save-base --description "Revisão adotada"
python entregas/EFCC_Final/src/scenario_v07.py save-comparative --description "Comparativo"
python entregas/EFCC_Final/src/report_v07.py --all --both-languages
```

Para uma comparação, selecionar `Base vs Selected`. Para várias, selecionar `Base vs Multiple` e indicar em B10 os IDs de versões comparativas, separados por vírgulas. Os comparativos guardados são recalculados com a Base atual preservada. O gerador Python requer as dependências do repositório e LibreOffice; os scripts de versionamento nativo acima são independentes desse gerador.

Para reproduzir um documento emitido sem consultar o XLSX atual:

```bash
python entregas/EFCC_Final/src/reproduce_report_v07.py caminho/relatorio.json --output copia.pdf
```

O reprodutor recusa substituir um ficheiro existente.

## Verificação e limites da entrega

Consultar `RELATORIO_IMPLEMENTACAO.md`, os JSON de testes e `Matriz_Navegacao.csv`. A validação estrutural usa Open XML SDK; o recálculo independente usa LibreOffice. **Não se afirma execução nem aceitação nativa no Microsoft Excel 365 Desktop**, indisponível no ambiente.

O binário do logótipo original não consta do pacote de origem; a área de inserção `V06 Identidade` foi preservada. Não se redesenhou a marca. Não foi disponibilizada medição certificada da área bruta exclusiva do Aparthotel; a área global de 47 000 m² continua identificada como proxy, não como área certificada do Aparthotel. Condições sem prova documental permanecem hipóteses.
