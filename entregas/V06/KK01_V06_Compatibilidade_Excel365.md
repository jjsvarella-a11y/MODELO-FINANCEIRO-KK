# KK-01 V06 — Compatibilidade Excel 365

22/10/2026 · V06 CANDIDATE

## Resultado Microsoft Excel 365 Desktop

NOT TESTED — Microsoft Excel 365 Desktop indisponível neste ambiente Linux. O livro é candidato para validação nativa, sem certificação de abertura, recálculo ou reabertura em Microsoft Excel. Excel Web e WPS: não testados.

## Evidência disponível

ZIP/XML, relações, referências de estilos/strings, ordem e unicidade de células, limite de 8.192 caracteres por fórmula, ausência de erros em cache, navegação e recálculo externo foram verificados. Fórmulas no livro final: 2031372; folhas: 384. O recálculo externo foi forçado pela remoção de caches numa cópia privada.

## Relatório de reparação recebido

error093440_01.xml refere remoção de informações de células na V05.1, incluindo sheet1 e várias folhas de pipeline de cenários. O ficheiro V05.1 não pôde ser inspecionado; a causa exata não foi confirmada. Não se afirma que a falha anterior tenha sido definitivamente resolvida. Foram criados novos componentes V06 e preservada a V04 verificada, sem copiar partes gráficas V05 potencialmente corrompidas.

## Protocolo de aceitação pendente

1. Abrir KK01_V06_Executive_Master.xlsx em Microsoft Excel 365 Desktop.

2. Confirmar ausência de aviso de reparação e de conteúdo removido.

3. Executar recálculo integral Ctrl+Alt+Shift+F9.

4. Verificar ausência de erros inesperados e comparar KPI com Referencia_Python_V06.json.

5. Guardar com outro nome, fechar Excel e reabrir.

6. Confirmar novamente ausência de reparação.

7. Testar os nove destinos da sidebar e os links de informação para o dicionário.

8. Alterar PT para EN em V06 Controlos; confirmar títulos e navegação sem alterar valores financeiros.

9. Testar juro, IVA e atraso de capital, observando caixa, dívida, dividendos e retornos.

10. Restaurar os valores base da tabela G:H em V06 Controlos.

11. Verificar impressão A3, gráficos e ausência de #### ou cortes.

12. Registar versão/build do Excel, resultados, screenshots e hashes antes/depois.

13. Se houver reparação: guardar error*.xml e a cópia reparada; não substituir o original. Identificar componentes/células e comparar com a cópia não reparada.

## Conclusão

Compatibilidade Microsoft 365: PENDENTE. A ausência de erros em LibreOffice e a conformidade estrutural não substituem o teste no Microsoft Excel.
