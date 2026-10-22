# Geração controlada

Execute a partir da raiz do repositório. Dependências: Python 3, xlsxwriter, lxml, scipy, numpy, openpyxl, Pillow, PyMuPDF, reportlab; `soffice` para os testes independentes. Não há Excel Desktop neste ambiente.

1. Criar `/workspace/kk01_efcc/checkpoints`; copiar o master `entregas/V06_DASH10/KK01_V06_Executive_Master.xlsx` para `C0_V06_DASH10.xlsx`. Se o JPG/PNG original estiver acessível, definir `KK01_LOGO` para o caminho absoluto. Não utilizar reconstrução de imagem.
2. Executar `src/preflight.py` e `src/build_final.py`. O gerador reutiliza o motor existente e os dados V04 de `entregas/V06/src/v04`. O build produz `V06_modulos.xlsx` e `Referencia_Python_V06.json`.
3. Executar `src/final_touches.py /workspace/kk01_efcc/V06_modulos.xlsx`.
4. Executar o recálculo com caches invalidados: `python entregas/V06/src/recalculate.py /workspace/kk01_efcc/V06_modulos.xlsx /workspace/kk01_efcc/final_recalc`. Depois `src/validate_final.py` sobre o ficheiro de saída. LibreOffice não certifica Excel.
5. Executar `src/assemble.py`. Recolhe apenas caches verificados sem regravar as fórmulas pelo LibreOffice e preserva as 313 folhas legadas.
6. Executar `src/polish_delivery.py` sobre o master montado para um caminho temporário e `src/adjust_final.py` sobre esse temporário para o master final. Adiciona apenas a projeção anual de gráficos e ajusta a apresentação.
7. Executar `src/audit_ooxml.py` sobre o master final; atualizar o SHA256 em `Preservacao.json`.
8. Executar `src/test_links.py` e `src/test_extended.py`. Os laboratórios privados congelam apenas ramos independentes não testados; o motor ativo, calendários, relatório e ramo testado continuam com fórmulas. O master mantém todos os ramos dinâmicos. Nos testes de desconto, todos os 30 VAL são recalculados a partir dos vetores de caixa, que não dependem da taxa de desconto.
9. Para os PDFs, aplicar polish/adjust também ao laboratório PT e ao laboratório EN criado pelo teste de idioma. Executar `src/export_all.py` para cada idioma. As cópias privadas removem freeze panes e áreas de impressão das folhas não selecionadas. Inspecionar todos os PDFs; não truncar com PageRange.
10. Executar `src/write_docs.py` após todos os testes e `src/package.py`. O JSON completo da referência é comprimido para distribuição.

Os scripts em `Excel365` só podem ser executados em Windows com Excel instalado. Nesta entrega, o resultado nativo é **NOT TESTED**. A causa dos erros #VALOR! reportados não foi demonstrada no Excel. O logótipo permanece pendente de acesso ao ficheiro original.

Atualização solicitada pelo utilizador: executar `src/print_edge.py` e depois `src/add_identity.py` antes da auditoria final. `V06 Identidade` é a área de inserção real do JPG/PNG no Excel; os cabeçalhos abrem essa folha. A sincronização das imagens é opcional e explícita no script Windows/Excel; não depende de funções financeiras, VBA ou IMAGE(). O seu teste nativo continua pendente.

Último passo de apresentação: executar `organise_tabs.py` sobre o master final antes de atualizar manifestos e empacotar. Preserva todos os componentes exceto a visibilidade no workbook.xml; mantém V06 Motor acessível para diagnóstico.

Separação definitiva: após organizar a interface, executar `separate_legacy.py` sobre o master. O script conserva uma cópia binária da V04 em `entregas/Arquivo/V04`, audita dependências, elimina as 313 folhas legadas do ZIP e as partes inacessíveis, remove nomes locais antigos e atualiza propriedades. Executar novamente a auditoria e o recálculo independente antes de empacotar.
