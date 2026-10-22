# Inserir o logótipo no próprio modelo

1. Abra `KK01_V06_Executive_Master.xlsx` no Excel e clique em **Inserir logótipo** no cabeçalho, ou abra `V06 Identidade`.
2. Use **Inserir → Imagens → Este dispositivo**. Escolha o JPG/PNG original e coloque **uma única imagem** na área B8:H20. Pode também colá-la nessa folha. Guarde o livro: o ficheiro da imagem fica incorporado no XLSX.
3. Para replicar o original nos nove dashboards, nas dez páginas do relatório e nas duas do brief, execute `Executar_Validacao.cmd` no Windows com Excel instalado. O script aplica as imagens **numa cópia**, testa os cálculos, exporta os PDFs e guarda a cópia e os resultados na nova pasta `Resultados_...`.

A área de identidade é independente das fórmulas financeiras. Não utiliza VBA ou a função IMAGE(). A inserção central não promete uma sincronização automática sem executar o passo 3. A rotina COM de sincronização está fornecida, mas **não foi executada neste ambiente Linux**. Se falhar, o teste regista a falha sem impedir o diagnóstico financeiro.

Microsoft Excel 365 continua **NOT TESTED** até execução real. Não considerar os `#VALOR!` resolvidos no Excel apenas porque os ensaios independentes passaram.

# Comportamento reportado dos #VALOR!

O utilizador confirmou que os valores/legendas aparecem antes de ativar edição e que os erros surgem depois de ativar edição e reabrir. O teste agora tenta explicitamente **Vista Protegida → Ativar Edição → cálculo automático → recálculo completo → guardar → fechar → reabrir**. Regista os valores guardados antes de editar e os erros imediatamente depois. Os primeiros registos priorizam o motor central e os calendários, para evitar confundir a origem com os efeitos nos dashboards. São guardados até 2.000 registos de erro.

Se a API de Vista Protegida não estiver disponível ou for bloqueada por política, esse passo fica NOT TESTED; o script prossegue com a abertura normal para diagnóstico. Esse fallback não comprova que o caso reportado foi reproduzido.
