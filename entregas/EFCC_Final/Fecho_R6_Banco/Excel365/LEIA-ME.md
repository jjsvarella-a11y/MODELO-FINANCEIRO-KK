Validação nativa pendente. Este script foi preparado para Windows com Microsoft Excel 365 Desktop; não foi executado no ambiente Linux.

Executar numa pasta local de confiança, com Excel fechado:

```powershell
.\Emitir-Banco-Excel365.ps1 -Source '..\KK01_Executive_Financial_Control_Center_R6_BANK_CHECKPOINT.xlsx' -Recipient 'Nome da instituição' -Scenario 'Proposal Bank A' -Comparison 'Official KK-01 6%'
```

O original é copiado, nunca gravado. A emissão verifica recálculo completo, erros de fórmulas, saldos dos 180 meses, gravação e reabertura. Mantém alertas visíveis: se surgir reparação, interromper e guardar o registo XML. Não considerar essa abertura aprovada. A Base Oficial é TAN 6%, carência 42 meses.

O teste deve também confirmar visualmente navegação dos dez dashboards, seletor PT/EN, cenários operacionais, impressão e matriz ADR × ocupação. O script não certifica sozinho estes comportamentos visuais.
