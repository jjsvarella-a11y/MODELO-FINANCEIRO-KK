O ZIP de recuperação contém fontes, referência Python, R4 e R5, preservando a hierarquia do repositório. O master R6.1 e o Snapshot emitido estão disponíveis separadamente nesta pasta, com hashes no manifesto; não são duplicados dentro do ZIP.

Para reconstruir a extensão a partir da R5 fornecida, no ambiente preparado:

```bash
python entregas/EFCC_Final/src/bank_r6.py
python entregas/EFCC_Final/src/bank_finish.py
python entregas/EFCC_Final/src/bank_cash_guards.py
python entregas/EFCC_Final/src/bank_ui_audit.py
python entregas/EFCC_Final/src/bank_polish.py
python entregas/EFCC_Final/src/bank_domains.py
python entregas/EFCC_Final/src/bank_view.py
python entregas/EFCC_Final/src/bank_chart_parts.py
python entregas/EFCC_Final/src/release_labels.py
```

É obrigatório recalcular integralmente, reconciliar e preencher caches verificados antes de entregar uma reconstrução. Os comandos de geração, por si só, não validam resultados. Usar `test_bank_r6.py` e os testes adicionais; o emissor `bank_emit.py` executa recálculo integral e bloqueia a emissão em caso de falha.

A validação Open XML é fornecida em `src/OpenXmlValidation`; exige .NET 8 e restauro dos pacotes oficiais. O Excel 365 Desktop é validado separadamente com o kit Windows. Não considerar esta recuperação ou um teste LibreOffice como prova de aceitação nativa.
