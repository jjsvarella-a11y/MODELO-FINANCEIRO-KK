# KK-01 V08 — entrega consolidada

Abra **KK01_Executive_Financial_Control_Center.xlsx**. O ficheiro abre no Dashboard Executivo e mantém o motor financeiro e o histórico de cenários.

- Base inicial preservada: TAN 6%, carência de capital 42 meses, prazo de amortização 120 meses após carência.
- PT/EN: seletor no cabeçalho do Dashboard Executivo, célula W5; acessível nos cabeçalhos dos dez dashboards.
- Base e Comparativo: filtros próprios no financiamento e Hotel & MICE. Outros inputs da Base estão diretamente nas páginas; o Comparativo completo mantém o gestor independente.
- Guardar/restaurar versões: ferramenta existente `Gerir_Cenarios.cmd`, com o Excel fechado. O mapa de entradas foi atualizado. A execução desta ferramenta no Excel Desktop continua por testar neste ambiente.
- A emissão dos relatórios exige Base guardada, recálculo e reconciliação. O gerador cloud existente é `../src/report_v07.py`; os quatro PDFs abaixo são emissões congeladas de referência, não ficheiros que mudam quando se editam os parâmetros.

## Relatórios de referência

1. [One-Page Executive](Relatorios/Referencia_V08/01_Executive_One_Page_PT.pdf)
2. [Bank Financing](Relatorios/Referencia_V08/02_Bank_Financing_PT.pdf)
3. [Equity Investor](Relatorios/Referencia_V08/03_Equity_Investor_PT.pdf)
4. [Comprehensive Project](Relatorios/Referencia_V08/04_Comprehensive_Project_PT.pdf)

Cada PDF tem um JSON de pressupostos/resultados e um identificador de emissão. Relatórios e versões anteriores foram preservados.

[Correções, testes e exceções](RELATORIO_IMPLEMENTACAO.md) · [Evidência dos testes](Verificacao_V08/)

**Âmbito:** a fonte foi a versão do GitHub. Não estão incluídas as alterações da cópia local enviada por OneDrive. A aplicação Microsoft Excel 365 Desktop não foi executada; os testes disponíveis são estruturais e de recálculo independente. O logótipo original em ficheiro e a ABC certificada continuam por disponibilizar/confirmar.
