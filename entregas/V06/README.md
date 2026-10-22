# KK-01 V06 — Executive Master Candidate

The workbook contains nine connected V06 dashboards and preserves the complete V04 worksheets as historical reference. Financially and structurally tested in available tools; native Microsoft Excel 365 Desktop certification is NOT TESTED. Exact V05 financial parity is not certified; see the reconciliation and assumptions.

## Build and verify

Python 3.12; xlsxwriter, openpyxl, scipy, lxml, reportlab and PyMuPDF. LibreOffice is needed for supplementary recalculation and selector tests. Extract the authoritative V04 Fixed from fontes/KK01_Fontes_Preservadas.zip; verify its hash against Manifesto_Fontes.json. Preserve the original and work only on copies.

Each cloud task is already isolated: use the existing checkout; do not create a Git worktree unless explicitly requested.

0. Extract fontes/KK01_Fontes_Preservadas.zip to a read-only source folder. The original complete V04 delivery ZIP remains preserved in the onboarding workspace and was not duplicated in this release.
1. Set KK01_WORKDIR to a writable absolute temporary build directory.
2. Run `python src/build.py` to create V06_modulos.xlsx and Referencia_Python_V06.json.
3. Run `python src/recalculate.py "$KK01_WORKDIR/V06_modulos.xlsx" "$KK01_WORKDIR/full_recalc"`.
4. Run `python src/validate.py "$KK01_WORKDIR/full_recalc/out/V06_modulos.xlsx"`.
5. Run `python src/test_selectors.py` to mutate real inputs and force recalculation.
6. Run `python src/merge.py /path/to/KK01_Modelo_Integrado_V04_Excel_Fix.xlsx "$KK01_WORKDIR/full_recalc/out/V06_modulos.xlsx" KK01_V06_Executive_Master.xlsx`.
7. Run `python src/validate.py KK01_V06_Executive_Master.xlsx --final`.
8. Run `python src/finish.py` to create reports and checksums. Package release files separately after validation.
9. Follow KK01_V06_Compatibilidade_Excel365.pdf for native Microsoft acceptance.

The source modules in src/v04 are preserved construction/recovery material. Do not execute the legacy build/publish scripts blindly: they use legacy paths and are not the V06 build entry point.

## Source limitations

Full V05 workbooks exceeded the available file-transfer limit. Eight calendars, some taxes and several convention details are explicitly reconstructed assumptions. IPI/customs/stamp inputs without an authoritative quantified source remain unresolved exposures; zero is not a granted exemption. Actual capital/cost fields are not fabricated. The model is a financial study, not legal or fiscal certification.
