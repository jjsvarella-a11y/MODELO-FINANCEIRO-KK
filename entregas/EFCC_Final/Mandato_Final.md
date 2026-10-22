# KK-01 | EXECUTIVE FINANCIAL CONTROL CENTER
## FINAL CONSOLIDATED EXECUTION MANDATE
### Financial Integrity · Investment Optimization · Premium Dashboards · Microsoft Excel 365

**Project:** KK Hotelaria e Turismo, S.A.
**Developer:** JS-CV Construções e Investimentos, Lda.
**Project location:** Kebra Kanela, Praia, Cabo Verde
**Working base:** Latest `KK01_V06_Executive_Master.xlsx`
**Authoritative recovery source:** V04 Fixed, where required
**Financial reference:** Existing V05/V06 reports, audit documentation and approved project specifications
**Official identity:** Attached JS-CV logo
**Target platform:** Microsoft Excel for Microsoft 365 Desktop
**Model currency:** EUR
**Financial horizon:** 180 months
**Interface:** Portuguese / English
**Final deliverable:** Fully integrated Executive Financial Control Center

---

# 0. EXECUTIVE MANDATE — COMPLETE THE PROJECT

This is the final consolidated development mandate for the KK-01 Executive Financial Control Center.

The project has already undergone extensive financial development, reconciliation, dashboard construction, technical corrections and graphical redesign.

The purpose of this mandate is to complete and stabilize that work.

**Do not rebuild the financial model from scratch.**

Preserve the existing V06 financial engine, recover valid V05/V04 elements where required, and implement the remaining functional, graphical and compatibility improvements.

Do not stop after producing a plan, inventory, prototype or readiness report.

Execute the work sequentially, preserving recoverable checkpoints, and deliver the completed model together with evidence of testing.

Priorities:

1. Eliminate calculation errors.
2. Resolve Excel opening and repair issues.
3. Reconcile the financial engine.
4. Complete the global PT/EN mechanism.
5. Integrate the official JS-CV visual identity.
6. Stabilize all ten dashboards.
7. Complete the financial scenario controls.
8. Implement the ADR × Occupancy sensitivity matrix.
9. Complete operator fee and reserve policy simulators.
10. Complete the Real Estate availability module.
11. Build the consolidated investment report.
12. Prepare all dashboards for PDF export.
13. Verify Microsoft Excel 365 behavior.
14. Publish all deliverables.

No new graphical element should be accepted if it compromises workbook integrity or financial calculations.

---

# PART I — FINANCIAL AND TECHNICAL STABILIZATION

## 1. Correct all unexpected Excel errors

The latest V06 screenshots still show multiple `#VALOR!` / `#VALUE!` errors.

These occur in several areas, including:

- Future CAPEX.
- Bank debt.
- Cash M30, M42 and M54.
- Construction and contingency indicators.
- Tax and insurance indicators.
- Residential financial indicators.
- Dividends.
- Investor IRR.
- MOIC.
- NPV / VAL.
- Terminal Equity Value.

These errors are blocking defects.

### Required actions

Perform a complete recalculation and dependency audit.

For every unexpected error:

1. Identify the originating formula.
2. Trace upstream dependencies.
3. Verify numeric data types.
4. Check date and text conversions.
5. Check named ranges.
6. Check links between technical sheets and dashboards.
7. Check Excel 365 function compatibility.
8. Check global-language selector dependencies.
9. Correct the originating defect.
10. Confirm downstream values recalculate correctly.

Do not hide financial errors using `IFERROR`, blank strings, artificial zeros or hard-coded historical values.

Create an auditable error-correction register.

A negative NPV is a legitimate financial result. `#VALUE!` is not.

Do not begin final investment optimization before the financial indicators have been restored.

## 2. Resolve Excel opening and content-repair errors

Previous versions triggered Microsoft Excel messages requesting content recovery.

Identify whether this issue still exists in the latest version.

If it does:

- Examine Excel repair logs where available.
- Identify the offending workbook component.
- Verify OOXML relationships.
- Examine defined names, formulas, charts, comments, shapes and styles.
- Correct problematic elements locally.
- Preserve valid financial content.
- Recalculate the workbook.
- Save, close and reopen.

Do not simply rename the workbook as an Excel Fixed version.

Do not use LibreOffice compatibility as evidence of native Microsoft Excel compatibility.

## 3. Financial engine integrity

Preserve and reconcile:

- 180 monthly periods.
- Construction investment.
- 94 residential units.
- Hotel with 200 rooms.
- MICE facility.
- Bank financing.
- Eight new-investor capital schedules.
- Fifteen shareholder blocks.
- VAT.
- IRPC.
- Insurance.
- Operator remuneration.
- FF&E.
- HOTUSA contingency.
- Cash flow.
- Dividend policy.
- Exit waterfall.
- Individual investor returns.

Maintain the existing financial model as the single source of truth.

No duplicated independent calculation engine may produce competing financial results.

Use the documented V05 results as scenario-specific regression benchmarks, not hard-coded outputs.

---

# PART II — PREMIUM VISUAL IDENTITY AND EXCEL NAVIGATION

## 4. Official JS-CV logo

Use the official logo supplied by the developer:

`logo jscv(1).jpg`

Embed the image into the Excel workbook.

Preserve:

- Original architectural symbol.
- JS-CV lettering.
- Corporate colors.
- Original proportions.
- Corporate signature.

Do not redesign, approximate or reinterpret the logo.

### Positioning

Place the logo in the upper-right portion of the red institutional header rectangle, with the project identity on the left.

Because the original logo contains a white background and red/black/grey artwork, use an appropriately proportioned white inset area where necessary.

Avoid distortion, stretching or loss of contrast.

The logo must remain embedded in the workbook and appear correctly in exported PDFs.

If a transparent derivative is created, preserve the original artwork precisely.

## 5. Global language selector

Place the PT | EN selector immediately below the red institutional rectangle.

The active language must be clearly visible.

The selector must control:

- Sidebar navigation.
- Dashboard titles.
- Section headings.
- KPI labels.
- Filters.
- Scenario descriptions.
- Tables.
- Explanatory messages where possible.
- Report titles.
- DASH 10 consolidated report.

Use centralized translations.

Changing language must not alter numeric financial calculations.

If long technical explanations cannot be translated robustly, use English as the fallback technical language.

## 6. Dashboard navigation

Adopt exactly ten numbered destinations:

**01 — Dashboard Executivo**
**01 — Executive Dashboard**

**02 — Investimento e Orçamento de Construção**
**02 — Investment & Construction Budget**

**03 — Fiscalidade, Incentivos e Seguros**
**03 — Taxation, Incentives & Insurance**

**04 — Desempenho Financeiro Consolidado**
**04 — Consolidated Financial Performance**

**05 — Promoção e Comercialização Imobiliária**
**05 — Real Estate Development & Sales**

**06 — Exploração Residencial e Sell & Leaseback**
**06 — Residential Operations & Sell & Leaseback**

**07 — Exploração Hoteleira e Centro MICE**
**07 — Hotel & MICE Operations**

**08 — Estrutura Acionista e Blocos de Investimento**
**08 — Shareholding & Investment Blocks**

**09 — Rentabilidade dos Acionistas e Distribuições**
**09 — Equity Returns & Distributions**

**10 — Relatório Executivo Consolidado de Viabilidade e Investimento**
**10 — Consolidated Investment & Feasibility Report**

The numbering is for executive navigation.

Preserve internal worksheet names and formula references unless there is a documented reason to change them.

Display only the selected language in the sidebar.

Allow controlled line breaks for longer titles.

Highlight the active dashboard in institutional red.

## 7. Screen delimitation and layout

Use approximately `A1:Q64` as the initial executive composition reference.

This is a visual guide, not a hard physical limit.

Each dashboard must present:

- Institutional header.
- Fixed sidebar.
- Fixed top navigation.
- Main financial content area.
- No worksheet gridlines.
- No technical row/column headers.
- Consistent spacing.
- Clearly organized cards and tables.
- Appropriate zoom.

Use Freeze Panes so the header remains visible during vertical scrolling and the sidebar remains visible during horizontal scrolling.

Avoid zoom levels of approximately 30% that make the dashboard unreadable.

Allow controlled expansion of tables, rows, columns and financial reports.

Do not use fixed merged-cell structures that prevent dynamic content expansion.

## 8. Currency and numerical presentation

Display monetary amounts with a clear currency unit.

Preferred formats:

- `63 168 000 €`
- `63,17 M€`
- `2 507 809 €`
- `−9,67 M€`

Apply pt-PT or EN regional presentation consistently.

Use real numeric Excel formats.

Do not transform monetary results into text.

Do not append euros to ratios, percentages, quantities or MOIC.

Use negative values consistently and visibly.

---

# PART III — TEN EXECUTIVE DASHBOARDS

Maintain the existing nine-dashboard structure and add the consolidated report as DASH 10.

Each dashboard must contain:

- Clearly displayed active scenario.
- Linked financial KPIs.
- Relevant charts.
- Scenario controls, where appropriate.
- Contextual KPI explanations.
- Navigation to related modules.
- Print / PDF access.
- Executive interpretation.
- Financial and contractual warnings where necessary.

All values must come from the live model.

---

# PART IV — HOTEL OPERATING ASSUMPTIONS

## 9. DASH 07 main simulator

Create a professional Hotel Operating Assumptions panel with the following primary parameters:

| Parameter | Sensitivity range | Step |
|---|---|---|
| ADR | 120–250 € | 5 € |
| Hotel Occupancy | 45%–78% | 1 pp |
| MICE Utilization | 20%–80% | 5 pp |
| FF&E Reserve | 2%–6% | 0.5 pp |
| Annual ADR Growth | 0%–6% | 0.5 pp |

These are study ranges and must not be represented as validated market forecasts.

The Base Case must reproduce the existing documented assumptions.

### Additional controls

Provide:

- Conservative.
- Base Case.
- Favorable.
- Upside Stress.
- Custom.

Preserve the hotel ramp-up and seasonality.

Do not use national bed occupancy directly as the hotel's room occupancy.

## 10. Operational results

Calculate:

- Available Room Nights.
- Occupied Room Nights.
- Room Revenue.
- RevPAR.
- Total Revenue.
- TRevPAR.
- GOP.
- GOPPAR.
- EBITDA.
- EBITDA Margin.
- Operator Fees.
- FF&E allocations.
- CFADS.
- Cash generation.
- Investor return impact.

Do not assume that all incremental revenue becomes EBITDA.

Hotel-room revenue, F&B and MICE revenues must be distinguished.

## 11. MICE Utilization

Link MICE utilization to:

- Available event-day capacity.
- Event days sold.
- Event volume.
- Participants.
- Revenue per event.
- Direct costs.
- Event F&B.
- MICE contribution margin.

The 1,500-participant venue capacity must not be treated as average attendance.

Preserve the underlying operational model where documentary inputs are insufficient.

Avoid duplication between hotel, events and F&B revenues.

## 12. FF&E Reserve

Calculate the reserve contribution and protected balance.

Distinguish:

- Initial FF&E already included in CAPEX.
- Annual FF&E reserve.
- Actual replacement expenditure.
- Cumulative reserve balance.
- Eligible cash after reserve.
- Terminal reserve treatment.

Do not double count FF&E as an operating expense and restricted cash movement.

---

# PART V — ADR × OCCUPANCY INVESTMENT SENSITIVITY

## 13. Dynamic sensitivity matrix

Implement a fully calculated matrix:

**PT:** Matriz de Sensibilidade ADR × Ocupação.

**EN:** ADR × Occupancy Sensitivity Matrix.

Suggested layout:

ADR columns:
150 €, 175 €, 200 €, 225 €, 250 €.

Occupancy rows:
50%, 55%, 60%, 65%, 70%, 75%.

Allow appropriate adjustments.

Each cell must calculate the selected financial metric through the underlying model.

### Metric selector

- New Investors NPV.
- New Investors IRR.
- Hotel EBITDA.
- RevPAR.
- Exit Equity Value.

### Discount-rate selector

- 8%.
- 10%.
- 12%.
- Custom.

### Break-even analysis

Identify:

- ADR needed for NPV ≥ 0.
- Occupancy needed for NPV ≥ 0.
- Required stabilized EBITDA.
- Distance from Base Case.
- Required investor return.

If no combination achieves a non-negative NPV, report that finding explicitly.

Do not extrapolate silently.

### Visual treatment

Use actual numerical values.

Muted red for negative NPV.

Amber for near break-even.

Muted green for positive NPV.

Do not use fixed illustrative heatmap results.

---

# PART VI — INTERNATIONAL OPERATOR FEE STRUCTURE

## 14. Hotel operator terms

Create the panel:

**PT:** Estrutura de Remuneração da Operadora Internacional.

**EN:** International Operator Fee Structure.

Four profiles:

1. Operator Proposed Terms.
2. KK Target Terms.
3. KK Optimized Terms.
4. Custom Terms.

### Proposed scenario parameters

| Component | Reference | KK Target | KK Optimized |
|---|---:|---:|---:|
| Base Management Fee | 4% | 2.5% | 2% |
| Incentive Management Fee | 10% | 6% | 5% |
| Marketing / Distribution / Systems | 4% aggregate | 3% aggregate | 1.5% aggregate |

The original reference must be verified against actual operator documentation.

Do not label undocumented amounts as formal contractual proposals.

### Editable controls

- Base Management Fee.
- Incentive Management Fee.
- Marketing Fee.
- Distribution Fee.
- Systems Fee.

Permit independent customization.

Document each fee's calculation base.

Do not simply add percentages applied to different bases.

### Optional negotiation mechanism

Provide an optional Owner's Priority Return mechanism for incentive fee eligibility.

Keep it disabled unless explicitly adopted as a contractual scenario.

Changes must affect actual fees, operating cash flow, debt coverage, dividends, investor NPV and exit valuation.

---

# PART VII — RESIDENTIAL OPERATOR CONDITIONS

## 15. DASH 06 residential controls

Provide equivalent fee scenarios for serviced residences.

Indicative study parameters:

| Component | Reference Study | KK Target | KK Optimized |
|---|---:|---:|---:|
| Management Fee | 8% | 6% | 5% |
| Marketing / Distribution / Systems | 3% | 2% | 1.5% |

Allow Custom Terms.

Preserve:

- Residential revenue.
- Residential OPEX.
- Operator fees.
- Owner remuneration.
- Owner yield.
- Maintenance reserve.
- Residential FF&E reserve.
- Available cash.

Do not combine residential fees or revenues incorrectly with the hotel operating model.

---

# PART VIII — RESERVE POLICY SIMULATOR

## 16. Three selectable reserve policies

Create a panel in DASH 09:

**PT:** Política de Reservas.

**EN:** Reserve Policy.

Provide:

| Policy | General liquidity reserve |
|---|---:|
| Prudente / Conservative | 3,000,000 € |
| Equilibrada / Balanced | 2,500,000 € |
| Flexível / Flexible | 2,000,000 € |

The Conservative setting should remain the reference policy unless the documented V06 baseline establishes another approved assumption.

These are study options, not permission to distribute restricted funds.

Maintain separately:

- HOTUSA 1,000,000 €.
- FF&E.
- Taxes payable.
- Debt-related restricted reserves.
- Other contractual obligations.

Preserve the 1.20× prior-12-month coverage test.

Do not assume reducing the liquidity reserve automatically increases dividends if another constraint remains binding.

Show the impact on:

- Available cash.
- Dividend capacity.
- Debt coverage.
- Investor IRR.
- Investor NPV.
- MOIC.
- Terminal Equity Value.

---

# PART IX — REAL ESTATE AVAILABILITY

## 17. Unit Availability view

From DASH 05, provide an icon to access:

**PT:** Disponibilidade de Frações.

**EN:** Unit Availability.

Use the existing single inventory of 94 apartments.

Show:

- Unit ID.
- Type.
- Floor.
- Interior area.
- Other commercial area, where applicable.
- Price per m².
- Total sale price.
- Commercial status.
- Deposit percentage.
- Deposit amount.
- Deposit date.
- Planned instalments.
- Due dates.
- Actual cash received.
- Outstanding receivables.
- Expected deed date.

Allow filters by:

- Apartment type.
- Floor.
- Area.
- Price.
- Availability status.
- Payment schedule.

Distinguish sold, reserved, available and deeded units.

Preserve actual contracted terms separately from forecast assumptions.

Reconcile every unit's payment schedule and the consolidated monthly Real Estate cash flow.

---

# PART X — SHAREHOLDER RETURNS AND FINANCING

## 18. Capital structure

Preserve the fifteen investment blocks:

- JSCV: 40%.
- JALS: 10%.
- New Investors: 50%.

The new-investor funding target is EUR 25 million through eight investment blocks.

Maintain separate subscription schedules and distinguish primary capital from secondary transfers.

## 19. Dividend policy

Preserve the documented flexible dividend targets:

- Platinum: 7%–10%.
- Gold: 5%–6.5%.
- Silver: 4%–5%.

Aggregate annual targets:

- Minimum: EUR 3 million.
- Intermediate: EUR 3.575 million.
- Maximum: EUR 4.15 million.

First eligible distribution: M54.

Preserve all binding liquidity, profit and debt restrictions.

## 20. Debt optimization

Provide sensitivity for:

- Bank interest rate.
- Principal repayment schedule.
- Loan tenor.
- Grace period.
- Cash sweep.
- Reserve requirements.

Evaluate consequences for:

- Total interest.
- Debt outstanding.
- Liquidity.
- Debt coverage.
- Dividend capacity.
- Investor NPV.

Do not confuse deferred debt repayment with an improvement in the project's underlying operating performance.

---

# PART XI — DASH 10 CONSOLIDATED REPORT

## 21. Consolidated Investment & Feasibility Report

Add:

**10 — Relatório Executivo Consolidado de Viabilidade e Investimento**

**10 — Consolidated Investment & Feasibility Report**

This dashboard must automatically capture all active assumptions and results from DASH 01–09.

Generate a complete institutional report of up to ten readable A4 pages.

### Ten-page report structure

**Page 1 — Executive Investment Summary**

Project description, principal KPIs, active scenario, investment assessment and alerts.

**Page 2 — Development Budget & Funding**

CAPEX, construction, historical investment, equity, financing and funding structure.

**Page 3 — Cash Flow & Debt Sustainability**

Cash, debt, interest, DSCR, reserves, funding gaps and coverage.

**Page 4 — Real Estate Development & Sales**

Apartment sales, inventory availability, prices, deposits, instalments and commercial cash flow.

**Page 5 — Residential Owner Returns**

Residential operating results, owner yield, costs, fees and FF&E.

**Page 6 — Hotel & MICE Operations**

ADR, occupancy, utilization, revenues, GOP, EBITDA, operator fees and FF&E.

**Page 7 — Taxation, Operator Terms & Reserves**

IRPC, VAT, insurance, applicable benefits, HOTUSA and active fee structures.

**Page 8 — Shareholding & Equity Returns**

Investment blocks, dividends, IRR, NPV, MOIC and equity waterfall.

**Page 9 — Sensitivities, Risks & Break-even**

ADR × Occupancy matrix, NPV sensitivities, debt stress, reserve policies and downside exposures.

**Page 10 — Investment Assessment & Key Decisions**

Consolidated assessment for the Board, investors, lenders, residential buyers and international operators.

### Report controls

Provide:

- Portuguese / English.
- Complete Report / Executive Brief.
- Active Scenario / Reference Case.
- Report reference date.
- Scenario ID.

Ensure the report can be reproduced from the recorded assumptions.

Do not label a report as updated if its financial recalculation has not completed.

### Stakeholder attractiveness

Evaluate separately:

- Promoter.
- Equity investors.
- Banks.
- Apartment buyers.
- Hotel operator.

Use documented financial thresholds.

Where data are insufficient, report the limitation rather than inventing conclusions.

---

# PART XII — PDF PRINTING FOR EVERY DASHBOARD

## 22. Print / Export PDF controls

Every dashboard must contain a visible control:

**PT:** Imprimir / Exportar PDF.

**EN:** Print / Export PDF.

Configure:

- Print area.
- Page orientation.
- Paper size.
- Scaling.
- Margins.
- Headers.
- Footers.
- Report name.
- Project name.
- Logo.
- Scenario reference.
- Page numbering.

For the consolidated report, use A4.

For wider dashboards, use A3 landscape where necessary.

The PDF must display the active scenario data and must remain readable.

Do not create a decorative button that has no real action.

If native Excel cannot provide automatic one-click PDF export without VBA, use the supported Print / Save as PDF workflow and clearly identify the action.

Do not introduce mandatory macros merely for printing.

Generate preview PDFs and inspect all pages for clipping, overlap, missing labels and incorrect currency formatting.

---

# PART XIII — KPI DICTIONARY AND EXPLANATIONS

## 23. KPI documentation

Maintain the bilingual KPI Dictionary.

Each important KPI should have:

- Name.
- Definition.
- Calculation.
- Data source.
- Interpretation.
- Target or relevant threshold.
- Limitations.

Provide contextual links from executive cards and tables.

Prioritize:

- IRR.
- NPV.
- MOIC.
- ADR.
- RevPAR.
- TRevPAR.
- GOP.
- GOPPAR.
- EBITDA.
- DSCR.
- CFADS.
- GDV.
- CAPEX.
- EAC.
- FF&E.
- Enterprise Value.
- Equity Value.

Technical explanations may use English when bilingual presentation would compromise Excel stability.

---

# PART XIV — MICROSOFT EXCEL 365 CERTIFICATION

## 24. Mandatory native testing

If Microsoft Excel for Microsoft 365 Desktop is available:

1. Open the completed workbook.
2. Confirm no content-repair warning.
3. Perform full calculation.
4. Confirm no unexpected formula errors.
5. Test PT/EN.
6. Test all ten dashboard links.
7. Test ADR and occupancy controls.
8. Test MICE utilization.
9. Test FF&E reserve.
10. Test hotel operator fee scenarios.
11. Test residential operator fee scenarios.
12. Test reserve policies.
13. Test debt scenarios.
14. Test the ADR × Occupancy sensitivity matrix.
15. Test the Real Estate availability view.
16. Test DASH 10 report.
17. Export and inspect PDF reports.
18. Save, close and reopen the workbook.
19. Confirm no repair warning after reopening.
20. Reconcile financial results after the native Excel save.

### Certification standard

Use only these states:

- PASS — Tested successfully.
- FAIL — Tested and failed.
- NOT TESTED — Environment unavailable.

Do not certify Excel 365 compatibility based solely on LibreOffice, Python, cached formula values or OOXML structural checks.

If native Excel testing is unavailable, deliver a candidate version and explicitly identify the outstanding Microsoft 365 acceptance test.

---

# PART XV — FINAL DELIVERABLES AND CLOSURE

## 25. Required outputs

Deliver:

1. `KK01_Executive_Financial_Control_Center.xlsx`
2. Complete technical and financial audit report.
3. Excel 365 compatibility report.
4. Financial regression and reconciliation report.
5. Register of corrected `#VALOR!` errors.
6. Ten executive dashboard PDFs.
7. DASH 10 consolidated investment report.
8. Two-page Executive Brief.
9. Active Scenario Parameters Register.
10. Updated bilingual KPI Dictionary.
11. Source code and automated tests.
12. Final deliverable ZIP.
13. SHA-256 delivery manifest.

Preserve earlier checkpoint versions for recovery.

## 26. Final acceptance requirements

The work is only considered financially and functionally complete when:

- Unexpected `#VALOR!` errors are eliminated.
- The financial engine reconciles.
- Language switching operates correctly.
- All ten dashboards are navigable.
- The JSCV logo is correctly embedded.
- The dashboard layout is professionally formatted.
- Numeric indicators display appropriate units.
- Inputs and scenario controls propagate correctly.
- The sensitivity matrix produces actual calculated results.
- Reserve policies respect binding constraints.
- The Real Estate availability schedule reconciles.
- Operator fee controls function.
- PDF exports are legible and complete.
- The consolidated report reflects the active scenario.
- All supported technical tests pass.
- Native Excel 365 tests have explicit evidence or are honestly declared pending.

A visually complete workbook with broken financial calculations is not acceptable.

A financially reconciled workbook with unusable dashboards is also not acceptable.

---

# FINAL EXECUTION INSTRUCTION

**Complete and stabilize the KK-01 Executive Financial Control Center.**

Do not repeat the development cycle from the beginning.

Do not stop at an implementation plan.

Do not hide errors.

Do not create fake controls or static dashboard results.

Do not claim native Excel 365 certification without actual native testing.

Preserve all validated financial functionality.

Optimize presentation without compromising integrity.

Deliver the completed workbook, its financial evidence, the ten PDF dashboards and the consolidated investment report.

**Final objective: a reliable, professionally presented and auditable financial control system capable of evaluating the development, financing, operations and stakeholder returns of the KK-01 project under configurable scenarios.**

EXECUTE, VALIDATE AND DELIVER.
