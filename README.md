# América Móvil (AMX): DCF + comparables valuation

An investment-banking interview-preparation project. It builds a complete valuation of **América Móvil, S.A.B. de C.V.** (BMV: AMX / NYSE: AMX) from primary SEC filings, so every line can be defended in a superday.

**Result (valuation date 24-Nov-2026):**
- DCF fair value: **MXN 25.54 per share** (USD 26.23 per ADR), against a price of MXN 23.98.
- Upside: **+6.5% → HOLD**, under a ±15% rule set before the model was run. No price target is issued.
- WACC 10.85%, terminal growth 3.0%. Terminal value is **67% of EV** (below the 70% red line, but close).
- Sensitivity range (WACC ±1 pp × g ±1 pp): MXN 19.6–35.5.
- Comps cross-check: AMX trades at 5.2x run-rate EV/EBITDA against a 4.6x peer median, which implies about MXN 20.

## Repository contents
| Path | What it is |
|---|---|
| `model/AMX_Valuation_Model.xlsx` | The model. Sheets: Cover · Inputs · Historicals · Drivers · Projections · WACC · DCF · Sensitivity · Comps · Checks |
| `docs/01_investment_memo.md` (+ `.pdf`) | 2-page investment memo |
| `docs/02_defense_sheet.md` | Answers to the six defence questions (a)–(f) |
| `docs/03_sources.md` | Every source, with dates accessed and known data gaps |
| `docs/scenarios.md` | Full-model reruns (FX ±10%, Desktop in, pension pre-tax, ARPU split, bear/bull) |
| `docs/model_audit.txt` | Automated scan for hard-coded numbers inside formulas (result: 0 of 2,421 formulas) |
| `scripts/build_model.py` | Generates the workbook from sourced inputs (reproducible) |
| `scripts/recalc.sh` | Recalculates the workbook headlessly with LibreOffice |
| `scripts/audit_model.py` | Hard-code audit |
| `scripts/scenarios.py` | Scenario reruns |
| `data/xbrl_hist_2021_2025.json` | FY2021–FY2025 figures extracted from AMX's 20-F XBRL (SEC EDGAR) |
| `data/input_map.json` | Input name → cell map, used by the scenario script |

## How the model works
1. **Inputs.** Every assumption is a blue cell with a source and a one-line rationale. The DCF, WACC, Sensitivity and Comps sheets contain formulas only.
2. **Historicals.**
   - FY2021–FY2025 income statement, balance sheet and cash flow come from the 20-F XBRL.
   - 9M25 and 9M26 come from the 3Q26 6-K.
   - Every year ties: assets = liabilities + equity, cash-flow closing cash = balance-sheet cash, EBIT and net income match the reported figures.
3. **Drivers.** Nine regions, each modelled in its local currency:
   - Postpaid subscribers × ARPU is the engine.
   - Prepaid revenue is driven by revenue growth, not subscriber counts, because base cleanups distort the counts.
   - Broadband accesses × revenue per access drives fixed revenue.
   - Translation to MXN uses spot FX plus relative-PPP drift.
   - The 3Q26 calibration reproduces reported revenue exactly in every region.
4. **Projections.**
   - FCFF = EBITDA − lease payments − tax on EBIT − capex − ΔNWC.
   - A projected income statement, balance sheet and cash flow statement tie every year, and sources equal uses.
5. **WACC.**
   - US 10Y risk-free rate.
   - Damodaran EM telecom unlevered beta, relevered to the company's target leverage.
   - Revenue-weighted country risk premium.
   - Conversion to MXN with the inflation differential.
   - Cost of debt taken from actual FY25 interest expense divided by average debt.
6. **DCF.**
   - Mid-year discounting and a Gordon terminal value.
   - An explicit TV-share test.
   - An equity bridge that deducts net debt (re-marked to spot FX), the post-quarter dividend, the after-tax pension deficit and NCI, and adds back affiliates.
7. **Sensitivity.** Full WACC × g, WACC × margin and g × margin grids. Each cell is a closed-form recomputation of the whole DCF, so the grids are live formulas rather than data tables.
8. **Checks.** 60 integrity tests, all passing.

### Key judgement calls (explained in the memo)
- **Desktop (Brazil) acquisition: excluded from the base case.** It closed 1-Oct-2026, after the balance-sheet date, and represents <1% of EV. A toggle on Inputs includes it: +MXN 0.25/share.
- **Leases.** FCFF is after lease payments, so lease liabilities are excluded from net debt. The comps use both an IFRS 16 basis and an after-lease basis.
- **Transcript vs. filing.** The call transcript misstates several figures (for example Brazil postpaid adds of 343k vs 743k). The 6-K always wins.

## Reproduce
```bash
python3 -m venv .venv && .venv/bin/pip install openpyxl pypdf
```
```bash
.venv/bin/python scripts/build_model.py
```
```bash
scripts/recalc.sh model/AMX_Valuation_Model.xlsx model/
```
```bash
.venv/bin/python scripts/audit_model.py model/AMX_Valuation_Model.xlsx
```
```bash
.venv/bin/python scripts/scenarios.py /tmp/amx_scenarios
```
LibreOffice (`soffice`) must be on your PATH for the recalculation and scenario steps. Excel recalculates on open.

## Data policy
- Public data only: SEC EDGAR filings, company press releases, Damodaran Online datasets and public market quotes.
- No credentials, API keys or private data are in this repository.
- Third-party figures (market prices, a few peer balance-sheet items, Desktop's pre-deal financials) are labelled wherever they are used.

*Educational exercise. Not investment advice.*
