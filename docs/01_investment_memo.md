# América Móvil (BMV/NYSE: AMX): Investment Memo

**Recommendation: HOLD.** The DCF gives a **fair value of MXN 25.54 per share** (USD 26.23 per ADR), **+6.5%** above the MXN 23.98 close on 24-Nov-2026. **No price target is set.** The model's upside sits inside its own error band: under the ±15% rule fixed before the model was run, upside of 6.5% is a HOLD. A target above the market would be a number the model does not support.

| | |
|---|---|
| Price (AMXB, 24-Nov-26) / ADR | MXN 23.98 / USD 24.63 |
| Market cap / EV (ex-leases, DCF) | MXN 1,432 bn / MXN 2,130 bn |
| WACC (nominal MXN) / terminal growth | 10.85% / 3.0% |
| Terminal value share of EV | 67%. Passes the 70% test, but not by much |
| Sensitivity range (WACC ±1 pp × g ±1 pp) | MXN 19.6 to 35.5 |
| Operating bear / bull reruns | MXN 21.27 (−11%) / MXN 28.05 (+17%) |
| Comps cross-check (EV/EBITDA, all-peer median 4.6x) | MXN 19.9 (−17%) |

## 1. Thesis in three lines
1. **The engine works.** Postpaid net adds were 3.1 m in 3Q26 (Brazil 743k, Colombia 348k, Peru 229k, Mexico 144k), and the postpaid base grew 8.9% y/y. Postpaid service revenue grew 7.5% and EBITDA margin rose 0.5 pp to 40.8%. Prepaid is noise: the −5.0 m disconnections are base cleanups, and the model treats them that way.
2. **The balance sheet is a strength, not a catalyst.** Net debt/EBITDAaL is 1.21x, the low end of the CFO's 1.2–1.5x target. 9M26 FCF of MXN 75 bn funded capex, MXN 27.5 bn of distributions and MXN 34 bn of net debt reduction.
3. **The price already reflects this.** AMX trades at 5.2x run-rate EV/EBITDA against a 4.6x peer median, and the DCF leaves only 6.5% upside. Upside of 15% or more needs either WACC below 10.4% or about 1.2 pp of extra margin. Neither is in the base case.

## 2. How the value is built (all inputs sourced on the *Inputs* sheet)
- **Data.** FY2021–25 come from the 20-F XBRL on SEC EDGAR (CIK 0001129137). 9M26 and 3Q26 come from the 6-K furnished 22-Oct-2026. All five 3Q26 anchors in the brief match the 6-K exactly: revenue 250,634 (+7.6%), EBITDA 102,156 (40.8%, +0.5 pp), EBIT 57,113 (+13.9%), net income 24,282 (+7.0%). The Motley Fool transcript misstates several figures: Brazil postpaid adds 343k (6-K: 743k), Argentina broadband adds 63k (53k), and 9M capex "676 bn" (75.9 bn). **The 6-K figure is used every time.**
- **Drivers, by country, in local currency.** The model covers nine regions. Postpaid revenue = average postpaid subs × postpaid ARPU, with ARPU calibrated so 3Q26 revenue is matched exactly. Prepaid revenue grows at an input rate below inflation. Fixed revenue = broadband accesses × revenue per access. Margins start from 9M26 and add 0.2 pp a year for three years. 2026–27 capex follows the USD 7 bn guidance, then 13.5% of revenue.
- **FX, stated explicitly.** Each currency is translated at the 24-Nov-2026 spot rate (USD/MXN 19.47, against 17.84 at 30-Sep), then moved by the inflation differential versus Mexico (relative PPP, central-bank targets). Argentina is modelled in real terms (IAS 29). The weaker peso also raises the MXN value of the USD/EUR/GBP debt, so that debt is re-marked by **+MXN 21.0 bn**.
- **FCFF and the equity bridge.** FCFF = EBITDA − lease payments − tax on EBIT − capex − ΔNWC. Because FCFF already pays the leases, lease liabilities are kept out of net debt. The bridge from EV to equity deducts net debt 400.1 bn, the FX re-mark 21.0 bn, the November dividend 17.0 bn, the after-tax pension deficit 127.2 bn and NCI 65.6 bn, and adds back affiliates of 25.4 bn.
- **WACC of 10.85%, built rather than copied.** Rf is the US 10Y at 4.90%. The beta starts from Damodaran's EM telecom unlevered beta of 0.641 and is relevered to 0.778 at the target D/E of 30.5% (1.35x × LTM EBITDAaL). The mature ERP is 4.23%. The **revenue-weighted CRP is 3.14%** (Mexico 2.46%, Brazil 3.24%, Argentina 9.71%, Ecuador 12.95%). That gives a USD cost of equity of 11.3%, or 12.4% in MXN. Kd of 8.08% is the actual FY25 interest on financial debt divided by average gross debt. Weights are 23/77 debt/equity.

## 3. Sensitivity: what moves value (MXN per share; base 25.54)
| Shock | Value | Δ |
|---|---:|---:|
| WACC −100 bp / +100 bp | 30.57 / 21.64 | **+19.7% / −15.2%** |
| Terminal g +100 bp / −100 bp | 29.02 / 22.84 | +13.7% / −10.6% |
| EBITDA margin +200 bp / −200 bp (all years) | 28.84 / 22.23 | ±13.0% |
| Country risk +100 bp (WACC +77 bp) | 22.45 | −12.1% |
| Capex +100 bp of revenue | 23.16 | −9.3% |
| Peso +10% / −10% vs. spot | 22.97 / 28.67 | −10.0% / +12.3% |

The full WACC × g, WACC × margin and g × margin grids are on the *Sensitivity* sheet, and each cell recomputes the full DCF. **The discount rate moves value most.**

## 4. Desktop (Brazil): excluded from the base case, deliberately
Claro NXT bought **72.83% of Desktop S.A. for BRL 1,747 m (BRL 20.63/share)**, closing on 1-Oct-2026 (Claro *fato relevante*; AMX 6-K of 1-Oct-2026). A mandatory tender offer for the remaining shares is pending. On a 100% basis the EV is about BRL 3.7 bn (≈5.6x LTM EBITDA, using third-party Desktop financials), which is about MXN 12 bn, or **0.6% of AMX's EV**. The deal is excluded for four reasons:

- It closed after the 30-Sep balance sheet the bridge uses.
- None of its revenue, EBITDA or cash outflow is in any reported number yet.
- The tender offer and the final price adjustment are open.
- It is immaterial.

Including it with the toggle on *Inputs* adds only **+MXN 0.25/share (+1%)**. A deal priced near AMX's own multiple neither creates nor destroys much value. The CEO's "BRL 4 bn" remark on the call could not be reconciled to the filing, so the filing is used.

## 5. Key risks
- **FX.** About 63% of revenue is earned outside Mexico. A *stronger* peso lowers MXN value (peso +10% → MXN 22.97). A weaker peso raises it, but in practice it usually comes with higher Mexican rates, and the model holds WACC fixed in that case.
- **Country risk and rates.** The US 10Y is near a 25-year high and the Fed hiked 25 bp on 16-Sep. CRP +179 bp alone turns the call into a SELL.
- **Mexico competition and regulation.** The CEO flagged a Televisa/AT&T Mexico combination as a "strong competitor". Prepaid registration runs until 31-Dec-2026, so line suspensions could last into 1Q27.
- **Terminal-value dependence.** TV is 67% of EV: below the 70% red line, but two-thirds of the value still sits beyond 2031.
- **Pension.** The Telmex labour obligation (≈MXN 182 bn pre-tax) is a large debt-like claim. Deducting it pre-tax lowers value to MXN 24.62.
- **Model limits.** The postpaid/prepaid ARPU split is an assumption, though its effect is small (k = 2 to 4 gives MXN 25.30 to 25.72). Peer data mixes quarters for some names.

## 6. What would change the recommendation
- **To BUY (value ≥ MXN 27.58):**
    - WACC ≤ 10.4%, i.e. a 44 bp lower risk-free rate or CRP;
    - or about +1.2 pp of sustained EBITDA margin above plan;
    - or terminal growth ≥ 3.6%;
    - or the bull rerun (margin +0.3 pp a year, postpaid adds at 100% of LTM, g of 3.5%) → MXN 28.05.
- **To SELL (value ≤ MXN 20.38):**
    - WACC ≥ 12.2%, for example a 180 bp rise in country risk such as a Mexico downgrade cycle;
    - or EBITDA margin about 3.1 pp below plan;
    - or terminal growth below 1%;
    - or the operating bear case combined with a stronger peso.

*Educational exercise built from public filings. Not investment advice.*
