# Defense sheet: answers without opening the model

Base case: **MXN 25.54 per share vs. MXN 23.98 (+6.5%) → HOLD**. WACC 10.85%, g 3.0%, TV 67% of EV.

---

### (a) Why AMX, and why now?
- **Why AMX.**
  - It is the largest telecom operator in Latin America: 332 m wireless subscribers, 153 m of them postpaid, and 38 m broadband accesses.
  - The engine is the same in every country (postpaid subscribers + fibre broadband + ARPU), so one consolidated DCF is legitimate.
  - The balance sheet is conservative: 1.21x net debt/EBITDAaL.
  - You can test the drivers country by country against a single filing.
- **Why now.**
  - 3Q26 showed the engine accelerating: postpaid revenue +7.5%, margin +0.5 pp, EBIT +13.9%.
  - The quarter also contains three things the market has not digested:
    - the Mexican prepaid cleanup (−5.0 m disconnections);
    - a large FX swing (the peso is ~9% weaker since 30-Sep);
    - the Desktop deal (1-Oct-26).
  - That makes it a good moment to separate signal from noise.
- **The honest conclusion.** Separating them gives a fair value only 6.5% above the price. The work supports "own it at the right price", not "buy it now".

### (b) Which assumption moves value most (growth, margin or WACC), and by how much?
**WACC.** Each move is one variable at a time, around MXN 25.54:

| Variable moved | Change in value |
|---|---|
| WACC ±100 bp | **−15.2% / +19.7%** (MXN 21.64 / 30.57) |
| Terminal growth ±100 bp | +13.7% / −10.6% |
| EBITDA margin ±200 bp, every year | ±13.0% |
| Capex +100 bp of revenue | −9.3% |
| Near-term growth (postpaid net adds at 50% vs 100% of LTM) | only −2% / +1% |

Near-term growth barely matters: value is a mature cash-flow annuity, so the rate it is discounted at dominates.

### (c) How was the WACC built, and what happens if country risk rises 100 bp?
- **Build in USD.**
  - Rf: US 10Y at 4.90% (24-Nov-26).
  - Beta: Damodaran Jan-2026 emerging-market unlevered betas, averaging Telecom Wireless (0.668) and Telecom Services (0.614) to get 0.641. Relevered with Hamada at the **target** D/E of 30.5% to get **0.778**.
  - Target D/E comes from the CFO's leverage target: 1.35x (midpoint of 1.2–1.5x) × LTM EBITDAaL of MXN 324 bn = MXN 437 bn of net debt, against a MXN 1,432 bn market cap.
  - Mature-market ERP: 4.23% (Damodaran).
  - CRP: 3.14%, weighted by where revenue is earned in 9M26 (Mexico 36.6% × 2.46%, Brazil 19.0% × 3.24%, Colombia 9.0% × 2.85%, Argentina 5.3% × 9.71%, Ecuador 2.0% × 12.95%, …). It is added in full (λ = 1).
  - Ke (USD) = 4.90% + 0.778 × 4.23% + 3.14% = **11.3%**.
- **Convert to MXN.** Multiply by (1.03/1.02), the Banxico vs Fed inflation targets, to get **Ke = 12.4%**.
- **Cost of debt.** Kd = FY25 interest on financial debt / average gross debt = 44,162 / 546,247 = **8.08%** pre-tax, or 5.66% after 30% tax. This sits above a UST + 100 bp market check (6.9% in MXN terms), so it is conservative.
- **WACC.** 76.6% × 12.43% + 23.4% × 5.66% = **10.85%**.
- **CRP +100 bp.** Ke rises 101 bp, so WACC rises **77 bp to 11.62%**, and value falls **12.1% to MXN 22.45**. That is still a HOLD.
- **What flips it to SELL.** A CRP rise of about 180 bp: WACC of about 12.2% gives MXN 20.38.

### (d) What % of value is terminal value, and is that a problem?
- **67% of EV** (PV of TV MXN 1,426 bn out of MXN 2,130 bn).
- It is below the 70% red line, so the case is not "weak" by the mandate's test. It is close, though, and I say so.
- Why it is not worse:
  - Explicit FCFF is already sizeable (≈15% of revenue each year).
  - The terminal year is normalised: capex = D&A at 13.5% of revenue, and g of 3% means zero real growth.
- Cross-check: the implied exit multiple is **4.8x 2031E EBITDA**, *below* where AMX and its peers trade today (4.6x–5.2x). The TV is not built on a heroic multiple.
- The real risk inside the TV is the discount rate, not the cash flow.

### (e) What would flip the recommendation?
- **To BUY (≥ MXN 27.58, +15%):** any one of
  - WACC ≤ 10.4% (−44 bp);
  - margin +1.2 pp every year;
  - g ≥ 3.6%;
  - the bull bundle (margin +0.3 pp/yr, postpaid adds 100% of LTM, g 3.5%) → MXN 28.05;
  - a further ~10% peso depreciation → MXN 28.67, if rates do not rise with it.
- **To SELL (≤ MXN 20.38):** any one of
  - WACC ≥ 12.2% (CRP +180 bp);
  - margin −3.1 pp;
  - g < 1%;
  - the operating bear (MXN 21.27) combined with a stronger peso.
- **The market-implied number.** Capex at 14.5% of revenue instead of 13.5% takes value exactly to the share price. The market is implicitly pricing a heavier investment cycle than guidance.

### (f) Why does the DCF differ from comps, and which do I believe more?
- **The gap.** Comps (run-rate, IFRS 16 basis) put AMX at **5.2x EV/EBITDA vs a 4.6x peer median** (LatAm 4.35x; EV/EBITDAaL 5.9x vs 5.0x). Applying the median gives about **MXN 20 (−17%)**. The DCF says MXN 25.5.
- **Why they differ.**
  1. **Peers are not like-for-like.** The LatAm median is driven by Vivo and TIM: single-country Brazil, high payout, Brazil CRP. AMX is diversified across 20+ countries and earns ~12% of revenue in Europe (A1).
  2. **Leverage and quality.** AMX carries 1.2x leverage against 2.5x for Millicom and 4.5x for LILA, and has a long record of FCF and buybacks. The market pays a premium for that.
  3. **Definitions.** Peer EVs do not deduct pension deficits, but my DCF deducts MXN 127 bn for Telmex. On a like-for-like basis AMX's premium to peers is even larger than the screen shows.
  4. **Data quality.** Run-rate multiples use the latest quarter. For Entel, Orange and KPN some balance-sheet items are older or third-party. I would not hang a recommendation on a 0.5x multiple difference.
- **Which I believe more: the DCF.** It prices AMX's own cash flows, FX mix and balance sheet line by line, and I can defend every input.
- **What the comps add.** They tell me the market already pays up for AMX. Further upside needs execution beyond what peers deliver, which is exactly why the answer is **HOLD, not BUY**.

---

### Quick facts to have ready
(6-K page numbers = PDF page of the 40-page 3Q26 filing.)
| Item | Figure | Source |
|---|---|---|
| 3Q26 revenue / EBITDA / margin | MXN 250,634 m / 102,156 m / 40.8% | 6-K p.11 |
| 3Q26 EBIT / net income | MXN 57,113 m (+13.9%) / 24,282 m (+7.0%) | 6-K p.11–12 |
| Postpaid net adds 3Q26 | 3.1 m (BR 743k, CO 348k, PE 229k, MX 144k, AR 126k) | 6-K p.3, p.5 |
| Prepaid | −5.0 m (MX −2.9 m, CO −2.8 m, cleanups) | 6-K p.16, p.20 |
| Broadband net adds | 390k (BR 94k, CO 67k, AR 53k) | 6-K p.3, p.5 |
| Peso in 3Q26 | ≈ −2.1% vs USD (gave up Jul–Aug gains in Sep) | 6-K p.9 |
| Net debt ex-leases / leverage | MXN 400,081 m / 1.21x EBITDAaL | 6-K p.37, p.10 |
| 9M26 capex / FCF / distributions | MXN 75.9 bn / 75.0 bn / 27.5 bn | 6-K p.10, p.15 |
| Desktop | 72.83% for BRL 1,747 m (BRL 20.63/sh), closed 1-Oct-26 | Claro fato relevante |
| Shares | 59.7 bn; 20 per ADR | 6-K p.4 |
