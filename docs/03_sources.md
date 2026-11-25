# Sources (all accessed 24–25 November 2026 unless stated)

Primary sources are used for every AMX figure. Third-party data appears only for (i) market prices and FX, (ii) some peer balance-sheet items, and (iii) Desktop's pre-deal financials. Each one is labelled where it is used.

## A. América Móvil: primary filings (SEC EDGAR, CIK 0001129137)
| # | Document | Used for | Link |
|---|---|---|---|
| A1 | 3Q26 earnings report, Form 6-K furnished 22-Oct-2026 (40 pp.) | 3Q26/9M26 P&L, Sep-26 balance sheet, debt by currency, FCF, subscribers and revenue/EBITDA by country, FX tables | http://d18rn0p25nwr6d.cloudfront.net/CIK-0001129137/6c9b7502-b913-423d-8aa5-f86eb0d982f9.pdf (same document as accession 0001140361-26-040507) |
| A2 | Annual report FY2025, Form 20-F filed 28-Apr-2026 (accession 0001140361-26-017486), XBRL instance `ef20060642_20f_htm.xml` | FY2023–FY2025 income statement, balance sheet, cash flow, pension, leases, interest | https://www.sec.gov/Archives/edgar/data/1129137/000114036126017486/ |
| A3 | XBRL company facts API (20-F filings FY2021–FY2024) | FY2021–FY2022 figures and Dec-2020 cash | https://data.sec.gov/api/xbrl/companyfacts/CIK0001129137.json |
| A4 | Form 6-K, 1-Oct-2026 ("America Movil acquires Desktop in Brazil") | Desktop closing | https://www.sec.gov/Archives/edgar/data/1129137/000129281426004830/amx20261001_6k.htm |
| A5 | Claro Telecom Participações, *Fato Relevante* and notice to Desktop S.A., 1-Oct-2026 | Desktop price (BRL 1,747,135,751.94; BRL 20.63/share; 72.83%), escrow, tender offer | https://a.storyblok.com/f/238969/x/93f7f1500d/claro_participacoes_fato_relevante_e_notificacao_claro_nxt.pdf |
| A6 | 3Q26 earnings call transcript, 21-Oct-2026 (The Motley Fool) | Capex guidance (≈USD 7 bn 2026, similar 2027), leverage target 1.2–1.5x, November dividend ≈MXN 17 bn, competition commentary | https://www.fool.com/earnings/call-transcripts/2026/10/21/america-movil-amx-q3-2026-earnings-call-transcript/ |

**Transcript vs. 6-K discrepancies (6-K used every time):**
- Brazil postpaid net adds: 343k in the transcript vs 743k in the 6-K.
- Argentina broadband net adds: 63k vs 53k.
- Net debt/EBITDAaL at end-2025: 1.51x vs 1.52x.
- Labour obligation reduction: "15 bn" vs 13.0 bn.
- 9M capex: "676 bn" vs 75.9 bn.
- Net income: "MXN 24 million" vs MXN 24,282 million.

## B. Cost of capital
| # | Source | Used for | Link |
|---|---|---|---|
| B1 | Damodaran Online, *Country Default Spreads and Risk Premiums* (`ctryprem.xlsx`, update 5-Jan-2026) | Mature-market ERP 4.23%; country risk premiums; statutory tax rates | https://pages.stern.nyu.edu/~adamodar/pc/datasets/ctryprem.xlsx |
| B2 | Damodaran Online, *Betas by sector, Emerging Markets* (`betaemerg.xls`, 5-Jan-2026) | Unlevered betas: Telecom (Wireless) 0.6684, Telecom Services 0.6141 (cash-corrected) | https://pages.stern.nyu.edu/~adamodar/pc/datasets/betaemerg.xls |
| B3 | CBOE 10-year Treasury yield index (^TNX), via Yahoo Finance, close 24-Nov-2026 | Risk-free rate 4.904% | https://finance.yahoo.com/quote/%5ETNX/ |
| B4 | Banco de México (3% target), Federal Reserve (2%), BCB/CMN (3%), BanRep (3%), BCRP (2%), BCCh (3%), ECB (2%) | Long-run inflation for the MXN conversion and PPP FX drift | Central-bank websites |

## C. Market data (Yahoo Finance chart / quote-summary API)
| Item | Value | Timestamp |
|---|---|---|
| AMXB.MX | MXN 23.98 | close 24-Nov-2026 |
| AMX (NYSE ADR) | USD 24.63 | close 24-Nov-2026 |
| USD/MXN · EUR/USD · GBP/USD · USD/BRL | 19.4718 · 1.0949 · 1.3090 · 5.9028 | 25-Nov-2026 ~05:00 UTC |
| USD/COP · USD/PEN · USD/CLP · USD/ARS · USD/SEK | 3,606.99 · 3.474 · 1,032.81 · 1,592.0 · 10.2664 | 24/25-Nov-2026 |
| Peer prices | VIVT3 35.30, TIMS3 21.88, TIGO 79.12, ENTEL 5,435, LILA 9.50, TEF 3.404, ORA 15.285, VOD 102.9p, KPN 4.342, TELIA 50.5 | close 24-Nov-2026 |

A quarterly-series cross-check against AMX's own 6-K matched revenue, lease liabilities and share count exactly. Yahoo's EBITDA field did **not** match company definitions, so it is not used for any peer that publishes its own figure.

## D. Peer company results (company-reported figures, located via search summaries)
| Peer | Figures used | Source |
|---|---|---|
| Telefônica Brasil | 3Q26 revenue 16,023; EBITDA 7,136; EBITDA AL 5,689; net debt incl. leases ≈13.8 bn (BRL m) | 3Q26 release, 10-Nov-2026 (via teletime.com.br, finsee.ai) |
| TIM Brasil | 3Q26 revenue 7,069; normalised EBITDA 3,701; EBITDA-AL 2,917; total net debt 11,126 | 3Q26 release (via quartr.com, nasdaq.com transcript) |
| Millicom | 3Q26 revenue 2,276; adj. EBITDA 1,092 (USD m) | Q3 2026 earnings release, 5-Nov-2026 (via finviz/GlobeNewswire) |
| Entel | 1H26 EBITDA CLP 433,125 m; net financial debt CLP 1,335.7 bn (1Q26) | La Tercera / DF (1H26); Quartr (1Q26), **older data, flagged** |
| Liberty Latin America | 3Q26 revenue 1,117; Adjusted OIBDA 454 (USD m) | SEC 8-K exhibit 99.1, Q3 2026 |
| Telefónica | Jan–Sep 26 adj. EBITDA 8,846; 3Q revenue 8,433; net financial debt 25,186 (EUR m) | telefonica.com press release, Nov-2026 |
| Orange | 3Q26 revenue 12,191; EBITDAaL 4,140; net financial debt 35.7 bn (30-Jun-26) | Orange Q3 and H1 2026 releases |
| Vodafone | H1 FY27 revenue 21,853; adj. EBITDAaL 6.5 bn; net debt 33,473 (EUR m) | Vodafone H1 FY27 results, 10-Nov-2026 |
| KPN | 3Q26 adj. EBITDA AL 671; net debt 6,254 (1Q26) | kpn.com Q3 2026 release; Q1 2026 release |
| Telia | 3Q26 revenue 21,306; adj. EBITDA 8,870 (SEK m) | Interim report Jan–Sep 2026, 22-Oct-2026 |
| Desktop S.A. (pre-deal) | LTM revenue 1,271, EBITDA 655, net debt 1,285 (BRL m) | Yahoo Finance fundamentals, **third-party, unverified** |

**Missing or not verifiable to primary-source standard (stated, not invented):**
- AMX does not disclose postpaid vs. prepaid ARPU, so the ratio k is an explicit assumption. Sensitivity: k = 2 to 4 moves value by ±1%.
- Southern Cone mobile service revenue is not disclosed. It is estimated from the reported ARPU (CLP 5,645).
- NCI at 30-Sep-2026 is not in the 6-K. Dec-2025 book value is used.
- Peer 3Q26 net debt was not found for Entel, Orange or KPN. The latest available quarter is used and flagged.
- Precedent multiples for 2025–26 LatAm deals (Millicom–Telefónica Colombia/Ecuador/Uruguay; Telecom Argentina–Telefónica Argentina) could not be verified from primary documents. They are not used.
