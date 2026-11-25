# Scenario reruns (full model recalculated for each line)

| Scenario | Value/share (MXN) | vs. price | TV % of EV | Rule output | Checks pass |
|---|---:|---:|---:|---|---|
| Base case | 25.54 | +6.5% | 67% | HOLD | True |
| FX: peso +10% vs. spot (stronger MXN) | 22.97 | -4.2% | 67% | HOLD | True |
| FX: peso −10% vs. spot (weaker MXN) | 28.67 | +19.5% | 67% | BUY | True |
| Desktop included (EBITDA in, full cost in net debt) | 25.79 | +7.5% | 67% | HOLD | True |
| Pension deficit deducted pre-tax | 24.62 | +2.7% | 67% | HOLD | True |
| Postpaid/prepaid ARPU ratio k = 2.0 everywhere | 25.30 | +5.5% | 67% | HOLD | True |
| Postpaid/prepaid ARPU ratio k = 4.0 everywhere | 25.72 | +7.3% | 67% | HOLD | True |
| Postpaid net adds 2027E = 50% of LTM (bear engine) | 24.99 | +4.2% | 67% | HOLD | True |
| Postpaid net adds 2027E = 100% of LTM (bull engine) | 25.90 | +8.0% | 67% | HOLD | True |
| No further margin expansion | 24.66 | +2.8% | 67% | HOLD | True |
| Capex 14.5% of revenue from 2028E (and terminal D&A) | 23.98 | +0.0% | 67% | HOLD | True |
| Terminal growth 2% (−1% real) | 22.77 | -5.0% | 64% | HOLD | True |
| Bear (operating): margin flat, net adds 50%, capex 14.5%, g 2.5% | 21.27 | -11.3% | 65% | HOLD | True |
| Bull: margin +0.3pp/yr, net adds 100%, g 3.5% | 28.05 | +17.0% | 69% | BUY | True |
