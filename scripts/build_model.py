"""
Build the América Móvil (AMX) valuation workbook.

Every assumption lives on the Inputs sheet (blue font) with its source and a
one-line rationale. Every other sheet is formulas only (black = same sheet,
green = link to another sheet). Historical figures are reported data from
SEC filings (20-F XBRL and the 3Q26 6-K) and are entered as data on the
Historicals sheet.

Usage:  python scripts/build_model.py  ->  model/AMX_Valuation_Model.xlsx
"""
import json, os, datetime
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter as CL
from openpyxl.worksheet.properties import PageSetupProperties

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT = os.path.join(ROOT, "model", "AMX_Valuation_Model.xlsx")
HIST = json.load(open(os.path.join(ROOT, "data", "xbrl_hist_2021_2025.json")))

# ----------------------------------------------------------------- styles
BLUE = Font(name="Calibri", size=10, color="0000FF")
BLACK = Font(name="Calibri", size=10, color="000000")
GREEN = Font(name="Calibri", size=10, color="008000")
BOLD = Font(name="Calibri", size=10, bold=True)
TITLE = Font(name="Calibri", size=14, bold=True, color="FFFFFF")
HDR = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
NOTE = Font(name="Calibri", size=9, italic=True, color="595959")
RED = Font(name="Calibri", size=10, bold=True, color="C00000")
FILL_T = PatternFill("solid", fgColor="1F3864")
FILL_H = PatternFill("solid", fgColor="2F5597")
FILL_S = PatternFill("solid", fgColor="D9E1F2")
FILL_IN = PatternFill("solid", fgColor="FFF2CC")
FILL_KEY = PatternFill("solid", fgColor="E2EFDA")
THIN = Side(style="thin", color="A6A6A6")
BOX = Border(top=THIN, bottom=THIN, left=THIN, right=THIN)

F_M = '#,##0;(#,##0);"–"'          # millions
F_M1 = '#,##0.0;(#,##0.0);"–"'
F_K = '#,##0;(#,##0);"–"'           # thousands of subs
F_P = '0.0%;(0.0%);"–"'
F_P2 = '0.00%;(0.00%);"–"'
F_X = '0.00"x"'
F_N2 = '#,##0.00;(#,##0.00)'
F_N4 = '#,##0.0000'
F_D = 'dd-mmm-yyyy'
F_N3 = '#,##0.000'


def font_for(v):
    if isinstance(v, str) and v.startswith("="):
        return GREEN if "!" in v else BLACK
    return BLUE


class S:
    """Thin wrapper that tracks named rows on a sheet."""

    def __init__(self, wb, name, title, widths):
        self.ws = wb.create_sheet(name)
        self.name = name
        self.rows = {}
        self.r = 1
        self.ws.sheet_view.showGridLines = False
        for i, w in enumerate(widths, start=1):
            self.ws.column_dimensions[CL(i)].width = w
        self.ws["A1"] = title
        self.ws["A1"].font = TITLE
        for c in range(1, len(widths) + 1):
            self.ws.cell(1, c).fill = FILL_T
        self.r = 3

    def ref(self, name, col, absolute=True, ext=True):
        row = self.rows[name]
        c = CL(col) if isinstance(col, int) else col
        a = f"${c}${row}" if absolute else f"{c}{row}"
        return f"'{self.name}'!{a}" if ext else a

    def section(self, text, ncols):
        self.r += 1
        for c in range(1, ncols + 1):
            self.ws.cell(self.r, c).fill = FILL_H
        self.ws.cell(self.r, 1, text).font = HDR
        self.r += 1

    def header(self, labels, start_col=1, fill=FILL_S):
        for i, lab in enumerate(labels):
            cell = self.ws.cell(self.r, start_col + i, lab)
            cell.font = BOLD
            cell.fill = fill
            cell.alignment = Alignment(horizontal="center" if i else "left", wrap_text=True)
        self.r += 1

    def put(self, row, col, v, fmt=None, font=None, fill=None, bold=False):
        cell = self.ws.cell(row, col, v)
        f = font or font_for(v)
        if bold:
            f = Font(name=f.name, size=f.size, color=f.color, bold=True)
        cell.font = f
        if fmt:
            cell.number_format = fmt
        if fill:
            cell.fill = fill
        return cell

    def line(self, name, label, values, start_col=3, fmt=F_M, bold=False, note=None, note_col=None, unit=None):
        """values: list (one per column from start_col); None leaves blank."""
        if name:
            self.rows[name] = self.r
        self.put(self.r, 1, label, font=BOLD if bold else BLACK)
        if unit:
            self.put(self.r, 2, unit, font=NOTE)
        for i, v in enumerate(values):
            if v is None:
                continue
            self.put(self.r, start_col + i, v, fmt=fmt, bold=bold)
        if note:
            self.put(self.r, note_col or (start_col + len(values) + 1), note, font=NOTE)
        self.r += 1
        return self.r - 1

    def note(self, text, col=1):
        self.put(self.r, col, text, font=NOTE)
        self.r += 1


wb = Workbook()
wb.remove(wb.active)

# =========================================================================
# COVER (filled at the end, created first so it is the first tab)
# =========================================================================
cover = S(wb, "Cover", "América Móvil, S.A.B. de C.V. (BMV/NYSE: AMX) — DCF + comparables valuation", [44, 22, 18, 60])

# =========================================================================
# INPUTS
# =========================================================================
inp = S(wb, "Inputs", "Inputs — every assumption, its source and its rationale (blue = hard-coded input)",
        [52, 16, 16, 16, 16, 16, 16, 16, 16, 16, 16, 62, 70])
I = {}  # name -> absolute external reference


def scalar(name, label, value, unit, source, why, fmt=F_M):
    r = inp.r
    inp.rows[name] = r
    inp.put(r, 1, label)
    c = inp.put(r, 2, value, fmt=fmt, fill=FILL_IN)
    inp.put(r, 3, unit, font=NOTE)
    inp.put(r, 12, source, font=NOTE)
    inp.put(r, 13, why, font=NOTE)
    I[name] = f"Inputs!$B${r}"
    inp.r += 1


inp.put(inp.r, 1, "Colour code: blue = hard-coded input; black = formula; green = link to another sheet. "
        "All money figures MXN millions unless stated. Subscribers in thousands. 6-K page numbers = PDF page of the 40-page 3Q26 filing.", font=NOTE)
inp.r += 1
inp.header(["Input", "Value", "Unit"] + [""] * 8 + ["Source (date accessed 24/25-Nov-2026 unless stated)", "Rationale / justification"])

inp.section("1. Market data and valuation date", 13)
scalar("val_date", "Valuation date", datetime.date(2026, 11, 24), "date", "Last close before model date", "Prices and FX taken at this close", F_D)
scalar("price", "AMX Series B share price (BMV: AMXB)", 23.98, "MXN/share", "Yahoo Finance chart API, AMXB.MX close 24-Nov-2026", "Market price the DCF is compared against; cross-checks with ADR: 24.63/20×19.4718 = 23.98")
inp.ws.cell(inp.rows["price"], 2).number_format = F_N2
scalar("adr_price", "AMX ADR price (NYSE: AMX)", 24.63, "USD/ADR", "Yahoo Finance chart API, AMX close 24-Nov-2026", "Cross-check of local price", F_N2)
scalar("shares", "Shares outstanding, end of period", 59700, "millions", "AMX 3Q26 6-K p.4 ('Shares Outstanding End of Period 59.7 bn')", "Latest reported count; buybacks after 30-Sep ignored (conservative, tiny)")
scalar("shares_per_adr", "Shares per ADR", 20, "shares", "AMX 3Q26 6-K p.4 footnote (2)", "ADR conversion", "0")
scalar("usdmxn", "USD/MXN spot", 19.4718, "MXN per USD", "Yahoo Finance MXN=X, 25-Nov-2026 05:20 UTC", "Spot FX used to translate foreign cash flows and to re-mark FX debt", F_N4)
scalar("eurusd", "EUR/USD spot", 1.0949, "USD per EUR", "Yahoo Finance EURUSD=X, 25-Nov-2026", "Cross rate for EUR", F_N4)
scalar("gbpusd", "GBP/USD spot", 1.3090, "USD per GBP", "Yahoo Finance GBPUSD=X, 25-Nov-2026", "Cross rate for GBP debt re-mark", F_N4)
scalar("usdbrl", "USD/BRL spot", 5.9028, "BRL per USD", "Yahoo Finance BRL=X, 25-Nov-2026", "Cross rate for BRL", F_N4)
scalar("usdcop", "USD/COP spot", 3606.99, "COP per USD", "Yahoo Finance COP=X, 25-Nov-2026", "Cross rate for COP", F_N2)
scalar("usdpen", "USD/PEN spot", 3.474, "PEN per USD", "Yahoo Finance PEN=X, 25-Nov-2026", "Cross rate for PEN", F_N4)
scalar("usdclp", "USD/CLP spot", 1032.81, "CLP per USD", "Yahoo Finance CLP=X, 24-Nov-2026", "Cross rate for CLP", F_N2)
scalar("usdars", "USD/ARS spot", 1592.0, "ARS per USD", "Yahoo Finance ARS=X, 25-Nov-2026", "Cross rate for ARS (official rate)", F_N2)
scalar("fx_shock", "FX scenario: extra MXN appreciation (+) / depreciation (−) vs spot on all non-MXN cash flows", 0.0, "%", "Scenario lever (base = 0)", "Lets the user flex the FX assumption explicitly; 0 = translate at spot then PPP drift", F_P)
scalar("ust10", "US 10-year Treasury yield (risk-free, USD)", 0.04904, "%", "CBOE ^TNX via Yahoo Finance, close 24-Nov-2026 (4.904%)", "Long-dated USD risk-free; WACC built in USD then converted to MXN", F_P2)

inp.section("2. Balance sheet items for the equity bridge (30-Sep-2026 unless stated)", 13)
scalar("debt_total", "Total financial debt, excl. leases", 483910, "MXN m", "AMX 3Q26 6-K p.37 debt table ('Total Debt 483,910')", "Reported gross debt at 30-Sep-26")
scalar("cash_sec", "Cash, marketable securities & ST investments", 83829, "MXN m", "AMX 3Q26 6-K p.37 / p.13", "Company net-debt definition (includes securities available for sale)")
scalar("debt_usd", "of which USD-denominated debt", 9001, "USD m", "AMX 3Q26 6-K p.37 (in currency millions)", "Re-marked to spot below; table excludes hedges (6-K note 1)")
scalar("debt_eur", "of which EUR-denominated debt", 3432, "EUR m", "AMX 3Q26 6-K p.37", "Re-marked to spot")
scalar("debt_gbp", "of which GBP-denominated debt", 2200, "GBP m", "AMX 3Q26 6-K p.37", "Re-marked to spot")
scalar("debt_brl", "of which BRL-denominated debt", 10000, "BRL m", "AMX 3Q26 6-K p.37", "Re-marked to spot")
scalar("usdmxn_sep", "USD/MXN at 30-Sep-2026", 17.8413, "MXN per USD", "AMX 3Q26 6-K p.35 (end of period)", "Rate at which the 6-K debt table is translated", F_N4)
scalar("eurusd_sep", "EUR per USD at 30-Sep-2026", 0.8826, "EUR per USD", "AMX 3Q26 6-K p.35 (end of period)", "Rate used in the 6-K", F_N4)
scalar("gbpusd_sep", "GBP/USD at 30-Sep-2026", 1.3266, "USD per GBP", "Implied by the 6-K debt table total (GBP 2,200m = MXN 52,072m); Yahoo GBPUSD=X 30-Sep close 1.323–1.326 corroborates", "GBP rate not printed in the 6-K; implied value reproduces the reported total", F_N4)
scalar("brlmxn_sep", "BRL per MXN at 30-Sep-2026", 0.2904, "BRL per MXN", "AMX 3Q26 6-K p.34 (end of period)", "Rate used in the 6-K", F_N4)
scalar("lease_liab", "Lease liabilities (current + non-current)", 218991, "MXN m", "AMX 3Q26 6-K p.13 (39,012 + 179,979)", "Excluded from net debt in the DCF because FCFF deducts lease payments; used in comps EV (IFRS 16 basis)")
scalar("div_post", "Dividend paid after 30-Sep (Nov-2026 instalment)", 17000, "MXN m", "AMX 3Q26 call, 21-Oct-2026 (CFO: ~MXN 17 bn due in 2–3 weeks)", "Cash leaves after the balance sheet date and the share now trades ex-dividend, so it is deducted")
scalar("pension_dec25", "Net defined-benefit (labour) obligation, 31-Dec-2025", 194736, "MXN m", "AMX FY2025 20-F XBRL (ifrs-full:LiabilityAssetOfDefinedBenefitPlans)", "Debt-like claim: funding payments sit outside FCFF")
scalar("pension_paydown", "Reduction in labour obligations, 9M26", 13013, "MXN m", "AMX 3Q26 6-K p.15 ('Labor Obligations 13,013')", "Rolls the deficit forward to 30-Sep-26 (estimate: ignores 9M accruals/remeasurement)")
scalar("pension_tax", "Tax rate applied to pension deficit (0 = pre-tax)", 0.30, "%", "Mexico statutory rate (Damodaran ctryprem.xlsx 'Country Tax Rates', Jan-2026)", "Contributions are deductible, so the economic claim is after tax; set 0 to deduct pre-tax", F_P)
scalar("nci", "Non-controlling interests (book)", 65602, "MXN m", "AMX FY2025 20-F XBRL (NoncontrollingInterests 31-Dec-25); not disclosed in the 3Q26 6-K", "Mostly the 39% free float of A1 Telekom Austria; book is above A1 minority market value (~MXN 48 bn) so conservative")
scalar("affiliates", "Investments in affiliates and other investments", 25375, "MXN m", "AMX 3Q26 6-K p.13", "Non-consolidated assets not in FCFF")

inp.section("3. Desktop (Brazil) acquisition — EXCLUDED from base case (toggle)", 13)
scalar("desk_on", "Include Desktop? (1 = yes, 0 = no)", 0, "switch", "Decision: see memo §4", "Closed 1-Oct-2026, after the 30-Sep balance sheet; not in any reported number; <1% of EV; tender offer pending", "0")
scalar("desk_px", "Price per Desktop share paid", 20.63, "BRL", "Claro Telecom Participações fato relevante, 1-Oct-2026, item (c)", "Same price assumed for the mandatory tender offer (art. 254-A)", F_N2)
scalar("desk_shares", "Desktop total shares", 116.276, "millions", "Fato relevante: 84,684,273 shares = 72.83% → 116.28 m total", "Derived from the filing", F_N3)
scalar("desk_netdebt", "Desktop net debt (incl. leases), 30-Jun-2026", 1285, "BRL m", "Yahoo Finance fundamentals (DESK3.SA): debt 1,777 − cash 492 — third-party, unverified", "Only used if toggle = 1")
scalar("desk_rev", "Desktop LTM revenue (to Jun-26)", 1271, "BRL m", "Yahoo Finance quarterly revenue DESK3.SA (3Q25–2Q26) — third-party", "Only used if toggle = 1")
scalar("desk_ebitda", "Desktop LTM EBITDA (to Jun-26)", 655, "BRL m", "Yahoo Finance quarterly EBITDA DESK3.SA (3Q25–2Q26) — third-party", "Only used if toggle = 1")

inp.section("4. Consolidated operating assumptions", 13)
scalar("rev_3q", "Consolidated revenue 3Q26", 250634, "MXN m", "AMX 3Q26 6-K p.11", "Calibrates eliminations vs. sum of regions")
scalar("ebitda_3q", "Consolidated EBITDA 3Q26", 102156, "MXN m", "AMX 3Q26 6-K p.11", "Calibrates corporate/elimination EBITDA")
scalar("rev_9m", "Consolidated revenue 9M26", 728549, "MXN m", "AMX 3Q26 6-K p.11", "2026E = 9M actual + modelled 4Q")
scalar("ebitda_9m", "Consolidated EBITDA 9M26", 292568, "MXN m", "AMX 3Q26 6-K p.11", "2026E = 9M actual + modelled 4Q")
scalar("lease_pct", "Lease payments (principal + interest), % of revenue", 0.0545, "%", "FY25 51,586/943,638 = 5.47% (20-F CF); 9M26 39,639/728,549 = 5.44% (6-K p.36)", "Stable for two periods; held flat", F_P2)
scalar("da_2026", "D&A on owned assets (ex right-of-use), % of revenue, 2026E", 0.150, "%", "FY25: (180,805 − 39,226 RoU dep.)/943,638 = 15.0% (20-F XBRL)", "Starting point; glides to the terminal level", F_P)
scalar("da_term", "D&A on owned assets, % of revenue, 2031E and terminal", 0.135, "%", "Set equal to long-run capex intensity", "In steady state with ~0% real growth capex ≈ D&A", F_P)
scalar("capex_usd_26", "Capex 2026E", 7000, "USD m", "AMX 3Q26 call, CEO: ~USD 7 bn for 2026 (±5%)", "Company guidance")
scalar("capex_usd_27", "Capex 2027E", 7000, "USD m", "AMX 3Q26 call, CEO: 2027 'roughly the same'", "Company guidance")
scalar("capex_pct", "Capex 2028E onward and terminal, % of revenue", 0.135, "%", "FY25 actual 13.9% (PP&E+intangibles, 20-F CF); 2027E guidance ≈13%", "Midpoint of recent actual and guided intensity", F_P)
scalar("nwc_pct", "Net working capital (receivables + inventories − payables), % of revenue", 0.109, "%", "FY25: (241,381 + 28,307 − 167,125)/943,638 = 10.9% (20-F XBRL)", "Latest year; ratio rose from 7.8% (FY23) so no further release assumed", F_P)
scalar("margin_drift", "EBITDA margin expansion per year, 2027E–2029E", 0.002, "pp/yr", "3Q26 +0.5 pp y/y; 9M26 +0.5 pp y/y (6-K p.11)", "Less than half the recent pace, then flat from 2030E", F_P2)
scalar("drift_years", "Number of years margin expansion applies", 3, "years", "Assumption", "Efficiency gains assumed to run out", "0")
scalar("post_2027", "Postpaid net adds 2027E as % of LTM net adds", 0.80, "%", "LTM = Sep-25→Sep-26 change in postpaid base (6-K country tables)", "Haircut: Brazil/Colombia promotional intensity, Mexico competition (Televisa/AT&T)", F_P)
scalar("post_fade", "Annual fade in postpaid net adds after 2027E", 0.20, "%", "Assumption", "Penetration of postpaid rises, so incremental adds shrink", F_P)
scalar("bb_2027", "Broadband net adds 2027E as % of LTM net adds", 0.90, "%", "LTM = Sep-25→Sep-26 change in broadband base (6-K)", "Fibre build continues (USD 7 bn capex)", F_P)
scalar("bb_fade", "Annual fade in broadband net adds after 2027E", 0.15, "%", "Assumption", "Footprint saturation", F_P)
scalar("nci_pct", "NCI share of net income", 0.0573, "%", "9M26: 4,376 / 76,391 (6-K p.11–12)", "Used only in the projected P&L/BS", F_P2)
scalar("payout", "Dividends, % of net income attributable", 0.40, "%", "FY25: 33,160 / 82,819 = 40% (20-F)", "Projected statements only; does not affect FCFF", F_P)
scalar("buyback", "Buybacks, % of net income attributable", 0.15, "%", "FY25: 11,944 / 82,819 = 14% (20-F)", "Projected statements only", F_P)
scalar("cash_yield", "Interest income yield on opening cash", 0.07, "%", "Assumption (Banxico policy rate area; FY25 interest income 9,127)", "Projected P&L only", F_P)
scalar("months_y", "Months per year", 12, "months", "Constant", "Kept as an input so no number is hard-coded in formulas", "0")
scalar("months_q", "Months per quarter", 3, "months", "Constant", "Idem", "0")
scalar("q_per_y", "Quarters per year", 4, "quarters", "Constant", "Idem", "0")
scalar("units_k", "Subscribers reported in thousands (k → units)", 1000, "x", "Constant", "Revenue in LC millions ÷ subs in thousands needs ×1,000 to give LC per sub", "0")
scalar("ebitdaal_9m", "EBITDAaL 9M26 (company definition)", 241238, "MXN m", "AMX 3Q26 6-K p.15", "Comps run-rate (×4/3)")

inp.section("5. WACC inputs", 13)
scalar("erp", "Mature-market equity risk premium", 0.0423, "%", "Damodaran Online, ctryprem.xlsx 'ERPs by country', update 5-Jan-2026 (implied S&P ERP 4.46% less US default spread)", "Latest Damodaran update available; implied, not historical", F_P2)
scalar("bu_wireless", "Unlevered beta, Emerging-market 'Telecom (Wireless)', cash-corrected", 0.6684, "x", "Damodaran Online betaemerg.xls (5-Jan-2026), 68 firms", "AMX is >60% mobile service revenue", F_N4)
scalar("bu_services", "Unlevered beta, Emerging-market 'Telecom. Services', cash-corrected", 0.6141, "x", "Damodaran Online betaemerg.xls (5-Jan-2026), 142 firms", "Fixed-line/integrated reference", F_N4)
scalar("target_lev", "Target net debt / EBITDAaL", 1.35, "x", "AMX 3Q26 call, CFO: target range 1.2x–1.5x; actual 1.21x (6-K p.3, p.10)", "Midpoint of the stated range sets target capital structure", F_X)
scalar("ebitdaal_ltm", "LTM EBITDAaL to Sep-26", 323600, "MXN m", "9M26 241,238 (6-K p.15) + FY25 (372,209 EBITDA − 51,586 lease pmts, 20-F) − 9M25 238,261 (6-K)", "FY25 EBITDAaL approximated as EBITDA − lease payments; check: same method gives 9M25 238,570 vs 238,261 reported")
scalar("tax_marg", "Marginal tax rate for debt shield", 0.30, "%", "Mexico statutory rate (Damodaran tax table, Jan-2026)", "Parent (Mexico) is the main issuer of debt", F_P)
scalar("infl_us", "Long-run US inflation", 0.02, "%", "Federal Reserve 2% target", "Converts USD cost of capital to MXN", F_P)
scalar("int_debt_25", "FY25 interest expense on financial debt", 44162, "MXN m", "20-F XBRL: InterestExpense 60,319 − InterestExpenseOnLeaseLiabilities 16,157", "Cost of debt numerator (ex-leases, consistent with net debt)")
scalar("debt_24", "Gross financial debt 31-Dec-2024", 567586, "MXN m", "20-F XBRL Borrowings", "Average-debt denominator")
scalar("debt_25", "Gross financial debt 31-Dec-2025", 524907, "MXN m", "20-F XBRL Borrowings", "Average-debt denominator")
scalar("kd_spread", "Cross-check only: assumed AMX USD credit spread over UST", 0.010, "%", "Assumption (investment-grade issuer); not used in WACC", "Market-based sanity check of the effective rate", F_P2)
scalar("crp_lambda", "Exposure of AMX to country risk (lambda)", 1.0, "x", "Damodaran 'bludgeon' approach (λ = 1)", "Conservative: full CRP added to cost of equity", F_N2)

inp.section("6. DCF settings and recommendation rule", 13)
scalar("g", "Terminal growth rate (nominal MXN)", 0.030, "%", "Banxico 3% inflation target → ~0% real growth", "Mature telecom; below MXN nominal GDP growth", F_P)
scalar("stub", "Fraction of 2026E FCFF in the stub (4Q26)", 0.25, "x", "Net debt is at 30-Sep-26, so cash generated from 1-Oct belongs to the DCF", "One quarter of the year", F_N2)
scalar("stub_t", "Discount time of the stub cash flow", 0.0, "years", "4Q26 mid-point (15-Nov) ≈ valuation date", "Rounded to zero", F_N2)
scalar("fy1_end", "Year-end of the base year", datetime.date(2026, 12, 31), "date", "Calendar", "Timing anchor", F_D)
scalar("mid", "Mid-year convention offset", 0.5, "years", "Standard", "Cash flows arrive on average mid-year", F_N2)
scalar("days", "Days per year", 365, "days", "Constant", "Time fractions", "0")
scalar("tv_limit", "Terminal value share flagged as weak above", 0.70, "%", "Mandate", "If TV > 70% of EV the case is weak", F_P)
scalar("buy_th", "Buy if upside ≥", 0.15, "%", "House rule (set before running the model)", "Must exceed model error band", F_P)
scalar("sell_th", "Sell if upside ≤", -0.15, "%", "House rule (set before running the model)", "Symmetric", F_P)
scalar("step_w", "Sensitivity step: WACC", 0.005, "%", "Choice", "", F_P2)
scalar("step_g", "Sensitivity step: terminal growth", 0.005, "%", "Choice", "", F_P2)
scalar("step_m", "Sensitivity step: EBITDA margin (all years)", 0.01, "pp", "Choice", "", F_P2)
scalar("crp_shock", "Country-risk shock for defence question (c)", 0.01, "%", "Mandate (+100 bp)", "", F_P2)
scalar("pct_lo", "Comps: lower percentile", 0.25, "%", "Choice", "Interquartile range of peer multiples", F_P)
scalar("pct_hi", "Comps: upper percentile", 0.75, "%", "Choice", "", F_P)
scalar("sw_w", "Swing size: WACC", 0.01, "%", "Choice", "One-variable swing table", F_P2)
scalar("sw_g", "Swing size: terminal growth", 0.01, "%", "Choice", "", F_P2)
scalar("sw_m", "Swing size: EBITDA margin", 0.02, "pp", "Choice", "", F_P2)
scalar("sw_c", "Swing size: capex % revenue", 0.01, "pp", "Choice", "", F_P2)

# ---------------- regional matrix
REG = ["Mexico", "Brazil", "Colombia", "Peru", "Ecuador", "Argentina", "Southern Cone", "C. America & Carib.", "Austria & E. Europe"]
CCY = ["MXN", "BRL", "COP", "PEN", "USD", "ARS (IAS 29)", "CLP", "USD", "EUR"]
NR = len(REG)
RC = {}  # (name, i) -> ref


def rrow(name, label, vals, unit, source, why, fmt=F_M):
    r = inp.r
    inp.rows[name] = r
    inp.put(r, 1, label)
    inp.put(r, 11, unit, font=NOTE)
    for i, v in enumerate(vals):
        inp.put(r, 2 + i, v, fmt=fmt, fill=FILL_IN if not (isinstance(v, str) and v.startswith("=")) else None)
        RC[(name, i)] = f"Inputs!${CL(2 + i)}${r}"
    inp.put(r, 12, source, font=NOTE)
    inp.put(r, 13, why, font=NOTE)
    inp.r += 1


inp.section("7. Regional base data (3Q26 6-K, local currency millions; subscribers thousands) and regional drivers", 13)
inp.header(["Item"] + REG + ["Unit", "Source", "Rationale"])
rrow("ccy", "Reporting currency", CCY, "", "AMX 3Q26 6-K country pages", "Colombia converted from COP bn to COP m; Argentina in constant ARS (IAS 29)", None)
P6 = "AMX 3Q26 6-K country subscriber tables (pp.17–33)"
rrow("post_sep26", "Postpaid subscribers, Sep-26", [16254, 61934, 12618, 7623, 2403, 10573, 5446, 6162, 17481], "k", P6, "Austria: 29,818 less 12,337 A1 Digital M2M units (near-zero ARPU) so the engine is not inflated by M2M", F_K)
rrow("pre_sep26", "Prepaid subscribers, Sep-26", [65527, 30883, 25900, 5953, 7430, 16146, 3347, 20600, 3737], "k", P6, "Memo only — prepaid revenue is modelled on revenue growth, not subs", F_K)
rrow("tot_jun26", "Total wireless subscribers, Jun-26 (ex-M2M for Austria)", [84531, 92264, 40943, 13321, 9687, 26401, 8642, 26426, 20841], "k", "AMX 3Q26 6-K p.6 (Jun-26 column); Austria = Sep ex-M2M 21,218 − 377 non-M2M Q3 net adds (p.32)", "Average base for ARPU calibration", F_K)
rrow("post_sep25", "Postpaid subscribers, Sep-25", [15784, 57774, 11485, 6873, 2324, 10078, 5016, 5655, 16937], "k", P6, "LTM net adds; Austria ex 8,479 M2M", F_K)
rrow("bb_jun26", "Broadband accesses, Jun-26", [12345, 10841, 3704, 1189, 458, 1841, 1554, 2980, 2923], "k", "AMX 3Q26 6-K p.8", "", F_K)
rrow("bb_sep26", "Broadband accesses, Sep-26", [12383, 10935, 3771, 1216, 474, 1894, 1555, 3040, 2956], "k", "AMX 3Q26 6-K p.8", "", F_K)
rrow("bb_sep25", "Broadband accesses, Sep-25", [11816, 10530, 3508, 1075, 407, 1665, 1558, 2776, 2852], "k", "AMX 3Q26 6-K p.8", "", F_K)
rrow("q4_post", "4Q26E postpaid net adds (assumption = 3Q26 run-rate)", [144, 743, 348, 229, 20, 126, 108, 121, 136], "k", "3Q26 disclosed net adds (6-K p.3, p.5); Ecuador, S. Cone, Austria not disclosed → LTM/4", "Quarter-on-quarter run-rate", F_K)
rrow("q4_pre", "4Q26E prepaid net adds (memo)", [-2900, -151, -500, -70, 19, -212, -106, 237, -34], "k", "Mexico: repeat of 3Q26 −2.9 m (registration deadline 31-Dec-26, 6-K p.16); Colombia: tail of cleanup; others LTM/4", "Memo only (no revenue impact by construction)", F_K)
rrow("q4_bb", "4Q26E broadband net adds (assumption = 3Q26 run-rate)", [37, 94, 67, 28, 17, 53, 1, 60, 33], "k", "AMX 3Q26 6-K country pages", "Run-rate", F_K)
rrow("mob_svc_3q", "Mobile service revenue 3Q26", [49289, 7656, 2176000, 1078, 205, 838407, None, 770, 677], "LC m", "AMX 3Q26 6-K country tables", "Southern Cone not disclosed → formula below", F_M)
r_sc = inp.rows["mob_svc_3q"]
inp.put(r_sc, 2 + 6, f"=5645*AVERAGE({CL(8)}{inp.rows['tot_jun26']},{CL(8)}{inp.rows['post_sep26']}+{CL(8)}{inp.rows['pre_sep26']})*$B${inp.rows['months_q']}/$B${inp.rows['units_k']}", fmt=F_M)
inp.ws.cell(r_sc, 8).font = BLUE
inp.put(r_sc, 13, "Southern Cone mobile service revenue = reported ARPU CLP 5,645 (6-K p.29) × average subs × 3 months (input-derived estimate)", font=NOTE)
rrow("svc_3q", "Total service revenue 3Q26", [71773, 12962, 3459000, 1438, 237, 1026220, 305116, 1170, 1205], "LC m", "AMX 3Q26 6-K country tables", "Fixed service = total service − mobile service", F_M)
rrow("rev_3q", "Total revenue 3Q26", [90295, 13804, 4441000, 1832, 268, 1125915, 368520, 1318, 1466], "LC m", "AMX 3Q26 6-K country tables", "", F_M)
rrow("ebitda_3q", "EBITDA 3Q26", [37486, 6174, 1907000, 730, 137, 502191, 88844, 602, 575], "LC m", "AMX 3Q26 6-K country tables", "", F_M)
rrow("rev_9m", "Total revenue 9M26", [263473, 40576, 13061000, 5461, 809, 3210189, 1116690, 3860, 4260], "LC m", "AMX 3Q26 6-K country tables", "", F_M)
rrow("ebitda_9m", "EBITDA 9M26", [109199, 17950, 5509000, 2169, 408, 1365370, 299555, 1731, 1618], "LC m", "AMX 3Q26 6-K country tables", "9M margin is the base margin (smooths quarterly one-offs, e.g. S. Cone)", F_M)
rrow("k_ratio", "Postpaid ARPU / prepaid ARPU", [3.0, 2.5, 3.0, 3.0, 3.0, 3.0, 3.0, 3.0, 3.0], "x", "ASSUMPTION — AMX does not disclose ARPU by segment", "Splits calibrated blended ARPU; Brazil lower (cheap 'Controle' plans). Only affects the mix effect, since 3Q26 revenue is matched exactly", F_N2)
rrow("g_post_arpu", "Postpaid ARPU growth p.a. (nominal LC), 2027E+", [0.030, 0.012, 0.019, 0.0, 0.01, 0.0, 0.027, 0.01, 0.01], "%", "Implied 3Q26: postpaid rev growth (6-K) − postpaid sub growth: MX 8.0−3.0, BR 8.5−7.2, CO 12−9.9, PE 7.5−10.9, CL 11.5−8.6", "Min(implied, local inflation target), floored at 0; Ecuador/CAC/Austria not disclosed → half of 2% target; Argentina real", F_P)
rrow("g_prepaid", "Prepaid service revenue growth p.a. (nominal LC), 2027E+", [0.02, 0.0, 0.03, -0.02, 0.02, 0.0, -0.03, 0.03, 0.0], "%", "3Q26 commentary (6-K): MX +3.5% decelerating (registration), BR decelerating, PE declining, EC +5.6%, CAC accelerating", "Prepaid is noise: grows below inflation as users migrate to postpaid", F_P)
rrow("g_fixed_arpu", "Fixed revenue per broadband access growth, 2027E (fades linearly to 0% by 2031E)", [-0.002, -0.010, -0.040, -0.041, -0.041, -0.069, -0.009, -0.026, 0.010], "%", "3Q26 fixed service revenue growth less broadband access growth (6-K country pages)", "Captures erosion of voice/Pay-TV inside fixed revenue", F_P)
rrow("infl", "Long-run inflation (for PPP FX drift)", [0.03, 0.03, 0.03, 0.02, 0.02, 0.03, 0.03, 0.02, 0.02], "%", "Central-bank targets: Banxico 3%, BCB 3%, BanRep 3%, BCRP 2%, Fed 2% (USD economies), BCCh 3%, ECB 2%", "Argentina modelled in real terms (= Mexico) — IAS 29 constant pesos; no ARS forecast is defensible", F_P)
rrow("fx_3q", "FX 3Q26 average (LC per MXN)", [1, 0.2978, 185.1585, 0.1963, 0.0580, 85.0274, 54.0607, 0.0580, 0.0504], "LC/MXN", "AMX 3Q26 6-K p.34; Argentina at closing rate (IAS 29 translation)", "Calibration of eliminations", F_N4)
rrow("fx_9m", "FX 9M26 average (LC per MXN)", [1, 0.2961, 201.2227, 0.1957, 0.0575, 85.0274, 52.0842, 0.0575, 0.0495], "LC/MXN", "AMX 3Q26 6-K p.34; Argentina at closing rate", "Memo: regional 9M in MXN", F_N4)
spot = [f"=1", f"={I['usdbrl']}/{I['usdmxn']}", f"={I['usdcop']}/{I['usdmxn']}", f"={I['usdpen']}/{I['usdmxn']}", f"=1/{I['usdmxn']}",
        f"={I['usdars']}/{I['usdmxn']}", f"={I['usdclp']}/{I['usdmxn']}", f"=1/{I['usdmxn']}", f"=1/({I['usdmxn']}*{I['eurusd']})"]
rrow("fx_spot", "FX spot (LC per MXN), valuation date", spot, "LC/MXN", "Derived from section 1 spot quotes", "Translation rate for 4Q26E and the anchor for PPP drift", F_N4)
rrow("crp", "Country risk premium (Damodaran, rating-based)", [0.02465, 0.03241, 0.02845, 0.02069, 0.12949, 0.09708, 0.01096, 0.03568, 0.01311], "%", "Damodaran ctryprem.xlsx (Jan-2026): MX 2.46, BR 3.24, CO 2.85, PE 2.07, EC 12.95, AR 9.71, CL 1.10; CAC = avg Guatemala 3.24/Dom. Rep. 3.90; A1 = 50% Austria 0.36 + 50% avg(BG 2.07, HR 1.55, RS 3.90, SI 1.55)", "Operations-weighted CRP (where revenue is earned, not where listed)", F_P2)
rrow("tax", "Statutory corporate tax rate", [0.30, 0.34, 0.35, 0.295, 0.25, 0.35, 0.27, 0.28, 0.23], "%", "Damodaran ctryprem.xlsx 'Country Tax Rates' (Jan-2026); CAC = avg of GT 25, DO 27, CR 30, HN 30, NI 30, SV 30", "EBITDA-weighted to get the operating tax rate", F_P)

# ---------------- peers (data for Comps)
inp.section("8. Peer data for comparables (local currency millions; run-rate = latest quarter × 4 unless stated)", 13)
inp.header(["Peer", "Price (LC)", "Shares (m)", "Net fin. debt ex-leases", "Lease liabilities", "NCI", "EBITDA IFRS16 (qtr)", "EBITDAaL (qtr)", "Revenue (qtr)", "LC per USD", "Region", "Sources", "Inclusion rationale"])
PEERS = [
    ("Telefônica Brasil (VIVT3)", 35.30, 3196, -2380, 16182, 75, 7136, 5689, 16023, f"={I['usdbrl']}", "LatAm",
     "Price/shares/leases: Yahoo Finance (24-Nov-26). 3Q26 release (10-Nov-26): revenue 16,023, EBITDA 7,136, EBITDA AL 5,689; net debt incl. leases ~13.8 bn (Finsee summary)",
     "Integrated mobile+fibre, #1 in Brazil — AMX's direct competitor (Claro Brasil)"),
    ("TIM Brasil (TIMS3)", 21.88, 2377, -2652, 13778, 0, 3701, 2917, 7069, f"={I['usdbrl']}", "LatAm",
     "Price/shares/leases: Yahoo Finance. 3Q26 release: revenue 7,069, normalised EBITDA 3,701, EBITDA-AL 2,917, total net debt incl. leases 11,126 (→ ex-leases −2,652)",
     "Mobile-led Brazilian operator, direct competitor"),
    ("Millicom (TIGO)", 79.12, 168, 7801, 3190, 40, 1092, None, 2276, "=1", "LatAm",
     "Price/shares: Yahoo. 3Q26 release (5-Nov-26): revenue 2,276, adj. EBITDA 1,092. Debt 11,839 incl. leases 3,190, cash 848: Yahoo balance sheet",
     "Central America/Colombia integrated operator; run-rate captures 2026 acquisitions (Coltel, Ecuador, Uruguay)"),
    ("Entel (Chile)", 5435.0, 302, 1335700, 686279, 0, 216563, None, 767074, f"={I['usdclp']}", "LatAm",
     "Price: Yahoo. EBITDA = 1H26 433,125 / 2 (company, via press); NFD 1,335,700 = 1Q26 (Quartr summary of company release); revenue 3Q26 & leases: Yahoo",
     "Chile/Peru integrated — competes with Claro Chile/Peru. 3Q26 release not retrieved (older data flagged)"),
    ("Liberty Latin America (LILA)", 9.50, 190, 7978, 0, 474, None, 454, 1117, "=1", "LatAm",
     "Price/shares/debt/cash/NCI: Yahoo (Sep-26). 3Q26 8-K (SEC): revenue 1,117, Adjusted OIBDA 454",
     "Caribbean/Puerto Rico/Central America overlap with AMX. US GAAP: operating leases inside OIBDA → EBITDAaL basis only"),
    ("Telefónica (TEF)", 3.404, 5628, 25186, 7923, 3640, 2949, None, 8433, f"=1/{I['eurusd']}", "Europe",
     "Price: Yahoo. Jan–Sep 26 release: adj. EBITDA 8,846 (÷3), 3Q revenue 8,433, net financial debt 25,186; leases & NCI: Yahoo (Jun-26)",
     "Integrated EU/LatAm (Spain, Germany, UK JV, Brazil) — closest European analogue"),
    ("Orange (ORA)", 15.285, 2658, 35700, 8859, 3321, None, 4140, 12191, f"=1/{I['eurusd']}", "Europe",
     "Price: Yahoo. 3Q26 release (27-Oct-26): revenue 12,191, EBITDAaL 4,140; net financial debt 35.7 bn at 30-Jun-26 (H1 release); leases/NCI Yahoo",
     "Integrated EU + Africa/Middle East (emerging-market exposure like AMX)"),
    ("Vodafone (VOD)", 1.029, 23087, 33473, 12388, 3732, None, 3250, 10927, f"=1/{I['eurusd']}", "Europe",
     "Price GBp 102.9 (Yahoo) → GBP; H1 FY27 (10-Nov-26): revenue 21,853, adj. EBITDAaL 6,500 (÷2), net debt 33,473 (ex-leases); shares/leases Mar-26 (Yahoo). Price in GBP converted with GBP/EUR",
     "Integrated EU + Africa (Vodacom/Safaricom)"),
    ("KPN (KPN)", 4.342, 3772, 6254, 687, 45, None, 671, 1463, f"=1/{I['eurusd']}", "Europe",
     "Price/shares/leases: Yahoo. 3Q26 release: adj. EBITDA AL 671; net debt 6,254 = 1Q26 (latest found); revenue = 2Q26 (Yahoo)",
     "Integrated fixed-mobile, mature market (upper-bound multiple)"),
    ("Telia (TELIA)", 50.5, 3932, 57794, 17889, 3963, 8870, None, 21306, f"={I['usdmxn']}/19.4718*10.2664", "Europe",
     "Price: Yahoo. Jan–Sep 26 report (22-Oct-26): 3Q revenue 21,306, adj. EBITDA 8,870 (IFRS 16); debt/leases/cash/NCI: Yahoo (Sep-26)",
     "Integrated Nordic/Baltic operator"),
]
PR = {}
for pi, p in enumerate(PEERS):
    r = inp.r
    PR[pi] = r
    inp.put(r, 1, p[0])
    for j, v in enumerate(p[1:10]):
        if v is None:
            continue
        fmt = F_N2 if j == 0 else (F_N4 if j == 8 else F_M)
        inp.put(r, 2 + j, v, fmt=fmt, fill=FILL_IN if not (isinstance(v, str)) else None)
    inp.put(r, 11, p[10], font=NOTE)
    inp.put(r, 12, p[11], font=NOTE)
    inp.put(r, 13, p[12], font=NOTE)
    inp.r += 1
# Telia FX: SEK per USD from spot quote (keep as plain input)
inp.ws.cell(PR[9], 10).value = 10.2664
inp.ws.cell(PR[9], 10).font = BLUE
inp.ws.cell(PR[9], 10).fill = FILL_IN
inp.put(PR[9], 12, inp.ws.cell(PR[9], 12).value + " SEK per USD 10.2664 (Yahoo SEK=X, 25-Nov-26)", font=NOTE)
# Vodafone: price is GBP, multiply into EUR inside Comps via GBP/EUR factor
inp.r += 1
inp.note("Excluded peers (memo): A1 Telekom Austria — 61% owned by AMX (circular); Telecom Argentina — IAS 29 hyperinflation accounting distorts multiples; "
         "Deutsche Telekom — value dominated by T-Mobile US (different market); Desktop — acquired by AMX.")
inp.note("Basis: EV/EBITDA uses EV including lease liabilities and IFRS 16 EBITDA; EV/EBITDAaL uses EV excluding leases and EBITDA after leases. "
         "Run-rate (latest quarter ×4) is used instead of consensus forward estimates so no third-party estimates enter the valuation.")

inp.ws.freeze_panes = "B4"

# =========================================================================
# HISTORICALS
# =========================================================================
his = S(wb, "Historicals", "Historicals — reported 3 statements, MXN millions (SEC 20-F XBRL FY2021–FY2025; 3Q26 6-K for 9M)",
        [58, 6, 13, 13, 13, 13, 13, 13, 13, 13, 70])
H = HIST
YRS = ["FY2021", "FY2022", "FY2023", "FY2024", "FY2025", "9M2025", "9M2026"]
his.note("Sources: FY2021–FY2025 = AMX annual reports on Form 20-F, SEC EDGAR CIK 0001129137 (XBRL 'ifrs-full' facts; FY2023–25 from the FY2025 20-F filed 28-Apr-2026, "
         "earlier years from the latest 20-F that reports them). 9M = AMX 3Q26 earnings report furnished on Form 6-K, 22-Oct-2026 (PDF pp.11–15, 36). All blue cells are reported data.")
his.header(["Income statement", ""] + YRS + ["", "Notes"])


def hv(tag):
    return [None if v is None else round(v, 3) for v in H[tag]]


C0 = 3  # first data col
nine25 = {"rev": 698741, "svc": 595485, "equip": 98761, "cogs": 176683 + 86222, "sga": 153971, "oth": 4584, "da": 135005,
          "fin": 34068 - 106, "tax": 40577, "nci": 4051, "ni": 63685}
nine26 = {"rev": 728549, "svc": 617702, "equip": 105498, "cogs": 177250 + 92661, "sga": 159160, "oth": 6910, "da": 133122,
          "fin": 45095 - 445, "tax": 38405, "nci": 4376, "ni": 72015}


def row7(tag, k9):
    return hv(tag) + [nine25[k9], nine26[k9]] if tag else [None] * 5 + [nine25[k9], nine26[k9]]


his.line("rev", "Total revenue", row7("Revenue", "rev"))
his.line("svc", "  of which service revenue", row7("RevenueFromRenderingOfServices", "svc"))
his.line("equip", "  of which equipment revenue", row7("RevenueFromSaleOfTelecommunicationEquipment", "equip"))
his.line("other_rev", "  of which other revenue", [f"={CL(c)}{his.rows['rev']}-{CL(c)}{his.rows['svc']}-{CL(c)}{his.rows['equip']}" for c in range(C0, C0 + 7)])
his.line("cogs", "Cost of sales (services + equipment)", row7("CostOfSales", "cogs"))
his.line("sga", "Selling, general & administrative", row7("SellingGeneralAndAdministrativeExpense", "sga"))
his.line("oth", "Other operating expenses", row7("OtherExpenseByFunction", "oth"))
his.line("ebitda", "EBITDA", [f"={CL(c)}{his.rows['rev']}-{CL(c)}{his.rows['cogs']}-{CL(c)}{his.rows['sga']}-{CL(c)}{his.rows['oth']}" for c in range(C0, C0 + 7)], bold=True)
his.line("ebitda_m", "  EBITDA margin", [f"={CL(c)}{his.rows['ebitda']}/{CL(c)}{his.rows['rev']}" for c in range(C0, C0 + 7)], fmt=F_P)
his.line("da", "Depreciation & amortisation (incl. right-of-use)", row7("DepreciationAndAmortisationExpense", "da"))
his.line("rou_dep", "  memo: of which depreciation of right-of-use assets", [None, None, 31839, 32402, 39226, None, None], note="20-F XBRL DepreciationRightofuseAssets", note_col=11)
his.line("ebit", "Operating profit (EBIT)", [f"={CL(c)}{his.rows['ebitda']}-{CL(c)}{his.rows['da']}" for c in range(C0, C0 + 7)], bold=True)
his.line("ebit_rep", "  memo: EBIT as reported", hv("ProfitLossFromOperatingActivities") + [142275, 159446], note="Tie-out row (Checks sheet)", note_col=11)
his.line("fin", "Comprehensive financing cost, net of equity-method income", [None] * 5 + [nine25["fin"], nine26["fin"]],
         note="FY: EBIT − reported pre-tax profit; 9M: CFC − equity participation (6-K p.11)", note_col=11)
his.line("ebt", "Profit before tax", [f"={CL(c)}{his.rows['ebit']}-{CL(c)}{his.rows['fin']}" for c in range(C0, C0 + 7)])
his.line("ebt_rep", "  memo: profit before tax as reported", hv("ProfitLossBeforeTax") + [None, None])
his.line("tax", "Income tax", row7("IncomeTaxExpenseContinuingOperations", "tax"))
his.line("ni_cont", "Net income from continuing operations", [f"={CL(c)}{his.rows['ebt']}-{CL(c)}{his.rows['tax']}" for c in range(C0, C0 + 7)])
his.line("disc", "Discontinued operations & other", [None] * 5 + [0, 0],
         note="FY21: gain on sale of TracFone (US). FY22: per 20-F", note_col=11)
his.line("ni_total", "Net income (incl. NCI)", [f"={CL(c)}{his.rows['ni_cont']}+{CL(c)}{his.rows['disc']}" for c in range(C0, C0 + 7)])
his.line("ni_total_rep", "  memo: profit for the year as reported", hv("ProfitLoss") + [None, None])
his.line("nci", "Non-controlling interests", row7("ProfitLossAttributableToNoncontrollingInterests", "nci"))
his.line("ni", "Net income attributable to AMX shareholders", [f"={CL(c)}{his.rows['ni_total']}-{CL(c)}{his.rows['nci']}" for c in range(C0, C0 + 7)], bold=True)
his.line("ni_rep", "  memo: as reported", hv("ProfitLossAttributableToOwnersOfParent") + [nine25["ni"], nine26["ni"]])
# fix forward references now that rows exist
for c in range(C0, C0 + 5):
    his.ws.cell(his.rows["fin"], c).value = f"={CL(c)}{his.rows['ebit']}-{CL(c)}{his.rows['ebt_rep']}"
    his.ws.cell(his.rows["disc"], c).value = f"={CL(c)}{his.rows['ni_total_rep']}-{CL(c)}{his.rows['ni_cont']}"

his.r += 1
his.header(["Balance sheet (period end)", "", "Dec-21", "Dec-22", "Dec-23", "Dec-24", "Dec-25", "", "Sep-26", "", "Notes"])
bs_sep = {"cash": 83829, "rec": 261578, "inv": 33410, "ca": 399970, "ppe": 688994, "intang": 136293, "rou": 200072, "gw": 157298, "assets": 1854716,
          "std": 85275, "stl": 39012, "pay": 141038, "cl": 534699, "ltd": 398634, "ltl": 179979, "ncl": 852641, "eq": 467376}


def bsrow(tag, k):
    v = hv(tag)
    return v + [None, bs_sep.get(k)] if k else v + [None, None]


his.line("cash", "Cash & equivalents (Sep-26: incl. marketable securities)", bsrow("CashAndCashEquivalents", "cash"))
his.line("rec", "Trade & other receivables", bsrow("TradeAndOtherCurrentReceivables", "rec"))
his.line("inv", "Inventories", bsrow("Inventories", "inv"))
his.line("oca", "Other current assets", [None] * 7)
r_oca = his.rows["oca"]
his.line("ca", "Total current assets", bsrow("CurrentAssets", "ca"), bold=True)
his.line("ppe", "Property, plant & equipment", bsrow("PropertyPlantAndEquipment", "ppe"))
his.line("intang", "Intangible assets (licences, other)", bsrow("IntangibleAssetsOtherThanGoodwill", "intang"))
his.line("rou", "Right-of-use assets", bsrow("RightofuseAssets", "rou"))
his.line("gw", "Goodwill", bsrow("Goodwill", "gw"))
his.line("onca", "Other non-current assets (deferred taxes, affiliates, other)", [None] * 7)
his.line("assets", "Total assets", [f"={CL(c)}{his.rows['ca']}+{CL(c)}{his.rows['ppe']}+{CL(c)}{his.rows['intang']}+{CL(c)}{his.rows['rou']}+{CL(c)}{his.rows['gw']}+{CL(c)}{his.rows['onca']}" if c != C0 + 5 else None for c in range(C0, C0 + 7)], bold=True)
his.line("assets_rep", "  memo: total assets as reported", bsrow("Assets", "assets"))
for c in range(C0, C0 + 7):
    if c == C0 + 5:
        continue
    his.ws.cell(r_oca, c).value = f"={CL(c)}{his.rows['ca']}-{CL(c)}{his.rows['cash']}-{CL(c)}{his.rows['rec']}-{CL(c)}{his.rows['inv']}"
    his.ws.cell(his.rows["onca"], c).value = f"={CL(c)}{his.rows['assets_rep']}-{CL(c)}{his.rows['ca']}-{CL(c)}{his.rows['ppe']}-{CL(c)}{his.rows['intang']}-{CL(c)}{his.rows['rou']}-{CL(c)}{his.rows['gw']}"
    for k in ("oca", "onca"):
        his.ws.cell(his.rows[k], c).number_format = F_M
        his.ws.cell(his.rows[k], c).font = BLACK
his.line("std", "Short-term debt & current portion of LT debt", bsrow("CurrentBorrowingsAndCurrentPortionOfNoncurrentBorrowings", "std"))
his.line("stl", "Current lease liabilities", bsrow("CurrentLeaseLiabilities", "stl"))
his.line("pay", "Trade & other payables", bsrow("TradeAndOtherCurrentPayables", "pay"))
his.line("ocl", "Other current liabilities", [None] * 7)
his.line("cl", "Total current liabilities", bsrow("CurrentLiabilities", "cl"), bold=True)
his.line("ltd", "Long-term debt", bsrow("LongtermBorrowings", "ltd"))
his.line("ltl", "Non-current lease liabilities", bsrow("NoncurrentLeaseLiabilities", "ltl"))
his.line("oncl", "Other non-current liabilities (pensions, deferred tax, other)", [None] * 7)
his.line("ncl", "Total non-current liabilities", bsrow("NoncurrentLiabilities", "ncl"), bold=True)
his.line("eqp", "Equity attributable to AMX shareholders", bsrow("EquityAttributableToOwnersOfParent", None))
his.line("nci_bs", "Non-controlling interests", bsrow("NoncontrollingInterests", None))
his.line("eq", "Total equity", [f"={CL(c)}{his.rows['eqp']}+{CL(c)}{his.rows['nci_bs']}" if c < C0 + 5 else None for c in range(C0, C0 + 7)], bold=True)
his.ws.cell(his.rows["eq"], C0 + 6).value = bs_sep["eq"]
his.ws.cell(his.rows["eq"], C0 + 6).font = BLUE
his.ws.cell(his.rows["eq"], C0 + 6).number_format = F_M
his.line("le", "Total liabilities & equity", [f"={CL(c)}{his.rows['cl']}+{CL(c)}{his.rows['ncl']}+{CL(c)}{his.rows['eq']}" if c != C0 + 5 else None for c in range(C0, C0 + 7)], bold=True)
his.line("eq_rep", "  memo: total equity as reported", bsrow("Equity", "eq"))
for c in range(C0, C0 + 7):
    if c == C0 + 5:
        continue
    his.ws.cell(his.rows["ocl"], c).value = f"={CL(c)}{his.rows['cl']}-{CL(c)}{his.rows['std']}-{CL(c)}{his.rows['stl']}-{CL(c)}{his.rows['pay']}"
    his.ws.cell(his.rows["oncl"], c).value = f"={CL(c)}{his.rows['ncl']}-{CL(c)}{his.rows['ltd']}-{CL(c)}{his.rows['ltl']}"
    for k in ("ocl", "oncl"):
        his.ws.cell(his.rows[k], c).number_format = F_M
        his.ws.cell(his.rows[k], c).font = BLACK
his.line("netdebt", "memo: net financial debt ex-leases (company definition)", [None, None, None, None, f"={CL(C0+4)}{his.rows['std']}+{CL(C0+4)}{his.rows['ltd']}-77384", None, f"={CL(C0+6)}{his.rows['std']}+{CL(C0+6)}{his.rows['ltd']}-{CL(C0+6)}{his.rows['cash']}"],
         note="Dec-25 uses cash + marketable securities 77,384 (6-K p.13/p.37); Sep-26 ties to 6-K net debt 400,081", note_col=11)
his.ws.cell(his.rows["netdebt"], C0 + 4).value = f"={CL(C0+4)}{his.rows['std']}+{CL(C0+4)}{his.rows['ltd']}-{CL(C0+4)}{his.rows['cash_sec_dec25']}" if False else his.ws.cell(his.rows["netdebt"], C0 + 4).value
his.line("cash_sec", "memo: cash + marketable securities (6-K definition)", [None, None, None, None, 77384, None, 83829], note="AMX 3Q26 6-K p.13 (Dec-25 and Sep-26 columns)", note_col=11)
his.ws.cell(his.rows["netdebt"], C0 + 4).value = f"={CL(C0+4)}{his.rows['std']}+{CL(C0+4)}{his.rows['ltd']}-{CL(C0+4)}{his.rows['cash_sec']}"
his.line("pension", "memo: net defined-benefit obligation", [None, 128400, 134421, 157867, 194736, None, None], note="20-F XBRL LiabilityAssetOfDefinedBenefitPlans", note_col=11)

his.r += 1
his.header(["Cash flow statement", "", "FY2021", "FY2022", "FY2023", "FY2024", "FY2025", "9M2025", "9M2026", "", "Notes"])
cf9_25 = {"cfo": 186665, "capex": -84870, "leasep": -38710, "intp": -27488, "div": -17757, "buy": -10815}
cf9_26 = {"cfo": 204793, "capex": -75915, "leasep": -39639, "intp": -26145, "div": -18132, "buy": -10821}
neg = lambda tag: [None if v is None else -v for v in hv(tag)]
his.line("cfo", "Net cash from operating activities", hv("CashFlowsFromUsedInOperatingActivities") + [cf9_25["cfo"], cf9_26["cfo"]])
his.line("capex_ppe", "Purchase of PP&E", neg("PurchaseOfPropertyPlantAndEquipmentClassifiedAsInvestingActivities") + [None, None])
his.line("capex_int", "Purchase of intangibles (incl. spectrum)", neg("PurchaseOfIntangibleAssetsClassifiedAsInvestingActivities") + [None, None])
his.line("capex", "Total capex", [f"={CL(c)}{his.rows['capex_ppe']}+{CL(c)}{his.rows['capex_int']}" for c in range(C0, C0 + 5)] + [cf9_25["capex"], cf9_26["capex"]], bold=True,
         note="9M: company 'capital expenditures' (6-K p.15 FCF table)", note_col=11)
his.line("cfi_oth", "Other investing (acquisitions, disposals, securities)", [None] * 7)
his.line("cfi", "Net cash from investing activities", hv("CashFlowsFromUsedInInvestingActivities") + [None, None])
his.line("div", "Dividends paid", neg("DividendsPaidClassifiedAsFinancingActivities") + [cf9_25["div"], cf9_26["div"]])
his.line("buy", "Share buybacks", neg("PurchaseOfTreasuryShares") + [cf9_25["buy"], cf9_26["buy"]])
his.line("leasep", "Lease payments", neg("PaymentsOfLeaseLiabilitiesClassifiedAsFinancingActivities") + [cf9_25["leasep"], cf9_26["leasep"]])
his.line("intp", "Interest paid", neg("InterestPaidClassifiedAsFinancingActivities") + [cf9_25["intp"], cf9_26["intp"]], note="9M: 'net interest paid' (6-K p.36)", note_col=11)
proc = hv("ProceedsFromBorrowingsClassifiedAsFinancingActivities")
proc[0] = 93675.127  # FY2021: ifrs-full:ProceedsFromNoncurrentBorrowings (FY2021 20-F)
rep = hv("RepaymentsOfBorrowingsClassifiedAsFinancingActivities")
his.line("netborr", "Net borrowing (proceeds − repayments)", [round(p - q, 3) for p, q in zip(proc, rep)] + [None, None], note="FY21 proceeds = ProceedsFromNoncurrentBorrowings tag", note_col=11)
his.line("cff_oth", "Other financing", [None] * 7)
his.line("cff", "Net cash from financing activities", hv("CashFlowsFromUsedInFinancingActivities") + [None, None])
his.line("fxc", "FX effect on cash", hv("EffectOfExchangeRateChangesOnCashAndCashEquivalents") + [None, None])
his.line("dcash", "Net change in cash", [f"={CL(c)}{his.rows['cfo']}+{CL(c)}{his.rows['cfi']}+{CL(c)}{his.rows['cff']}+{CL(c)}{his.rows['fxc']}" for c in range(C0, C0 + 5)] + [None, None], bold=True)
his.line("open_cash", "Opening cash & equivalents", [35917.907] + [f"={CL(c-1)}{his.rows['cash']}" for c in range(C0 + 1, C0 + 5)] + [None, None], note="Dec-20 cash: FY2021 20-F XBRL", note_col=11)
his.line("close_cash", "Closing cash (opening + change)", [f"={CL(c)}{his.rows['open_cash']}+{CL(c)}{his.rows['dcash']}" for c in range(C0, C0 + 5)] + [None, None])
his.line("fcf_co", "memo: company FCF (CFO − leases − interest − capex)", [f"={CL(c)}{his.rows['cfo']}+{CL(c)}{his.rows['leasep']}+{CL(c)}{his.rows['intp']}+{CL(c)}{his.rows['capex']}" for c in range(C0, C0 + 7)],
         note="9M26 company-reported FCF 74,960 also includes other items (6-K p.15, p.36)", note_col=11)
for c in range(C0, C0 + 5):
    his.put(his.rows["cfi_oth"], c, f"={CL(c)}{his.rows['cfi']}-{CL(c)}{his.rows['capex']}", fmt=F_M)
    his.put(his.rows["cff_oth"], c, f"={CL(c)}{his.rows['cff']}-{CL(c)}{his.rows['div']}-{CL(c)}{his.rows['buy']}-{CL(c)}{his.rows['leasep']}-{CL(c)}{his.rows['intp']}-{CL(c)}{his.rows['netborr']}", fmt=F_M)
his.r += 1
his.note("9M columns: the 6-K does not include a full IFRS cash-flow statement (only the FCF reconciliation), so the 9M cash flow is shown for the available lines and is not tied.")
his.ws.freeze_panes = "C4"

# =========================================================================
# DRIVERS (regional build)
# =========================================================================
drv = S(wb, "Drivers", "Regional drivers — subscribers × ARPU by country, local currency → MXN (postpaid = engine, prepaid = noise)",
        [52, 10, 14, 14, 14, 14, 14, 14, 14, 14, 60])
DY = ["3Q26A calib.", "4Q26E", "FY2026E", "FY2027E", "FY2028E", "FY2029E", "FY2030E", "FY2031E"]
YCOL = {"cal": 3, "q4": 4, "y26": 5, 2027: 6, 2028: 7, 2029: 8, 2030: 9, 2031: 10}
drv.note("Method: 3Q26 blended mobile ARPU is calibrated to reported mobile service revenue and split postpaid/prepaid with the ratio k (Inputs). "
         "Postpaid revenue = average postpaid subs × postpaid ARPU × 12. Prepaid revenue grows at an input rate (subscriber counts are distorted by base cleanups). "
         "Fixed service revenue = average broadband accesses × fixed revenue per access × 12. FY2026E = 9M26 actual + modelled 4Q26.")
drv.line("year", "Year index (years after 2026)", [None, None, 0, 1, 2, 3, 4, 5], fmt="0", note="Used for FX drift, fades and margin drift", note_col=11)
drv.ws.cell(drv.rows["year"], 3).value = None
for c in range(5, 11):
    drv.ws.cell(drv.rows["year"], c).font = BLUE
DR = {}  # (key, region) -> row


def dref(key, i, col):
    return f"{CL(col)}{DR[(key, i)]}"


for i, reg in enumerate(REG):
    drv.r += 1
    drv.section(f"{reg}  ({CCY[i]})", 11)
    drv.header(["", ""] + DY + ["", "Comment"])
    R = lambda key: DR[(key, i)]
    g = lambda name: RC[(name, i)]
    yr = drv.rows["year"]

    def L_(key, label, vals, fmt=F_M, bold=False, note=None):
        DR[(key, i)] = drv.line(None, label, vals, fmt=fmt, bold=bold, note=note, note_col=11)

    # Postpaid subscribers
    L_("post_na", "Postpaid net adds (k)", [f"={g('post_sep26')}-{g('post_sep25')}", f"={g('q4_post')}", f"=C{drv.r}*0+0" if False else None,
                                            f"=$C{drv.r}*{I['post_2027']}", f"=F{drv.r}*(1-{I['post_fade']})", f"=G{drv.r}*(1-{I['post_fade']})", f"=H{drv.r}*(1-{I['post_fade']})", f"=I{drv.r}*(1-{I['post_fade']})"],
       fmt=F_K, note="3Q26 column = LTM net adds (Sep-25→Sep-26)")
    L_("post_eop", "Postpaid subscribers, end of period (k)", [f"={g('post_sep26')}", f"=C{drv.r}+D{R('post_na')}", f"=D{drv.r}",
                                                             f"=E{drv.r}+F{R('post_na')}", f"=F{drv.r}+G{R('post_na')}", f"=G{drv.r}+H{R('post_na')}", f"=H{drv.r}+I{R('post_na')}", f"=I{drv.r}+J{R('post_na')}"], fmt=F_K)
    L_("post_avg", "Average postpaid subscribers (k)", [f"=AVERAGE({g('post_sep26')}-{g('q4_post')},{g('post_sep26')})", f"=AVERAGE(C{R('post_eop')},D{R('post_eop')})", None] + [f"=AVERAGE({CL(c-1)}{R('post_eop')},{CL(c)}{R('post_eop')})" for c in range(6, 11)], fmt=F_K,
       note="3Q26: Sep-26 less half of the quarter's net adds")
    L_("pre_avg", "Average prepaid subscribers 3Q26 (k) — calibration only", [f"=AVERAGE({g('tot_jun26')},{g('post_sep26')}+{g('pre_sep26')})-C{R('post_avg')}"] + [None] * 7, fmt=F_K)
    L_("pre_eop", "Prepaid subscribers, end of period (k) — memo", [f"={g('pre_sep26')}", f"=C{drv.r}+{g('q4_pre')}", f"=D{drv.r}"] + [None] * 5, fmt=F_K, note="Base cleanups (MX, CO) make prepaid subs a poor revenue driver")
    L_("mix", "Postpaid share of wireless base, Sep-26 (memo)", [f"={g('post_sep26')}/({g('post_sep26')}+{g('pre_sep26')})"] + [None] * 7, fmt=F_P)
    L_("blend_arpu", "Blended mobile ARPU, monthly (LC) — memo vs. 6-K reported ARPU", [f"={g('mob_svc_3q')}/(C{R('post_avg')}+C{R('pre_avg')})/{I['months_q']}*{I['units_k']}"] + [None] * 7, fmt=F_N2,
       note="Revenue in LC m, subs in k → ×1,000")
    L_("pre_arpu", "Prepaid ARPU, monthly (LC) — calibrated", [f"={g('mob_svc_3q')}/({g('k_ratio')}*C{R('post_avg')}+C{R('pre_avg')})/{I['months_q']}*{I['units_k']}"] + [None] * 7, fmt=F_N2,
       note="Solves revenue = avg post × k × p + avg pre × p")
    L_("post_arpu", "Postpaid ARPU, monthly (LC)", [f"=C{R('pre_arpu')}*{g('k_ratio')}", f"=C{drv.r}", None, f"=C{drv.r}*(1+{g('g_post_arpu')})"] + [f"={CL(c-1)}{drv.r}*(1+{g('g_post_arpu')})" for c in range(7, 11)], fmt=F_N2)
    L_("post_rev", "Postpaid service revenue (LC m)", [f"=C{R('post_avg')}*C{R('post_arpu')}*{I['months_q']}/{I['units_k']}",
                                                    f"=D{R('post_avg')}*D{R('post_arpu')}*{I['months_q']}/{I['units_k']}", None] +
       [f"={CL(c)}{R('post_avg')}*{CL(c)}{R('post_arpu')}*{I['months_y']}/{I['units_k']}" for c in range(6, 11)], bold=True, note="THE ENGINE")
    L_("pre_rev", "Prepaid service revenue (LC m)", [f"={g('mob_svc_3q')}-C{R('post_rev')}", f"=C{drv.r}", None, f"=C{drv.r}*{I['q_per_y']}*(1+{g('g_prepaid')})"] +
       [f"={CL(c-1)}{drv.r}*(1+{g('g_prepaid')})" for c in range(7, 11)], note="2027E = 3Q26 run-rate ×4 × (1+g)")
    L_("mob_rev", "Mobile service revenue (LC m)", [f"=C{R('post_rev')}+C{R('pre_rev')}", f"=D{R('post_rev')}+D{R('pre_rev')}", None] + [f"={CL(c)}{R('post_rev')}+{CL(c)}{R('pre_rev')}" for c in range(6, 11)])
    # Broadband
    L_("bb_na", "Broadband net adds (k)", [f"={g('bb_sep26')}-{g('bb_sep25')}", f"={g('q4_bb')}", None, f"=$C{drv.r}*{I['bb_2027']}"] + [f"={CL(c-1)}{drv.r}*(1-{I['bb_fade']})" for c in range(7, 11)], fmt=F_K, note="3Q26 column = LTM")
    L_("bb_eop", "Broadband accesses, end of period (k)", [f"={g('bb_sep26')}", f"=C{drv.r}+D{R('bb_na')}", f"=D{drv.r}"] + [f"={CL(c-1)}{drv.r}+{CL(c)}{R('bb_na')}" for c in range(6, 11)], fmt=F_K)
    L_("bb_avg", "Average broadband accesses (k)", [f"=AVERAGE({g('bb_jun26')},{g('bb_sep26')})", f"=AVERAGE(C{R('bb_eop')},D{R('bb_eop')})", None] + [f"=AVERAGE({CL(c-1)}{R('bb_eop')},{CL(c)}{R('bb_eop')})" for c in range(6, 11)], fmt=F_K)
    L_("fix_arpu", "Fixed service revenue per broadband access, monthly (LC)", [f"=({g('svc_3q')}-{g('mob_svc_3q')})/C{R('bb_avg')}/{I['months_q']}*{I['units_k']}", f"=C{drv.r}", None, f"=C{drv.r}*(1+{g('g_fixed_arpu')})"] +
       [f"={CL(c-1)}{drv.r}*(1+{g('g_fixed_arpu')}*($J${yr}-{CL(c)}${yr})/($J${yr}-$F${yr}))" for c in range(7, 11)], fmt=F_N2, note="Includes voice, Pay-TV, corporate networks; growth fades to 0% by 2031E")
    L_("fix_rev", "Fixed service revenue (LC m)", [f"={g('svc_3q')}-{g('mob_svc_3q')}", f"=D{R('bb_avg')}*D{R('fix_arpu')}*{I['months_q']}/{I['units_k']}", None] + [f"={CL(c)}{R('bb_avg')}*{CL(c)}{R('fix_arpu')}*{I['months_y']}/{I['units_k']}" for c in range(6, 11)])
    L_("svc", "Total service revenue (LC m)", [f"=C{R('mob_rev')}+C{R('fix_rev')}", f"=D{R('mob_rev')}+D{R('fix_rev')}", None] + [f"={CL(c)}{R('mob_rev')}+{CL(c)}{R('fix_rev')}" for c in range(6, 11)], bold=True)
    L_("eq_ratio", "Equipment & other revenue, % of service revenue", [f"=({g('rev_3q')}-{g('svc_3q')})/{g('svc_3q')}"] + [None] * 7, fmt=F_P, note="3Q26 ratio held (low-margin handset sales)")
    L_("eq_rev", "Equipment & other revenue (LC m)", [f"={g('rev_3q')}-{g('svc_3q')}", f"=D{R('svc')}*$C{drv.r-1}", None] + [f"={CL(c)}{R('svc')}*$C{R('eq_ratio')}" for c in range(6, 11)])
    desk = (i == 1)
    if desk:
        L_("desk_rev", "Desktop revenue (toggle; 0 in base case)", [None, None, None, f"={I['desk_on']}*{I['desk_rev']}"] + [f"={CL(c-1)}{drv.r}*(1+{g('infl')})" for c in range(7, 11)], note="Grows with Brazil inflation if included")
    L_("rev", "Total revenue (LC m)", [f"={g('rev_3q')}", f"=D{R('svc')}+D{R('eq_rev')}", f"={g('rev_9m')}+D{drv.r}"] +
       [f"={CL(c)}{R('svc')}+{CL(c)}{R('eq_rev')}" + (f"+{CL(c)}{R('desk_rev')}" if desk else "") for c in range(6, 11)], bold=True)
    L_("rev_g", "  growth y/y (LC)", [None, None, None, f"=F{R('rev')}/E{R('rev')}-1"] + [f"={CL(c)}{R('rev')}/{CL(c-1)}{R('rev')}-1" for c in range(7, 11)], fmt=F_P)
    L_("margin", "EBITDA margin", [f"={g('ebitda_9m')}/{g('rev_9m')}", f"=C{drv.r}", f"=E{drv.r+1}/E{R('rev')}"] +
       [f"=$C{drv.r}+{I['margin_drift']}*MIN({CL(c)}${yr},{I['drift_years']})" for c in range(6, 11)], fmt=F_P, note="3Q26 column = 9M26 margin (base)")
    L_("ebitda", "EBITDA (LC m)", [f"={g('ebitda_3q')}", f"=D{R('rev')}*D{R('margin')}", f"={g('ebitda_9m')}+D{drv.r}"] +
       [f"={CL(c)}{R('rev')}*{CL(c)}{R('margin')}" + (f"+{I['desk_on']}*{I['desk_ebitda']}*{CL(c)}{R('desk_rev')}/MAX({I['desk_rev']},1)/MAX({I['desk_on']},1)" if desk else "") for c in range(6, 11)], bold=True)
    L_("fx", "FX: local currency per MXN", [f"={g('fx_3q')}", f"={g('fx_spot')}*(1+{I['fx_shock']})", None] +
       [f"=$D{drv.r}*((1+{g('infl')})/(1+{RC[('infl', 0)]}))^{CL(c)}${yr}" for c in range(6, 11)], fmt=F_N4, note="Spot at valuation date, then relative-PPP drift (inflation differential vs Mexico)")
    L_("rev_mxn", "Revenue (MXN m)", [f"=C{R('rev')}/C{R('fx')}", f"=D{R('rev')}/D{R('fx')}", f"={g('rev_9m')}/{g('fx_9m')}+D{drv.r}"] + [f"={CL(c)}{R('rev')}/{CL(c)}{R('fx')}" for c in range(6, 11)], bold=True)
    L_("ebitda_mxn", "EBITDA (MXN m)", [f"=C{R('ebitda')}/C{R('fx')}", f"=D{R('ebitda')}/D{R('fx')}", f"={g('ebitda_9m')}/{g('fx_9m')}+D{drv.r}"] + [f"={CL(c)}{R('ebitda')}/{CL(c)}{R('fx')}" for c in range(6, 11)], bold=True)
    L_("chk", "Check: calibrated 3Q26 revenue − reported (must be 0)", [f"=C{R('svc')}+C{R('eq_rev')}-{g('rev_3q')}"] + [None] * 7, fmt=F_M1)
    # format fix for blank cells kept None

# consolidation block
drv.r += 1
drv.section("Consolidation: sum of regions, eliminations / corporate", 11)
drv.header(["", ""] + DY + ["", "Comment"])
cols = [3, 4, 5, 6, 7, 8, 9, 10]
drv.line("sum_rev", "Sum of regional revenue (MXN m)", [("=" + "+".join(f"{CL(c)}{DR[('rev_mxn', i)]}" for i in range(NR))) for c in cols], bold=True)
drv.line("sum_ebitda", "Sum of regional EBITDA (MXN m)", [("=" + "+".join(f"{CL(c)}{DR[('ebitda_mxn', i)]}" for i in range(NR))) for c in cols], bold=True)
drv.line("elim_rev", "Other businesses & eliminations, % of regional revenue", [f"={I['rev_3q']}/C{drv.rows['sum_rev']}-1"] + [None] * 7, fmt=F_P2, note="Calibrated on 3Q26 (consolidated vs. sum of countries at 3Q average FX) and held")
drv.line("elim_ebitda", "Corporate & eliminations, % of regional EBITDA", [f"={I['ebitda_3q']}/C{drv.rows['sum_ebitda']}-1"] + [None] * 7, fmt=F_P2)
drv.line("cons_rev", "Consolidated revenue (MXN m)", [f"={I['rev_3q']}", f"=D{drv.rows['sum_rev']}*(1+$C{drv.rows['elim_rev']})", f"={I['rev_9m']}+D{drv.r}"] +
         [f"={CL(c)}{drv.rows['sum_rev']}*(1+$C{drv.rows['elim_rev']})" for c in range(6, 11)], bold=True)
drv.line("cons_ebitda", "Consolidated EBITDA (MXN m)", [f"={I['ebitda_3q']}", f"=D{drv.rows['sum_ebitda']}*(1+$C{drv.rows['elim_ebitda']})", f"={I['ebitda_9m']}+D{drv.r}"] +
         [f"={CL(c)}{drv.rows['sum_ebitda']}*(1+$C{drv.rows['elim_ebitda']})" for c in range(6, 11)], bold=True)
drv.line("tax_w", "Operating tax rate (EBITDA-weighted statutory)", [None, None, None] +
         [("=(" + "+".join(f"{CL(c)}{DR[('ebitda_mxn', i)]}*{RC[('tax', i)]}" for i in range(NR)) + f")/{CL(c)}{drv.rows['sum_ebitda']}") for c in range(6, 11)], fmt=F_P)
drv.ws.cell(drv.rows["tax_w"], 5).value = f"=F{drv.rows['tax_w']}"
drv.ws.cell(drv.rows["tax_w"], 5).number_format = F_P
drv.line("w_rev", "Revenue weights 9M26 (for CRP) — computed below per region", [None] * 8)
for i in range(NR):
    drv.line(f"w_{i}", f"  {REG[i]}: share of 9M26 regional revenue (MXN)", [f"=({RC[('rev_9m', i)]}/{RC[('fx_9m', i)]})/(" + "+".join(f"{RC[('rev_9m', j)]}/{RC[('fx_9m', j)]}" for j in range(NR)) + ")"] + [None] * 7, fmt=F_P)
drv.line("post_total", "Memo: total postpaid subs, end of period (k)", [("=" + "+".join(f"{CL(c)}{DR[('post_eop', i)]}" for i in range(NR))) for c in cols], fmt=F_K)
drv.line("bb_total", "Memo: total broadband accesses, end of period (k)", [("=" + "+".join(f"{CL(c)}{DR[('bb_eop', i)]}" for i in range(NR))) for c in cols], fmt=F_K)
drv.ws.freeze_panes = "C3"

# =========================================================================
# PROJECTIONS (consolidated, 3 statements)
# =========================================================================
prj = S(wb, "Projections", "Projections — consolidated FCFF build and projected 3 statements, MXN millions",
        [56, 6, 14, 14, 14, 14, 14, 14, 14, 16, 60])
PY = ["FY2025A", "FY2026E", "FY2027E", "FY2028E", "FY2029E", "FY2030E", "FY2031E", "Terminal"]
PC = {"A": 3, 2026: 4, 2027: 5, 2028: 6, 2029: 7, 2030: 8, 2031: 9, "T": 10}
DCOL = {4: 5, 5: 6, 6: 7, 7: 8, 8: 9, 9: 10}  # projections col -> drivers col
HD = lambda key: f"Historicals!$G${his.rows[key]}"  # FY2025 column G
prj.note("Leases: FCFF deducts lease payments (EBITDAaL basis), so lease liabilities are excluded from net debt; right-of-use assets and lease liabilities are held flat in the projected balance sheet.")
prj.section("A. Operating build (FCFF)", 11)
prj.header(["", ""] + PY + ["Comment"])
yrs = [2026, 2027, 2028, 2029, 2030, 2031]
prj.line("rev", "Revenue", [f"={HD('rev')}"] + [f"=Drivers!{CL(DCOL[c])}{drv.rows['cons_rev']}" for c in range(4, 10)] + [f"=I{prj.r}*(1+{I['g']})"], bold=True)
prj.line("rev_g", "  growth", [None] + [f"={CL(c)}{prj.rows['rev']}/{CL(c-1)}{prj.rows['rev']}-1" for c in range(4, 11)], fmt=F_P)
prj.line("ebitda", "EBITDA", [f"={HD('ebitda')}"] + [f"=Drivers!{CL(DCOL[c])}{drv.rows['cons_ebitda']}" for c in range(4, 10)] + [f"=J{prj.rows['rev']}*I{prj.r+1}"], bold=True)
prj.line("margin", "  EBITDA margin", [f"={CL(c)}{prj.rows['ebitda']}/{CL(c)}{prj.rows['rev']}" for c in range(3, 10)] + [f"=I{prj.r}"], fmt=F_P)
prj.ws.cell(prj.rows["margin"], 10).value = f"=I{prj.rows['margin']}"
prj.ws.cell(prj.rows["ebitda"], 10).value = f"=J{prj.rows['rev']}*J{prj.rows['margin']}"
prj.line("lease", "Lease payments", [f"={HD('leasep')}"] + [f"=-{CL(c)}{prj.rows['rev']}*{I['lease_pct']}" for c in range(4, 11)])
prj.line("ebitdaal", "EBITDAaL", [f"={CL(c)}{prj.rows['ebitda']}+{CL(c)}{prj.rows['lease']}" for c in range(3, 11)], bold=True)
prj.line("da_pct", "  D&A on owned assets, % revenue", [None] + [f"={I['da_2026']}+({I['da_term']}-{I['da_2026']})*Drivers!{CL(DCOL[c])}{drv.rows['year']}/Drivers!$J${drv.rows['year']}" for c in range(4, 10)] + [f"={I['da_term']}"], fmt=F_P)
prj.line("da", "D&A on owned assets (ex right-of-use)", [f"=-({HD('da')}-39226)"] + [f"=-{CL(c)}{prj.rows['rev']}*{CL(c)}{prj.rows['da_pct']}" for c in range(4, 11)], note="FY25: total D&A less RoU depreciation 39,226 (20-F XBRL)")
prj.ws.cell(prj.rows["da"], 3).value = f"=-{HD('da')}+{HD('rou_dep')}"
prj.line("ebit", "EBIT after leases", [f"={CL(c)}{prj.rows['ebitdaal']}+{CL(c)}{prj.rows['da']}" for c in range(3, 11)], bold=True)
prj.line("taxr", "  operating tax rate", [None] + [f"=Drivers!{CL(DCOL[c])}{drv.rows['tax_w']}" for c in range(4, 10)] + [f"=I{prj.r}"], fmt=F_P)
prj.line("tax_ebit", "Taxes on EBIT", [None] + [f"=-{CL(c)}{prj.rows['ebit']}*{CL(c)}{prj.rows['taxr']}" for c in range(4, 11)])
prj.line("nopat", "NOPAT", [None] + [f"={CL(c)}{prj.rows['ebit']}+{CL(c)}{prj.rows['tax_ebit']}" for c in range(4, 11)], bold=True)
fxusd = lambda c: f"(1/Drivers!{CL(DCOL[c])}{DR[('fx', 4)]})"
prj.line("capex", "Capex (PP&E + intangibles)", [f"={HD('capex')}",
                                                 None,
                                                 f"=-{I['capex_usd_27']}/Drivers!F{DR[('fx', 4)]}"] +
         [f"=-{CL(c)}{prj.rows['rev']}*{I['capex_pct']}" for c in range(6, 11)], note="2026E–27E: USD 7 bn guidance × FX (2026: 9M avg / 4Q spot weighted 3:1); 2028E+: % revenue")
# cleaner 2026 capex formula (weights 9M/12 and 3/12 from month inputs)
prj.ws.cell(prj.rows["capex"], 4).value = (f"=-{I['capex_usd_26']}*((1/{RC[('fx_9m', 4)]})*({I['months_y']}-{I['months_q']})/{I['months_y']}"
                                           f"+(1/Drivers!D{DR[('fx', 4)]})*{I['months_q']}/{I['months_y']})")
prj.line("capex_pct", "  capex % revenue", [f"=-{CL(c)}{prj.rows['capex']}/{CL(c)}{prj.rows['rev']}" for c in range(3, 11)], fmt=F_P)
prj.line("nwc", "Net working capital", [f"={HD('rec')}+{HD('inv')}-{HD('pay')}"] + [f"={CL(c)}{prj.rows['rev']}*{I['nwc_pct']}" for c in range(4, 11)])
prj.line("dnwc", "(Increase) / decrease in NWC", [None] + [f"=-({CL(c)}{prj.rows['nwc']}-{CL(c-1)}{prj.rows['nwc']})" for c in range(4, 11)])
prj.line("fcff", "Free cash flow to firm (FCFF)", [None] + [f"={CL(c)}{prj.rows['nopat']}-{CL(c)}{prj.rows['da']}+{CL(c)}{prj.rows['capex']}+{CL(c)}{prj.rows['dnwc']}" for c in range(4, 11)], bold=True,
         note="NOPAT + D&A − capex − ΔNWC; terminal column uses capex = D&A (steady state)")
for c in range(4, 11):
    prj.ws.cell(prj.rows["fcff"], c).fill = FILL_KEY
prj.line("fcff_m", "  FCFF % revenue", [None] + [f"={CL(c)}{prj.rows['fcff']}/{CL(c)}{prj.rows['rev']}" for c in range(4, 11)], fmt=F_P)

prj.section("B. Projected income statement (pre-IFRS 16 equivalent: lease payments in operating costs)", 11)
prj.header(["", ""] + PY[:-1] + ["", "Comment"])
prj.line("p_ebit", "EBIT after leases", [None] + [f"={CL(c)}{prj.rows['ebit']}" for c in range(4, 10)])
prj.line("p_int", "Interest on financial debt", [None] + [f"=-{CL(c-1)}{prj.r+9}*WACC!$C$KD" for c in range(4, 10)], note="Effective cost of debt × opening gross debt (no circularity)")
prj.line("p_inc", "Interest income on cash", [None] + [f"={CL(c-1)}{prj.r+9}*{I['cash_yield']}" for c in range(4, 10)], note="Opening cash × yield")
prj.line("p_ebt", "Profit before tax", [None] + [f"=SUM({CL(c)}{prj.rows['p_ebit']}:{CL(c)}{prj.rows['p_inc']})" for c in range(4, 10)])
prj.line("p_tax", "Income tax", [None] + [f"=-{CL(c)}{prj.rows['p_ebt']}*{CL(c)}{prj.rows['taxr']}" for c in range(4, 10)], note="FX gains/losses assumed nil")
prj.line("p_ni_t", "Net income incl. NCI", [None] + [f"={CL(c)}{prj.rows['p_ebt']}+{CL(c)}{prj.rows['p_tax']}" for c in range(4, 10)])
prj.line("p_nci", "Non-controlling interests", [None] + [f"=-{CL(c)}{prj.rows['p_ni_t']}*{I['nci_pct']}" for c in range(4, 10)])
prj.line("p_ni", "Net income attributable", [None] + [f"={CL(c)}{prj.rows['p_ni_t']}+{CL(c)}{prj.rows['p_nci']}" for c in range(4, 10)], bold=True)
prj.line("p_eps", "  EPS (MXN, on 30-Sep-26 share count)", [None] + [f"={CL(c)}{prj.rows['p_ni']}/{I['shares']}" for c in range(4, 10)], fmt=F_N2)

prj.section("C. Projected balance sheet (31-Dec)", 11)
prj.header(["", ""] + PY[:-1] + ["", "Comment"])
b_cash = prj.r
prj.line("b_cash", "Cash & equivalents", [f"={HD('cash')}"] + [f"={CL(c-1)}{prj.r}+{CL(c)}{{CF}}" for c in range(4, 10)])
prj.line("b_nwc", "Net working capital (receivables + inventories − payables)", [f"={CL(c)}{prj.rows['nwc']}" for c in range(3, 10)])
prj.line("b_oca", "Other current assets", [f"={HD('oca')}"] + [f"={CL(c-1)}{prj.r}" for c in range(4, 10)], note="Held flat")
prj.line("b_ppe", "PP&E and intangibles, net", [f"={HD('ppe')}+{HD('intang')}"] + [f"={CL(c-1)}{prj.r}-{CL(c)}{prj.rows['capex']}+{CL(c)}{prj.rows['da']}" for c in range(4, 10)], note="Opening + capex − D&A")
prj.line("b_rou", "Right-of-use assets", [f"={HD('rou')}"] + [f"={CL(c-1)}{prj.r}" for c in range(4, 10)], note="Held flat (renewals = depreciation)")
prj.line("b_gw", "Goodwill", [f"={HD('gw')}"] + [f"={CL(c-1)}{prj.r}" for c in range(4, 10)], note="No acquisitions in base case")
prj.line("b_onca", "Other non-current assets", [f"={HD('onca')}"] + [f"={CL(c-1)}{prj.r}" for c in range(4, 10)], note="Held flat")
prj.line("b_ta", "Total assets", [f"=SUM({CL(c)}{prj.rows['b_cash']}:{CL(c)}{prj.rows['b_onca']})" for c in range(3, 10)], bold=True)
prj.line("b_debt", "Gross financial debt", [f"={HD('std')}+{HD('ltd')}"] + [f"={CL(c-1)}{prj.r}" for c in range(4, 10)], note="Held flat (refinanced at maturity)")
prj.line("b_lease", "Lease liabilities", [f"={HD('stl')}+{HD('ltl')}"] + [f"={CL(c-1)}{prj.r}" for c in range(4, 10)])
prj.line("b_ocl", "Other current liabilities", [f"={HD('ocl')}"] + [f"={CL(c-1)}{prj.r}" for c in range(4, 10)])
prj.line("b_oncl", "Other non-current liabilities", [f"={HD('oncl')}"] + [f"={CL(c-1)}{prj.r}" for c in range(4, 10)])
prj.line("b_eqp", "Equity attributable to AMX", [f"={HD('eqp')}"] + [f"={CL(c-1)}{prj.r}+{CL(c)}{prj.rows['p_ni']}+{CL(c)}{{DIV}}+{CL(c)}{{BUY}}" for c in range(4, 10)])
prj.line("b_nci", "Non-controlling interests", [f"={HD('nci_bs')}"] + [f"={CL(c-1)}{prj.r}-{CL(c)}{prj.rows['p_nci']}+{CL(c)}{{NCID}}" for c in range(4, 10)])
prj.line("b_tle", "Total liabilities & equity", [f"=SUM({CL(c)}{prj.rows['b_debt']}:{CL(c)}{prj.rows['b_nci']})" for c in range(3, 10)], bold=True)
prj.line("b_chk", "Balance check (assets − L&E)", [f"={CL(c)}{prj.rows['b_ta']}-{CL(c)}{prj.rows['b_tle']}" for c in range(3, 10)], fmt=F_M1)

prj.section("D. Projected cash flow statement and sources & uses", 11)
prj.header(["", ""] + PY[:-1] + ["", "Comment"])
prj.line("c_ni", "Net income incl. NCI", [None] + [f"={CL(c)}{prj.rows['p_ni_t']}" for c in range(4, 10)])
prj.line("c_da", "D&A", [None] + [f"=-{CL(c)}{prj.rows['da']}" for c in range(4, 10)])
prj.line("c_nwc", "Change in NWC", [None] + [f"={CL(c)}{prj.rows['dnwc']}" for c in range(4, 10)])
prj.line("c_cfo", "Cash from operations", [None] + [f"=SUM({CL(c)}{prj.rows['c_ni']}:{CL(c)}{prj.rows['c_nwc']})" for c in range(4, 10)], bold=True)
prj.line("c_capex", "Capex", [None] + [f"={CL(c)}{prj.rows['capex']}" for c in range(4, 10)])
prj.line("c_div", "Dividends to AMX shareholders", [None] + [f"=-{CL(c)}{prj.rows['p_ni']}*{I['payout']}" for c in range(4, 10)])
prj.line("c_buy", "Share buybacks", [None] + [f"=-{CL(c)}{prj.rows['p_ni']}*{I['buyback']}" for c in range(4, 10)])
prj.line("c_ncid", "Distributions to NCI (= NCI income)", [None] + [f"={CL(c)}{prj.rows['p_nci']}" for c in range(4, 10)])
prj.line("c_debt", "Net debt issuance / (repayment)", [None] + [f"={CL(c)}{prj.rows['b_debt']}-{CL(c-1)}{prj.rows['b_debt']}" for c in range(4, 10)])
prj.line("c_net", "Net change in cash", [None] + [f"={CL(c)}{prj.rows['c_cfo']}+SUM({CL(c)}{prj.rows['c_capex']}:{CL(c)}{prj.rows['c_debt']})" for c in range(4, 10)], bold=True)
prj.line("su_src", "Sources: CFO + new debt", [None] + [f"={CL(c)}{prj.rows['c_cfo']}+MAX({CL(c)}{prj.rows['c_debt']},0)" for c in range(4, 10)])
prj.line("su_use", "Uses: capex + distributions + debt repaid + increase in cash", [None] + [f"=-{CL(c)}{prj.rows['c_capex']}-{CL(c)}{prj.rows['c_div']}-{CL(c)}{prj.rows['c_buy']}-{CL(c)}{prj.rows['c_ncid']}-MIN({CL(c)}{prj.rows['c_debt']},0)+{CL(c)}{prj.rows['c_net']}" for c in range(4, 10)])
prj.line("su_chk", "Sources − uses (must be 0)", [None] + [f"={CL(c)}{prj.rows['su_src']}-{CL(c)}{prj.rows['su_use']}" for c in range(4, 10)], fmt=F_M1)
prj.line("nd_ebitdaal", "Memo: net debt ex-leases / EBITDAaL", [None] + [f"=({CL(c)}{prj.rows['b_debt']}-{CL(c)}{prj.rows['b_cash']}-{CL(c)}{prj.rows['b_oca']}*0)/{CL(c)}{prj.rows['ebitdaal']}" for c in range(4, 10)], fmt=F_X,
         note="Cash grows because debt is held flat; excess cash would fund buybacks or deleveraging (value-neutral in FCFF)")
# patch placeholders
for c in range(4, 10):
    cc = CL(c)
    prj.ws.cell(prj.rows["b_cash"], c).value = f"={CL(c-1)}{prj.rows['b_cash']}+{cc}{prj.rows['c_net']}"
    prj.ws.cell(prj.rows["b_eqp"], c).value = f"={CL(c-1)}{prj.rows['b_eqp']}+{cc}{prj.rows['p_ni']}+{cc}{prj.rows['c_div']}+{cc}{prj.rows['c_buy']}"
    prj.ws.cell(prj.rows["b_nci"], c).value = f"={CL(c-1)}{prj.rows['b_nci']}-{cc}{prj.rows['p_nci']}+{cc}{prj.rows['c_ncid']}"
    prj.ws.cell(prj.rows["p_int"], c).value = f"=-{CL(c-1)}{prj.rows['b_debt']}*WACC!$C$KD"
    prj.ws.cell(prj.rows["p_inc"], c).value = f"={CL(c-1)}{prj.rows['b_cash']}*{I['cash_yield']}"
    prj.ws.cell(prj.rows["nd_ebitdaal"], c).value = f"=({cc}{prj.rows['b_debt']}-{cc}{prj.rows['b_cash']})/{cc}{prj.rows['ebitdaal']}"
prj.ws.freeze_panes = "C3"

# =========================================================================
# WACC
# =========================================================================
wac = S(wb, "WACC", "WACC — built, not copied (USD build converted to MXN with inflation differential)", [62, 6, 16, 70])
W = {}


def wl(name, label, formula, fmt=F_P2, note=None, bold=False):
    r = wac.r
    wac.rows[name] = r
    wac.put(r, 1, label, font=BOLD if bold else BLACK)
    wac.put(r, 3, formula, fmt=fmt, bold=bold)
    if note:
        wac.put(r, 4, note, font=NOTE)
    W[name] = f"WACC!$C${r}"
    wac.r += 1


wac.section("Cost of equity", 4)
wl("rf", "Risk-free rate (US 10Y, USD)", f"={I['ust10']}", note="Inputs §1")
wl("bu", "Unlevered beta (avg. of EM Telecom Wireless & EM Telecom Services, cash-corrected)", f"=AVERAGE({I['bu_wireless']},{I['bu_services']})", F_N4, "Damodaran Jan-2026 industry betas")
wl("mcap", "Market capitalisation (MXN m)", f"={I['price']}*{I['shares']}", F_M)
wl("nd_target", "Target net debt = target ND/EBITDAaL × LTM EBITDAaL (MXN m)", f"={I['target_lev']}*{I['ebitdaal_ltm']}", F_M, "Company leverage target, not today's balance sheet")
wl("de", "Target D/E", f"={W['nd_target']}/{W['mcap']}", F_P)
wl("bl", "Relevered beta = βu × (1 + (1 − t) × D/E)", f"={W['bu']}*(1+(1-{I['tax_marg']})*{W['de']})", F_N4, "Hamada")
wl("erp", "Mature-market ERP", f"={I['erp']}")
crp_formula = "=" + "+".join(f"Drivers!$C${drv.rows[f'w_{i}']}*{RC[('crp', i)]}" for i in range(NR))
wl("crp", "Country risk premium, weighted by 9M26 revenue mix", crp_formula, note="Mexico 36%, Brazil 19%, Austria/EE 12%, CAC 9%, Colombia 9%, Argentina 5% …")
wl("ke_usd", "Cost of equity (USD) = Rf + βL × ERP + λ × CRP", f"={W['rf']}+{W['bl']}*{W['erp']}+{I['crp_lambda']}*{W['crp']}")
wl("infl_mx", "Long-run Mexican inflation", f"={RC[('infl', 0)]}", note="Banxico target")
wl("ke", "Cost of equity (MXN) = (1 + Ke USD) × (1 + π MX)/(1 + π US) − 1", f"=(1+{W['ke_usd']})*(1+{W['infl_mx']})/(1+{I['infl_us']})-1", bold=True)
wac.section("Cost of debt", 4)
wl("kd", "Pre-tax cost of debt = FY25 interest on debt / average gross debt", f"={I['int_debt_25']}/AVERAGE({I['debt_24']},{I['debt_25']})", note="Effective rate on AMX's actual (multi-currency) debt; ex-lease interest")
wl("kd_chk", "Cross-check: (UST + assumed spread) converted to MXN", f"=(1+{I['ust10']}+{I['kd_spread']})*(1+{W['infl_mx']})/(1+{I['infl_us']})-1", note="Effective rate is above the market-implied rate → conservative")
wl("kd_at", "After-tax cost of debt", f"={W['kd']}*(1-{I['tax_marg']})", bold=True)
wac.section("Weights and WACC", 4)
wl("wd", "D / (D + E) at target", f"={W['de']}/(1+{W['de']})", F_P)
wl("we", "E / (D + E) at target", f"=1-{W['wd']}", F_P)
wl("wacc", "WACC (nominal MXN)", f"={W['we']}*{W['ke']}+{W['wd']}*{W['kd_at']}", bold=True)
wac.ws.cell(wac.rows["wacc"], 3).fill = FILL_KEY
wac.section("Defence question (c): country risk +100 bp", 4)
wl("ke_shock", "Cost of equity (MXN) with CRP + shock", f"=(1+{W['ke_usd']}+{I['crp_lambda']}*{I['crp_shock']})*(1+{W['infl_mx']})/(1+{I['infl_us']})-1")
wl("wacc_shock", "WACC with CRP + shock", f"={W['we']}*{W['ke_shock']}+{W['wd']}*{W['kd_at']}")
wl("wacc_delta", "Change in WACC", f"={W['wacc_shock']}-{W['wacc']}")
# patch KD placeholder used in projections
for c in range(4, 10):
    for key in ("p_int",):
        cell = prj.ws.cell(prj.rows[key], c)
        cell.value = cell.value.replace("WACC!$C$KD", W["kd"])

# =========================================================================
# DCF
# =========================================================================
dcf = S(wb, "DCF", "DCF — FCFF discounted at WACC, Gordon terminal value, equity bridge", [58, 6, 14, 14, 14, 14, 14, 14, 16, 60])
dcf.section("Cash flows and discounting", 10)
dcf.header(["", "", "4Q26E stub", "FY2027E", "FY2028E", "FY2029E", "FY2030E", "FY2031E", "Terminal", "Comment"])
PJ = lambda key, c: f"Projections!{CL(c)}{prj.rows[key]}"
dcf.line("fcff", "FCFF", [f"={PJ('fcff', 4)}*{I['stub']}"] + [f"={PJ('fcff', c)}" for c in range(5, 10)] + [f"={PJ('fcff', 10)}"], bold=True, note="Stub = 2026E FCFF × stub fraction")
dcf.line("rev", "Revenue (for margin sensitivity)", [f"={PJ('rev', 4)}*{I['stub']}"] + [f"={PJ('rev', c)}" for c in range(5, 10)] + [f"={PJ('rev', 10)}"])
dcf.line("taxr", "Operating tax rate", [f"={PJ('taxr', 4)}"] + [f"={PJ('taxr', c)}" for c in range(5, 10)] + [f"={PJ('taxr', 10)}"], fmt=F_P)
dcf.line("t", "Discount time (years from valuation date)", [f"={I['stub_t']}"] + [f"=({I['fy1_end']}-{I['val_date']})/{I['days']}+Drivers!{CL(c+2)}{drv.rows['year']}-{I['mid']}" for c in range(4, 9)] +
         [f"=H{dcf.r}+{I['mid']}"], fmt=F_N2, note="Mid-year convention; terminal value discounted from end-2031")
dcf.line("df", "Discount factor", [f"=1/(1+{W['wacc']})^{CL(c)}{dcf.rows['t']}" for c in range(3, 10)], fmt=F_N4)
dcf.line("tv", "Terminal value at end-2031 = FCFF(T) / (WACC − g)", [None] * 6 + [f"=I{dcf.rows['fcff']}/({W['wacc']}-{I['g']})"], note="FCFF(T) already grown by g in Projections")
dcf.line("pv", "Present value", [f"={CL(c)}{dcf.rows['fcff']}*{CL(c)}{dcf.rows['df']}" for c in range(3, 9)] + [f"=I{dcf.rows['tv']}*I{dcf.rows['df']}"], bold=True)
dcf.section("Enterprise value and terminal-value dependence", 10)
V = {}


def dl(name, label, formula, fmt=F_M, note=None, bold=False, fill=None):
    r = dcf.r
    dcf.rows[name] = r
    dcf.put(r, 1, label, font=BOLD if bold else BLACK)
    dcf.put(r, 3, formula, fmt=fmt, bold=bold, fill=fill)
    if note:
        dcf.put(r, 4, note, font=NOTE)
    V[name] = f"DCF!$C${r}"
    dcf.r += 1


dl("pv_fcff", "PV of explicit FCFF (4Q26–2031)", f"=SUM(C{dcf.rows['pv']}:H{dcf.rows['pv']})")
dl("pv_tv", "PV of terminal value", f"=I{dcf.rows['pv']}")
dl("ev", "Enterprise value", f"={V['pv_fcff']}+{V['pv_tv']}", bold=True)
dl("tv_pct", "Terminal value as % of EV", f"={V['pv_tv']}/{V['ev']}", F_P, bold=True)
dl("tv_flag", "Terminal-value test", f'=IF({V["tv_pct"]}>{I["tv_limit"]},"WEAK CASE: terminal value exceeds "&TEXT({I["tv_limit"]},"0%")&" of EV","Pass: terminal value below "&TEXT({I["tv_limit"]},"0%")&" of EV")', None)
dl("exit_mult", "Implied exit multiple: TV / 2031E EBITDA", f"=I{dcf.rows['tv']}/{PJ('ebitda', 9)}", F_X, "Sanity check vs. comps")
dl("ev_ebitda26", "Implied EV / 2026E EBITDA", f"={V['ev']}/{PJ('ebitda', 4)}", F_X)
dcf.section("Equity bridge (MXN m)", 10)
dl("nd_rep", "Net financial debt ex-leases, 30-Sep-26 (reported)", f"={I['debt_total']}-{I['cash_sec']}", note="6-K p.37: 400,081")
SPOT = f"({I['usdmxn']}/(1+{I['fx_shock']}))"
remark = (f"={I['debt_usd']}*({SPOT}-{I['usdmxn_sep']})+{I['debt_eur']}*({SPOT}*{I['eurusd']}-{I['usdmxn_sep']}/{I['eurusd_sep']})"
          f"+{I['debt_gbp']}*({SPOT}*{I['gbpusd']}-{I['usdmxn_sep']}*{I['gbpusd_sep']})+{I['debt_brl']}*({SPOT}/{I['usdbrl']}-1/{I['brlmxn_sep']})")
dl("nd_fx", "FX re-mark of foreign-currency debt to spot", remark, note="Peso weaker since 30-Sep raises MXN value of USD/EUR/GBP debt; hedges ignored (conservative)")
dl("div", "Dividend paid after 30-Sep", f"={I['div_post']}")
dl("pension", "Labour obligations (pension deficit), after tax", f"=({I['pension_dec25']}-{I['pension_paydown']})*(1-{I['pension_tax']})")
dl("nci", "Non-controlling interests", f"={I['nci']}")
dl("aff", "Less: investments in affiliates", f"=-{I['affiliates']}")
dl("desk", "Desktop acquisition cost (only if toggle = 1)", f"={I['desk_on']}*({I['desk_shares']}*{I['desk_px']}+{I['desk_netdebt']})*{I['usdmxn']}/{I['usdbrl']}", note="100% equity at BRL 20.63 (incl. tender) + net debt")
dl("bridge", "Total claims deducted from EV", f"=SUM({V['nd_rep']}:{V['desk']})".replace("DCF!$C$", "C").replace("$", ""), bold=True)
dcf.ws.cell(dcf.rows["bridge"], 3).value = f"=SUM(C{dcf.rows['nd_rep']}:C{dcf.rows['desk']})"
dl("eqv", "Equity value", f"={V['ev']}-{V['bridge']}", bold=True)
dl("vps", "Value per share (MXN)", f"={V['eqv']}/{I['shares']}", F_N2, bold=True, fill=FILL_KEY)
dl("vadr", "Value per ADR (USD)", f"={V['vps']}*{I['shares_per_adr']}/{I['usdmxn']}", F_N2)
dl("px", "Share price (MXN)", f"={I['price']}", F_N2)
dl("upside", "Upside / (downside) vs. price", f"={V['vps']}/{V['px']}-1", F_P, bold=True)
dl("reco", "Recommendation (rule in Inputs §6)", f'=IF({V["upside"]}>={I["buy_th"]},"BUY",IF({V["upside"]}<={I["sell_th"]},"SELL","HOLD"))', None, bold=True)
dl("impl_px_mult", "Memo: EV / LTM-run-rate EBITDA at market price", f"=({W['mcap']}+{V['nd_rep']}+{I['lease_liab']}+{I['nci']})/({I['ebitda_3q']}*{I['q_per_y']})", F_X, "IFRS 16 basis, same as Comps")

# =========================================================================
# SENSITIVITY
# =========================================================================
sen = S(wb, "Sensitivity", "Sensitivity — value per share (MXN) vs. WACC, terminal growth and EBITDA margin (full grids, live formulas)",
        [40, 12, 12, 12, 12, 12, 12, 12, 12, 50])
fc = f"DCF!$D${dcf.rows['fcff']}:$H${dcf.rows['fcff']}"
rv = f"DCF!$D${dcf.rows['rev']}:$H${dcf.rows['rev']}"
tx = f"DCF!$D${dcf.rows['taxr']}:$H${dcf.rows['taxr']}"
tt = f"DCF!$D${dcf.rows['t']}:$H${dcf.rows['t']}"
s_f, s_r, s_x, s_t = (f"DCF!$C${dcf.rows[k]}" for k in ("fcff", "rev", "taxr", "t"))
T_f, T_r, T_x, T_t = (f"DCF!$I${dcf.rows[k]}" for k in ("fcff", "rev", "taxr", "t"))
BR = V["bridge"]
SH = I["shares"]


def vps_formula(w, gg, dm, dc="0"):
    """Value per share as a closed-form function of WACC, g, margin delta (dm) and capex delta (dc).
    Margin delta raises EBITDA by dm × revenue (taxed); capex delta lowers FCFF by dc × revenue."""
    adj = f"({dm})*{{R}}*(1-{{X}})-({dc})*{{R}}"
    stub = f"({s_f}+" + adj.format(R=s_r, X=s_x) + f")/(1+{w})^{s_t}"
    exp = f"SUMPRODUCT(({fc}+({dm})*{rv}*(1-{tx})-({dc})*{rv})/(1+{w})^{tt})"
    term = f"({T_f}+" + adj.format(R=T_r, X=T_x) + f")/(({w})-({gg}))/(1+{w})^{T_t}"
    return f"=({stub}+{exp}+{term}-{BR})/{SH}"


def grid(title, rlab, clab, rbase, rstep, cbase, cstep, fn, rfmt=F_P2, cfmt=F_P2, n_r=7, n_c=7):
    sen.section(title, 10)
    top = sen.r
    sen.put(top, 1, f"Rows: {rlab}  |  Columns: {clab}", font=NOTE)
    ccen = CL(3 + n_c // 2)
    rcen = top + 1 + n_r // 2
    for j in range(n_c):
        sen.put(top, 2 + j + 1, f"={cbase}+(COLUMN()-COLUMN(${ccen}${top}))*{cstep}", fmt=cfmt, bold=True, fill=FILL_S)
    for i in range(n_r):
        rr = top + 1 + i
        sen.put(rr, 2, f"={rbase}+(ROW()-ROW($B${rcen}))*{rstep}", fmt=rfmt, bold=True, fill=FILL_S)
        for j in range(n_c):
            rc = f"$B{rr}"
            ccell = f"{CL(3 + j)}${top}"
            cell = sen.put(rr, 3 + j, fn(rc, ccell), fmt=F_N2)
            if i == n_r // 2 and j == n_c // 2:
                cell.fill = FILL_KEY
    sen.r = top + n_r + 2
    return top


sen.note("Each cell recomputes the full DCF (stub + 2027–31 FCFF + terminal value − equity bridge) in closed form. Centre cell = base case and equals DCF!C"
         f"{dcf.rows['vps']}. Margin delta is applied to every forecast year and the terminal year; taxes adjust.")
g1 = grid("Table 1 — WACC (rows) × terminal growth (columns); margin at base", "WACC", "terminal growth g", W["wacc"], I["step_w"], I["g"], I["step_g"],
          lambda r, c: vps_formula(r, c, "0"))
g2 = grid("Table 2 — WACC (rows) × EBITDA-margin change (columns, pp); g at base", "WACC", "EBITDA margin change", W["wacc"], I["step_w"], "0*" + I["step_m"], I["step_m"],
          lambda r, c: vps_formula(r, I["g"], c))
g3 = grid("Table 3 — terminal growth (rows) × EBITDA-margin change (columns, pp); WACC at base", "terminal growth g", "EBITDA margin change", I["g"], I["step_g"], "0*" + I["step_m"], I["step_m"],
          lambda r, c: vps_formula(W["wacc"], r, c))
sen.section("Table 4 — one-variable swings (which assumption moves value most?)", 10)
sen.header(["Driver", "Shock", "Value/share", "Δ vs base", "Δ %"], start_col=1)
base = f"DCF!$C${dcf.rows['vps']}"
swings = [
    ("WACC +100 bp", f"={I['sw_w']}", vps_formula(f"{W['wacc']}+{I['sw_w']}", I["g"], "0")),
    ("WACC −100 bp", f"=-{I['sw_w']}", vps_formula(f"{W['wacc']}-{I['sw_w']}", I["g"], "0")),
    ("Terminal growth +100 bp", f"={I['sw_g']}", vps_formula(W["wacc"], f"{I['g']}+{I['sw_g']}", "0")),
    ("Terminal growth −100 bp", f"=-{I['sw_g']}", vps_formula(W["wacc"], f"{I['g']}-{I['sw_g']}", "0")),
    ("EBITDA margin +200 bp (all years)", f"={I['sw_m']}", vps_formula(W["wacc"], I["g"], I['sw_m'])),
    ("EBITDA margin −200 bp (all years)", f"=-{I['sw_m']}", vps_formula(W["wacc"], I["g"], f"-{I['sw_m']}")),
    ("Capex +100 bp of revenue (all years)", f"={I['sw_c']}", vps_formula(W["wacc"], I["g"], "0", I['sw_c'])),
    ("Country risk +100 bp (via WACC)", f"={W['wacc_delta']}", vps_formula(W["wacc_shock"], I["g"], "0")),
]
SW = {}
for lab, shock, f in swings:
    r = sen.r
    SW[lab] = r
    sen.put(r, 1, lab)
    sen.put(r, 2, shock, fmt=F_P2)
    sen.put(r, 3, f, fmt=F_N2)
    sen.put(r, 4, f"=C{r}-{base}", fmt=F_N2)
    sen.put(r, 5, f"=C{r}/{base}-1", fmt=F_P)
    sen.r += 1
sen.r += 1
sen.put(sen.r, 1, "Sensitivity range (Table 1, WACC ±1 pp × g ±1 pp around base): min / max", font=BOLD)
sen.put(sen.r, 3, f"=MIN(D{g1+2}:H{g1+6})", fmt=F_N2)
sen.put(sen.r, 4, f"=MAX(D{g1+2}:H{g1+6})", fmt=F_N2)
sen.rows["range"] = sen.r
sen.r += 1

# =========================================================================
# COMPS
# =========================================================================
cmp_ = S(wb, "Comps", "Comparable companies — EV/EBITDA and EV/Sales (cross-check; the DCF leads)",
         [34, 10, 14, 14, 14, 14, 14, 14, 12, 12, 12, 12, 40])
cmp_.note("All peer values in local currency millions; multiples are currency-neutral. EV (IFRS16) = market cap + net financial debt + leases + NCI. "
          "EV ex-leases = market cap + net financial debt + NCI. Run-rate = latest reported quarter × 4. Data and sources: Inputs §8.")
cmp_.header(["Company", "Region", "Market cap", "EV (incl. leases)", "EV ex-leases", "EBITDA run-rate", "EBITDAaL run-rate", "Revenue run-rate",
             "EV/EBITDA", "EV/EBITDAaL", "EV/Sales", "Net debt/EBITDA", "Note"])
cr0 = cmp_.r
q = I["q_per_y"]
for pi, p in enumerate(PEERS):
    r = cmp_.r
    ir = PR[pi]
    A = lambda col: f"Inputs!${col}${ir}"
    mc = f"{A('B')}*{A('C')}" + (f"*{I['gbpusd']}/{I['eurusd']}" if "Vodafone" in p[0] else "")
    cmp_.put(r, 1, f"={A('A')}")
    cmp_.put(r, 2, f"=Inputs!$K${ir}")
    cmp_.put(r, 3, f"={mc}", fmt=F_M)
    cmp_.put(r, 4, f"=C{r}+{A('D')}+{A('E')}+{A('F')}", fmt=F_M)
    cmp_.put(r, 5, f"=C{r}+{A('D')}+{A('F')}", fmt=F_M)
    cmp_.put(r, 6, f'=IF(ISNUMBER({A("G")}),{A("G")}*{q},"n/a")', fmt=F_M)
    cmp_.put(r, 7, f'=IF(ISNUMBER({A("H")}),{A("H")}*{q},"n/a")', fmt=F_M)
    cmp_.put(r, 8, f"={A('I')}*{q}", fmt=F_M)
    cmp_.put(r, 9, f'=IF(ISNUMBER(F{r}),D{r}/F{r},"n/a")', fmt=F_X)
    cmp_.put(r, 10, f'=IF(ISNUMBER(G{r}),E{r}/G{r},"n/a")', fmt=F_X)
    cmp_.put(r, 11, f"=D{r}/H{r}", fmt=F_X)
    cmp_.put(r, 12, f'=IF(ISNUMBER(F{r}),({A("D")}+{A("E")})/F{r},"n/a")', fmt=F_X)
    cmp_.r += 1
cr1 = cmp_.r - 1
cmp_.r += 1
stats = [("Median — LatAm", (cr0, cr0 + 4)), ("Median — Europe", (cr0 + 5, cr1)), ("Median — all peers", (cr0, cr1))]
ST = {}
for lab, (a_, b_) in stats:
    r = cmp_.r
    ST[lab] = r
    cmp_.put(r, 1, lab, font=BOLD)
    for col in (9, 10, 11):
        cmp_.put(r, col, f"=MEDIAN({CL(col)}{a_}:{CL(col)}{b_})", fmt=F_X, bold=True)
    cmp_.r += 1
for lab, fn in (("25th percentile — all", "QUARTILE"), ("75th percentile — all", "QUARTILE")):
    r = cmp_.r
    ST[lab] = r
    cmp_.put(r, 1, lab, font=BOLD)
    k = I["pct_lo"] if lab.startswith("25") else I["pct_hi"]
    for col in (9, 10, 11):
        cmp_.put(r, col, f"=PERCENTILE({CL(col)}{cr0}:{CL(col)}{cr1},{k})", fmt=F_X, bold=True)
    cmp_.r += 1
cmp_.r += 1
r = cmp_.r
cmp_.rows["amx"] = r
cmp_.put(r, 1, "América Móvil (AMX) — at market", font=BOLD, fill=FILL_KEY)
cmp_.put(r, 2, "LatAm/Europe")
cmp_.put(r, 3, f"={W['mcap']}", fmt=F_M)
cmp_.put(r, 4, f"=C{r}+{V['nd_rep']}+{I['lease_liab']}+{I['nci']}", fmt=F_M)
cmp_.put(r, 5, f"=C{r}+{V['nd_rep']}+{I['nci']}", fmt=F_M)
cmp_.put(r, 6, f"={I['ebitda_3q']}*{q}", fmt=F_M)
cmp_.put(r, 7, f"={I['ebitdaal_9m']}/({I['months_y']}-{I['months_q']})*{I['months_y']}", fmt=F_M)
cmp_.put(r, 8, f"={I['rev_3q']}*{q}", fmt=F_M)
cmp_.put(r, 9, f"=D{r}/F{r}", fmt=F_X, bold=True)
cmp_.put(r, 10, f"=E{r}/G{r}", fmt=F_X, bold=True)
cmp_.put(r, 11, f"=D{r}/H{r}", fmt=F_X, bold=True)
cmp_.put(r, 12, f"=({V['nd_rep']}+{I['lease_liab']})/F{r}", fmt=F_X)
cmp_.put(r, 13, "EBITDAaL run-rate = 9M26 EBITDAaL 241,238 × 4/3 (6-K p.15)", font=NOTE)
cmp_.r += 1
r = cmp_.r
cmp_.rows["amx_dcf"] = r
cmp_.put(r, 1, "AMX — at DCF value", font=BOLD)
cmp_.put(r, 3, f"={V['eqv']}", fmt=F_M)
cmp_.put(r, 4, f"=C{r}+(D{r-1}-C{r-1})", fmt=F_M)
cmp_.put(r, 5, f"=C{r}+(E{r-1}-C{r-1})", fmt=F_M)
cmp_.put(r, 9, f"=D{r}/F{r-1}", fmt=F_X)
cmp_.put(r, 10, f"=E{r}/G{r-1}", fmt=F_X)
cmp_.put(r, 11, f"=D{r}/H{r-1}", fmt=F_X)
cmp_.put(r, 13, "Note: AMX's pension deficit is deducted in the DCF but not in market EVs — part of why AMX screens cheaper on comps than its DCF value implies", font=NOTE)
cmp_.r += 2
cmp_.section("Implied AMX value per share from peer multiples (MXN)", 13)
cmp_.header(["Multiple applied to AMX run-rate", "", "Multiple", "Implied EV", "Less claims", "Equity", "Per share", "vs price"])
amx = cmp_.rows["amx"]
imp = [("EV/EBITDA — LatAm median", f"I{ST['Median — LatAm']}", f"F{amx}", f"({V['nd_rep']}+{I['lease_liab']}+{I['nci']})"),
       ("EV/EBITDA — all-peer median", f"I{ST['Median — all peers']}", f"F{amx}", f"({V['nd_rep']}+{I['lease_liab']}+{I['nci']})"),
       ("EV/EBITDA — 25th percentile", f"I{ST['25th percentile — all']}", f"F{amx}", f"({V['nd_rep']}+{I['lease_liab']}+{I['nci']})"),
       ("EV/EBITDA — 75th percentile", f"I{ST['75th percentile — all']}", f"F{amx}", f"({V['nd_rep']}+{I['lease_liab']}+{I['nci']})"),
       ("EV/EBITDAaL — all-peer median", f"J{ST['Median — all peers']}", f"G{amx}", f"({V['nd_rep']}+{I['nci']})"),
       ("EV/Sales — all-peer median", f"K{ST['Median — all peers']}", f"H{amx}", f"({V['nd_rep']}+{I['lease_liab']}+{I['nci']})")]
for lab, m, base_, claims in imp:
    r = cmp_.r
    cmp_.rows[lab] = r
    cmp_.put(r, 1, lab)
    cmp_.put(r, 3, f"={m}", fmt=F_X)
    cmp_.put(r, 4, f"=C{r}*{base_}", fmt=F_M)
    cmp_.put(r, 5, f"={claims}", fmt=F_M)
    cmp_.put(r, 6, f"=D{r}-E{r}", fmt=F_M)
    cmp_.put(r, 7, f"=F{r}/{I['shares']}", fmt=F_N2)
    cmp_.put(r, 8, f"=G{r}/{I['price']}-1", fmt=F_P)
    cmp_.r += 1
cmp_.r += 1
cmp_.section("Precedent transaction (cross-check)", 13)
cmp_.header(["Deal", "Date", "EV (BRL m)", "LTM EBITDA", "EV/EBITDA", "LTM revenue", "EV/Sales", "", "", "", "", "", "Source"])
r = cmp_.r
cmp_.rows["desk"] = r
cmp_.put(r, 1, "Claro NXT (AMX) / Desktop — 72.83% control")
cmp_.put(r, 2, "1-Oct-26")
cmp_.put(r, 3, f"={I['desk_shares']}*{I['desk_px']}+{I['desk_netdebt']}", fmt=F_M)
cmp_.put(r, 4, f"={I['desk_ebitda']}", fmt=F_M)
cmp_.put(r, 5, f"=C{r}/D{r}", fmt=F_X)
cmp_.put(r, 6, f"={I['desk_rev']}", fmt=F_M)
cmp_.put(r, 7, f"=C{r}/F{r}", fmt=F_X)
cmp_.put(r, 13, "Price: Claro fato relevante 1-Oct-26 (primary). Net debt/EBITDA/revenue: Yahoo (third party)", font=NOTE)
cmp_.r += 1
cmp_.note("Other recent LatAm deals (Millicom/Telefónica Colombia, Ecuador, Uruguay 2025–26; Telecom Argentina/Telefónica Argentina 2025) were not used: "
          "their EV/EBITDA could not be verified from primary documents within this project. Listing them without verified multiples would be invented precision.")
# fix AMX EBITDAaL formula (clean)


# =========================================================================
# CHECKS
# =========================================================================
chk = S(wb, "Checks", "Checks — integrity tests (0 / TRUE = pass)", [74, 16, 16, 60])
chk.header(["Test", "Value", "Pass?", "What it tests"])
CK = []


def ck(label, formula, test, what):
    r = chk.r
    chk.put(r, 1, label)
    chk.put(r, 2, formula, fmt=F_M1)
    chk.put(r, 3, test.replace("{v}", f"B{r}"))
    chk.put(r, 4, what, font=NOTE)
    CK.append(r)
    chk.r += 1


tol = "0.5"
for j, y in enumerate(["FY2021", "FY2022", "FY2023", "FY2024", "FY2025"]):
    c = CL(3 + j)
    ck(f"Historical BS balances {y}: assets − (liabilities + equity)", f"=Historicals!{c}{his.rows['assets']}-Historicals!{c}{his.rows['le']}", f"=ABS({{v}})<{tol}", "Reported balance sheet ties")
    ck(f"Historical BS {y}: total assets vs reported", f"=Historicals!{c}{his.rows['assets']}-Historicals!{c}{his.rows['assets_rep']}", f"=ABS({{v}})<{tol}", "Aggregation loses nothing")
    ck(f"Historical CF {y}: closing cash vs balance sheet cash", f"=Historicals!{c}{his.rows['close_cash']}-Historicals!{c}{his.rows['cash']}", f"=ABS({{v}})<2", "Cash flow statement ties to balance sheet")
    ck(f"Historical IS {y}: EBIT rebuilt vs reported", f"=Historicals!{c}{his.rows['ebit']}-Historicals!{c}{his.rows['ebit_rep']}", f"=ABS({{v}})<2", "Income statement lines tie")
    ck(f"Historical IS {y}: net income rebuilt vs reported", f"=Historicals!{c}{his.rows['ni']}-Historicals!{c}{his.rows['ni_rep']}", f"=ABS({{v}})<2", "Net income ties")
    ck(f"Historical equity {y}: parent + NCI vs reported total", f"=Historicals!{c}{his.rows['eq']}-Historicals!{c}{his.rows['eq_rep']}", f"=ABS({{v}})<2", "Equity split ties")
ck("Sep-26 BS balances (6-K): assets − (liabilities + equity)", f"=Historicals!I{his.rows['assets']}-Historicals!I{his.rows['le']}", "=ABS({v})<3", "6-K balance sheet ties (rounding ≤ 2)")
ck("9M26 IS: EBIT rebuilt vs 6-K", f"=Historicals!I{his.rows['ebit']}-Historicals!I{his.rows['ebit_rep']}", f"=ABS({{v}})<2", "")
ck("9M26 IS: net income rebuilt vs 6-K", f"=Historicals!I{his.rows['ni']}-Historicals!I{his.rows['ni_rep']}", f"=ABS({{v}})<2", "")
ck("Sep-26 net debt vs 6-K (400,081)", f"=Historicals!I{his.rows['netdebt']}-400081", f"=ABS({{v}})<2", "Net debt definition matches company")
for j, y in enumerate(["FY2026E", "FY2027E", "FY2028E", "FY2029E", "FY2030E", "FY2031E"]):
    c = CL(4 + j)
    ck(f"Projected BS balances {y}", f"=Projections!{c}{prj.rows['b_chk']}", f"=ABS({{v}})<{tol}", "Projected three statements tie")
    ck(f"Sources = uses {y}", f"=Projections!{c}{prj.rows['su_chk']}", f"=ABS({{v}})<{tol}", "Cash sources equal uses")
ck("Projected BS opening (FY2025A) balances", f"=Projections!C{prj.rows['b_chk']}", f"=ABS({{v}})<{tol}", "Opening balance sheet ties")
for i in range(NR):
    ck(f"Driver calibration {REG[i]}: 3Q26 revenue rebuilt − reported", f"=Drivers!C{DR[('chk', i)]}", f"=ABS({{v}})<{tol}", "Subscribers × ARPU reproduces reported revenue")
ck("Consolidated 3Q26 revenue = sum of regions × (1 + elim)", f"=Drivers!C{drv.rows['sum_rev']}*(1+Drivers!C{drv.rows['elim_rev']})-{I['rev_3q']}", f"=ABS({{v}})<{tol}", "Regional build reconciles to consolidated")
ck("WACC > terminal growth", f"={W['wacc']}-{I['g']}", "={v}>0", "Gordon model valid")
ck("Sensitivity centre cell = DCF value per share", f"=Sensitivity!F{g1+4}-{V['vps']}", f"=ABS({{v}})<0.001", "Closed-form sensitivity reproduces the DCF")
ck("Terminal value share of EV (flag only, not a pass/fail)", f"={V['tv_pct']}", f"={{v}}<={I['tv_limit']}", "Mandate: say plainly if > 70%")
chk.ws.cell(CK[-1], 2).number_format = F_P
chk.ws.cell(CK[-2], 2).number_format = F_N4
chk.ws.cell(CK[-3], 2).number_format = F_P2
chk.r += 1
chk.put(chk.r, 1, "ALL CHECKS PASS?", font=BOLD)
chk.put(chk.r, 3, "=AND(" + ",".join(f"C{r}" for r in CK[:-1]) + ")", bold=True, fill=FILL_KEY)
chk.rows["all"] = chk.r
chk.r += 2
chk.note("Hard-code audit: performed by scripts/audit_model.py — scans every formula on Drivers, Projections, WACC, DCF, Sensitivity, Comps and Checks "
         "for numeric literals other than 0 and 1 (structural). Result is written to docs/model_audit.txt.")

# =========================================================================
# COVER content
# =========================================================================
cv = cover
cv.r = 3
cv.put(cv.r, 1, "Investment-banking interview preparation project. MXN millions unless stated. Valuation date: see Inputs.", font=NOTE); cv.r += 2
cv.header(["Output", "Value", "", "Comment"])
outs = [
    ("Value per share (DCF, MXN)", f"={V['vps']}", F_N2, "FCFF discounted at WACC; Gordon terminal value"),
    ("Value per ADR (USD)", f"={V['vadr']}", F_N2, ""),
    ("Share price (MXN), 24-Nov-2026", f"={V['px']}", F_N2, "BMV: AMXB"),
    ("Upside / (downside)", f"={V['upside']}", F_P, ""),
    ("Recommendation (pre-set rule: ±15%)", f"={V['reco']}", None, "Rule fixed before the model was run"),
    ("WACC (nominal MXN)", f"={W['wacc']}", F_P2, "See WACC sheet"),
    ("Terminal growth", f"={I['g']}", F_P, ""),
    ("Terminal value as % of EV", f"={V['tv_pct']}", F_P, ""),
    ("Terminal-value test", f"={V['tv_flag']}", None, ""),
    ("Enterprise value (MXN m)", f"={V['ev']}", F_M, ""),
    ("Equity value (MXN m)", f"={V['eqv']}", F_M, ""),
    ("Implied EV / 2026E EBITDA (DCF)", f"={V['ev_ebitda26']}", F_X, "EV here excludes leases (EBITDAaL-based FCFF)"),
    ("Sensitivity range, WACC ±1pp × g ±1pp (MXN)", f'=TEXT(Sensitivity!C{sen.rows["range"]},"0.00")&" – "&TEXT(Sensitivity!D{sen.rows["range"]},"0.00")', None, ""),
    ("Comps: implied value/share at all-peer median EV/EBITDA (MXN)", f"=Comps!G{cmp_.rows['EV/EBITDA — all-peer median']}", F_N2, "Cross-check only"),
    ("Desktop included?", f'=IF({I["desk_on"]}=1,"Yes","No — excluded from base case")', None, "See memo"),
    ("All integrity checks pass?", f"=Checks!C{chk.rows['all']}", None, "See Checks sheet"),
]
for lab, f, fmt, com in outs:
    cv.put(cv.r, 1, lab)
    cv.put(cv.r, 2, f, fmt=fmt, bold=True, fill=FILL_KEY)
    cv.put(cv.r, 4, com, font=NOTE)
    cv.r += 1
cv.r += 1
cv.header(["Sheet", "Contents", "", ""])
for s_, d_ in [("Inputs", "Every assumption with source and rationale; regional base data; peer data"),
               ("Historicals", "IS / BS / CF FY2021–FY2025 (20-F XBRL) + 9M25/9M26 (6-K)"),
               ("Drivers", "Country build: postpaid subs × ARPU, prepaid revenue, broadband × fixed ARPU, margins, FX"),
               ("Projections", "Consolidated FCFF + projected IS/BS/CF with sources & uses"),
               ("WACC", "Rf, relevered beta, operations-weighted CRP, cost of debt from actual interest"),
               ("DCF", "Discounting, terminal value, TV %, equity bridge, recommendation"),
               ("Comps", "LatAm + European peers, EV/EBITDA, EV/EBITDAaL, EV/Sales, precedent"),
               ("Sensitivity", "Full grids: WACC × g, WACC × margin, g × margin; one-variable swings"),
               ("Checks", "Statement ties, sources = uses, calibration, sensitivity reproduction")]:
    cv.put(cv.r, 1, s_, font=BOLD)
    cv.put(cv.r, 2, d_)
    cv.r += 1
cv.r += 1
cv.note("Colour code: blue = hard-coded input (Inputs sheet / reported data on Historicals); black = formula; green = cross-sheet link. Yellow fill = input cell.")
cv.note("Not investment advice. Educational valuation exercise built from public filings.")

for ws in wb.worksheets:
    ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
wb.save(OUT)
print("saved", OUT)

# map of input names -> cells (used by scripts/scenarios.py)
json.dump({"scalar": I, "regional": {f"{k[0]}|{k[1]}": v for k, v in RC.items()}},
          open(os.path.join(ROOT, "data", "input_map.json"), "w"), indent=1)
