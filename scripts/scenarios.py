"""Re-run the full workbook under alternative inputs (LibreOffice recalculation) and report value per share.
Usage: python scripts/scenarios.py <scratch_dir>  -> docs/scenarios.md"""
import json, os, sys, subprocess, shutil
from openpyxl import load_workbook
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
MODEL = os.path.join(ROOT, "model", "AMX_Valuation_Model.xlsx")
MAP = json.load(open(os.path.join(ROOT, "data", "input_map.json")))
TMP = sys.argv[1]
os.makedirs(TMP, exist_ok=True)

def cell(ref):
    sh, a = ref.split("!")
    return sh.strip("'"), a.replace("$", "")

def run(name, changes):
    wb = load_workbook(MODEL)
    for key, val in changes.items():
        if "|" in key:  # regional: name|* or name|i
            base, idx = key.split("|")
            idxs = range(9) if idx == "*" else [int(idx)]
            for i in idxs:
                sh, a = cell(MAP["regional"][f"{base}|{i}"])
                wb[sh][a].value = val(wb[sh][a].value) if callable(val) else val
        else:
            sh, a = cell(MAP["scalar"][key])
            wb[sh][a].value = val(wb[sh][a].value) if callable(val) else val
    f = os.path.join(TMP, f"scn_{len(os.listdir(TMP))}.xlsx")
    wb.save(f)
    out = os.path.join(TMP, "out")
    subprocess.run(["soffice", "--headless", "--calc", "--convert-to", "xlsx", "--outdir", out, f], capture_output=True)
    r = load_workbook(os.path.join(out, os.path.basename(f)), data_only=True)
    c = r["Cover"]
    vals = {c.cell(i, 1).value: c.cell(i, 2).value for i in range(5, 25)}
    return (name, vals["Value per share (DCF, MXN)"], vals["Upside / (downside)"], vals["Terminal value as % of EV"],
            vals["Recommendation (pre-set rule: ±15%)"], vals["All integrity checks pass?"])

S = [
    ("Base case", {}),
    ("FX: peso +10% vs. spot (stronger MXN)", {"fx_shock": 0.10}),
    ("FX: peso −10% vs. spot (weaker MXN)", {"fx_shock": -0.10}),
    ("Desktop included (EBITDA in, full cost in net debt)", {"desk_on": 1}),
    ("Pension deficit deducted pre-tax", {"pension_tax": 0.0}),
    ("Postpaid/prepaid ARPU ratio k = 2.0 everywhere", {"k_ratio|*": 2.0}),
    ("Postpaid/prepaid ARPU ratio k = 4.0 everywhere", {"k_ratio|*": 4.0}),
    ("Postpaid net adds 2027E = 50% of LTM (bear engine)", {"post_2027": 0.50}),
    ("Postpaid net adds 2027E = 100% of LTM (bull engine)", {"post_2027": 1.00}),
    ("No further margin expansion", {"margin_drift": 0.0}),
    ("Capex 14.5% of revenue from 2028E (and terminal D&A)", {"capex_pct": 0.145, "da_term": 0.145}),
    ("Terminal growth 2% (−1% real)", {"g": 0.02}),
    ("Bear (operating): margin flat, net adds 50%, capex 14.5%, g 2.5%", {"margin_drift": 0.0, "post_2027": 0.5, "capex_pct": 0.145, "da_term": 0.145, "g": 0.025}),
    ("Bull: margin +0.3pp/yr, net adds 100%, g 3.5%", {"margin_drift": 0.003, "post_2027": 1.0, "g": 0.035}),
]
rows = [run(n, ch) for n, ch in S]
with open(os.path.join(ROOT, "docs", "scenarios.md"), "w") as f:
    f.write("# Scenario reruns (full model recalculated for each line)\n\n")
    f.write("| Scenario | Value/share (MXN) | vs. price | TV % of EV | Rule output | Checks pass |\n|---|---:|---:|---:|---|---|\n")
    for n, v, u, tv, rec, ok in rows:
        f.write(f"| {n} | {v:.2f} | {u:+.1%} | {tv:.0%} | {rec} | {ok} |\n")
for r_ in rows:
    print(f"{r_[0]:60s} {r_[1]:7.2f} {r_[2]:+.1%} {r_[3]:.0%} {r_[4]} {r_[5]}")
