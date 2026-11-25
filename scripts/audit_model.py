"""Scan model-logic sheets for numeric literals inside formulas (other than 0 and 1).
Writes docs/model_audit.txt. Usage: python scripts/audit_model.py model/AMX_Valuation_Model.xlsx"""
import re, sys, os
from openpyxl import load_workbook
path = sys.argv[1]
wb = load_workbook(path)  # formulas, not values
SHEETS = ["Drivers", "Projections", "WACC", "DCF", "Sensitivity", "Comps", "Cover"]
num = re.compile(r'(?<![A-Za-z_$!:\d.])(\d+\.?\d*(?:[eE][-+]?\d+)?)(?![\d.]*[A-Za-z_(!$:])')
total, flagged = 0, []
for sh in SHEETS:
    for row in wb[sh].iter_rows():
        for c in row:
            v = c.value
            if isinstance(v, str) and v.startswith("="):
                total += 1
                body = re.sub(r'"[^"]*"', '""', v)                    # ignore text literals
                body = re.sub(r"'[^']*'!", "", body)                    # ignore quoted sheet names
                body = re.sub(r"\$?[A-Z]{1,3}\$?\d+", "", body)          # ignore cell references
                lits = [x for x in num.findall(body) if float(x) not in (0.0, 1.0)]
                if lits:
                    flagged.append(f"{sh}!{c.coordinate}: {lits}  ::  {v[:120]}")
os.makedirs("docs", exist_ok=True)
with open("docs/model_audit.txt", "w") as f:
    f.write(f"Hard-code audit of {os.path.basename(path)}\nSheets scanned: {', '.join(SHEETS)}\n")
    f.write(f"Formulas scanned: {total}\nFormulas containing numeric literals other than 0/1: {len(flagged)}\n\n")
    f.write("\n".join(flagged) + ("\n" if flagged else "None.\n"))
    f.write("\nNot scanned: Inputs (assumptions live there by design), Historicals (reported data), Checks (tolerances such as 0.5 are test thresholds, not model logic).\n")
print(total, len(flagged)); print("\n".join(flagged[:40]))
