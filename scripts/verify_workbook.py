#!/usr/bin/env python3
"""Verify a Trip Plan workbook before delivery.
Usage: python3 verify_workbook.py new.xlsx [previous.xlsx]
1. Recalculates with LibreOffice (headless) and scans every cell for Excel errors.
2. Prints Budget GRAND TOTAL / VARIANCE and Cost Sharing totals.
3. If previous.xlsx is given: per-sheet diff count + first changed cells,
   so you can prove that out-of-scope sheets have 0 diffs.
Exit code 1 if any formula error is found."""
import os, subprocess, sys, tempfile
import openpyxl

ERRORS = ("#REF!", "#VALUE!", "#DIV/0!", "#NAME?", "#N/A", "#NUM!", "#NULL!")

def recalc(path):
    out = tempfile.mkdtemp()
    subprocess.run(["soffice", "--headless", "--convert-to", "xlsx", "--outdir", out, path],
                   stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, timeout=180)
    p = os.path.join(out, os.path.basename(path))
    if not os.path.exists(p): sys.exit("LibreOffice recalculation failed — is soffice installed?")
    return openpyxl.load_workbook(p, data_only=True)

def find_row(ws, col, label):
    for r in range(1, ws.max_row + 1):
        if ws.cell(r, col).value == label: return r

def main():
    new = sys.argv[1]; prev = sys.argv[2] if len(sys.argv) > 2 else None
    wb = recalc(new)
    errs = [(ws.title, c.coordinate, c.value) for ws in wb for row in ws.iter_rows() for c in row
            if isinstance(c.value, str) and c.value.strip() in ERRORS]
    print("Sheets:", wb.sheetnames)
    print("Formula errors:", errs or "none")
    if "Budget" in wb.sheetnames:
        b = wb["Budget"]
        for lab in ("GRAND TOTAL", "TARGET", "VARIANCE (Target − Total)"):
            r = find_row(b, 2, lab)
            if r: print(f"Budget {lab}: {b.cell(r, 7).value:,.0f}")
    if "Cost Sharing" in wb.sheetnames:
        cs = wb["Cost Sharing"]
        for col in range(1, cs.max_column + 1):
            if cs.cell(2, col).value == "Member":
                r = find_row(cs, col, "Total")
                if r: print("Cost Sharing totals (Share / Prepay / Balance):",
                            [cs.cell(r, col + k).value for k in (1, 2, 3)])
                break
    if prev:
        a = openpyxl.load_workbook(prev); b = openpyxl.load_workbook(new)
        print("\nDiff vs", os.path.basename(prev))
        for name in sorted(set(a.sheetnames) | set(b.sheetnames), key=lambda n: (b.sheetnames + a.sheetnames).index(n)):
            if name not in a.sheetnames: print(f"  {name}: NEW sheet"); continue
            if name not in b.sheetnames: print(f"  {name}: REMOVED"); continue
            wa, wb2 = a[name], b[name]
            diffs = [f"{openpyxl.utils.get_column_letter(c)}{r}"
                     for r in range(1, max(wa.max_row, wb2.max_row) + 1)
                     for c in range(1, max(wa.max_column, wb2.max_column) + 1)
                     if wa.cell(r, c).value != wb2.cell(r, c).value]
            print(f"  {name}: {len(diffs)} diffs" + (f" (first: {', '.join(diffs[:8])})" if diffs else ""))
    sys.exit(1 if errs else 0)

if __name__ == "__main__":
    main()
