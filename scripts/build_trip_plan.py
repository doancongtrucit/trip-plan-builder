#!/usr/bin/env python3
"""Build a Trip Plan workbook (6 sheets) from a JSON spec.
Usage: python3 build_trip_plan.py spec.json out.xlsx
Sheets: Schedule, Checklists, Food & Drink, Attraction, Budget, Cost Sharing.
Formulas (Budget, Cost Sharing) stay live so the user can edit inputs later."""
import json, sys
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter as L
from openpyxl.worksheet.datavalidation import DataValidation

NAVY, BLUE, GREEN, RED = "1F3864", "2E5C9A", "00B050", "C00000"
HILITE, HILITE_TXT, SECTION = "FFE699", "7F4F00", "D9D9D9"
MEAL_FILL, SIGHT_FILL, CRIT_FILL = "F4CCCC", "F2EFEA", "FF0000"
BACKUP_FILL, INPUT_FILL = "FFF2CC", "FFFF00"
TIER = {"prebook": "B6D7A8", "advance": "FCE4D6", "onsite": None}
STATUS_DONE = "E2EFDA"

def fill(c): return PatternFill("solid", fgColor=c) if c else PatternFill()
def border(c="D8D8D8"):
    s = Side(style="thin", color=c); return Border(left=s, right=s, top=s, bottom=s)

def title_row(ws, text, ncol, size=15, row=1, col=1):
    ws.merge_cells(start_row=row, start_column=col, end_row=row, end_column=col + ncol - 1)
    c = ws.cell(row, col, text)
    c.font = Font(name="Arial", size=size, bold=True, color="FFFFFF")
    c.fill = fill(NAVY); c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[row].height = 30

def header(ws, row, labels, colors=None, col=1, bcol="D8D8D8"):
    for i, h in enumerate(labels):
        c = ws.cell(row, col + i, h)
        c.font = Font(name="Arial", size=10, bold=True, color="FFFFFF")
        c.fill = fill((colors or {}).get(h, BLUE))
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = border(bcol)
    ws.row_dimensions[row].height = 32

def band(ws, row, text, ncol, color, font_color="000000", font="Arial", size=10):
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=ncol)
    c = ws.cell(row, 1, text)
    c.font = Font(name=font, size=size, bold=True, color=font_color)
    c.fill = fill(color); c.alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[row].height = 22

def widths(ws, ws_widths):
    for k, v in ws_widths.items(): ws.column_dimensions[k].width = v

def body(c, font="Arial", center=False, wrap=True, bcol="D8D8D8", bold=False):
    c.font = Font(name=font, size=10 if font == "Arial" else 11, bold=bold)
    c.alignment = Alignment(horizontal="center" if center else None, vertical="center", wrap_text=wrap)
    c.border = border(bcol)

# ---------------- Schedule ----------------
def sheet_schedule(wb, spec):
    ws = wb.active; ws.title = "Schedule"
    ws.sheet_properties.tabColor = "F1C232"; ws.sheet_view.showGridLines = False
    cols = ["Time", "Duration", "Type", "Activity / Location", "Area", "Transport", "Notes", "Backup Plan 1", "Backup Plan 2"]
    title_row(ws, spec["schedule_title"], len(cols))
    header(ws, 2, cols, {"Backup Plan 1": GREEN, "Backup Plan 2": GREEN}, bcol="CCCCCC")
    r = 3
    for d in spec["days"]:
        band(ws, r, d["banner"], len(cols), RED, "FFFFFF", "Calibri", 11); r += 1
        if d.get("highlights"):
            band(ws, r, d["highlights"], len(cols), HILITE, HILITE_TXT, "Calibri", 11); r += 1
        for it in d["rows"]:
            vals = [it.get(k, "") for k in ("time", "duration", "type", "activity", "area", "transport", "notes", "backup1", "backup2")]
            for i, v in enumerate(vals, 1):
                c = ws.cell(r, i, v or None)
                body(c, "Calibri", center=i in (1, 2, 3), bcol="CCCCCC")
            t = (it.get("type") or "").upper()
            act = ws.cell(r, 4)
            if t == "MEAL": act.fill = fill(MEAL_FILL); act.font = Font(name="Calibri", size=11, bold=True)
            elif t == "SIGHTSEEING": act.fill = fill(SIGHT_FILL); act.font = Font(name="Calibri", size=11, bold=True)
            if it.get("critical"):
                ws.cell(r, 1).fill = fill(CRIT_FILL); ws.cell(r, 1).font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
            r += 1
    widths(ws, {"A": 13, "B": 10, "C": 16, "D": 50, "E": 16, "F": 34, "G": 80, "H": 50, "I": 50})
    ws.freeze_panes = "D3"

# ---------------- Checklists ----------------
def sheet_checklists(wb, spec):
    ws = wb.create_sheet("Checklists"); ws.sheet_view.showGridLines = False
    cols = ["No.", "Category", "What / Why", "When", "Link", "Status"]
    title_row(ws, "CHECKLISTS", len(cols), size=11)
    header(ws, 2, cols, bcol="BFBFBF")
    r = 3
    ck = spec["checklists"]
    for key, label in (("before", "A. BEFORE THE TRIP (sorted by deadline)"), ("apps", "B. APPS TO INSTALL"), ("items", "C. ESSENTIAL ITEMS TO BRING")):
        band(ws, r, ck.get(key + "_label", label), len(cols), SECTION); r += 1
        for n, it in enumerate(ck.get(key, []), 1):
            vals = [n, it.get("category"), it.get("detail"), it.get("when"), it.get("link"), it.get("status")]
            for i, v in enumerate(vals, 1):
                c = ws.cell(r, i, v); body(c, center=i in (1, 4, 6), bcol="BFBFBF")
            if it.get("link", "").startswith("http"): ws.cell(r, 5).hyperlink = it["link"]; ws.cell(r, 5).font = Font(name="Arial", size=10, color="0563C1", underline="single")
            if (it.get("status") or "").lower() == "done": ws.cell(r, 6).fill = fill(STATUS_DONE)
            if (it.get("when") or "").upper() in ("OVERDUE", "QUÁ HẠN"): ws.cell(r, 4).font = Font(name="Arial", size=10, bold=True, color=RED)
            r += 1
    widths(ws, {"A": 5, "B": 42, "C": 59, "D": 15, "E": 45, "F": 9})
    ws.freeze_panes = "C3"

# ---------------- Food & Drink ----------------
def sheet_food(wb, spec):
    ws = wb.create_sheet("Food & Drink"); ws.sheet_view.showGridLines = False
    ln = spec.get("local_name_label", "Local Name")
    cols = ["No.", "Day", "Food & Drink", ln, "Area", "Price Range", "Booked? (x)", "Notes", "Address"]
    title_row(ws, "FOOD & DRINK", len(cols), size=11)
    header(ws, 2, cols)
    r, n = 3, 1
    for sec in spec["food"]:
        band(ws, r, sec["title"], len(cols), SECTION); r += 1
        for it in sec["rows"]:
            vals = [n, it.get("day", "—"), it.get("name"), it.get("local"), it.get("area"), it.get("price"), it.get("booked", "—"), it.get("notes"), it.get("address")]
            for i, v in enumerate(vals, 1):
                c = ws.cell(r, i, v); body(c, center=i in (1, 2, 7))
                if sec.get("backup"): c.fill = fill(BACKUP_FILL)
            r += 1; n += 1
    widths(ws, {"A": 5, "B": 6, "C": 34, "D": 20, "E": 22, "F": 16, "G": 12, "H": 50, "I": 45})
    ws.freeze_panes = "E3"

# ---------------- Attraction ----------------
def sheet_attraction(wb, spec):
    ws = wb.create_sheet("Attraction"); ws.sheet_view.showGridLines = False
    ln = spec.get("local_name_label", "Local Name")
    cols = ["No.", "City", "Attraction", ln, "Category", "Area / Access", "Ticket", "Time Needed", "Notes"]
    title_row(ws, spec.get("attraction_title", "ATTRACTION"), len(cols), size=11)
    header(ws, 2, cols)
    r, n = 3, 1
    for sec in spec["attractions"]:
        band(ws, r, sec["title"], len(cols), SECTION); r += 1
        for it in sec["rows"]:
            vals = [n, it.get("city"), it.get("name"), it.get("local"), it.get("category"), it.get("access"), it.get("ticket"), it.get("time"), it.get("notes")]
            for i, v in enumerate(vals, 1):
                c = ws.cell(r, i, v); body(c, center=i in (1, 2, 8))
                if sec.get("backup"): c.fill = fill(BACKUP_FILL)
            if it.get("must_go"): ws.cell(r, 3).font = Font(name="Arial", size=10, bold=True)
            r += 1; n += 1
    widths(ws, {"A": 5, "B": 10, "C": 34, "D": 20, "E": 20, "F": 28, "G": 22, "H": 11, "I": 55})
    ws.freeze_panes = "E3"

# ---------------- Budget ----------------
def sheet_budget(wb, spec):
    ws = wb.create_sheet("Budget"); ws.sheet_view.showGridLines = False
    b = spec["budget"]; cur = b.get("currency", "VND")
    cols = ["No.", "Category", "Detail", "Unit", "Qty", f"Price / Unit ({cur})", f"Total ({cur})", "Booked? (x)"]
    title_row(ws, b["title"], len(cols))
    header(ws, 2, cols)
    r = 3
    for n, it in enumerate(b["rows"], 1):
        vals = [n, it["category"], it["detail"], it.get("unit"), it.get("qty", 1), it.get("price"), f"=E{r}*F{r}", it.get("booked", "—")]
        for i, v in enumerate(vals, 1):
            c = ws.cell(r, i, v); body(c, center=i in (1, 4, 5, 8))
            c.fill = fill(TIER.get(it.get("tier", "onsite")))
            if i in (6, 7): c.number_format = "#,##0"
        r += 1
    last = r - 1
    tot = r
    ws.cell(tot, 2, "GRAND TOTAL"); ws.cell(tot, 7, f"=SUM(G3:G{last})")
    ws.cell(tot, 8, f'=COUNTIF(H3:H{last},"x")&" / "&COUNTIF(H3:H{last},"<>—")')
    for i in range(1, 9):
        c = ws.cell(tot, i); c.fill = fill(NAVY); c.font = Font(name="Arial", size=10, bold=True, color="FFFFFF"); c.border = border()
        c.alignment = Alignment(horizontal="center", vertical="center")
    ws.cell(tot, 7).number_format = "#,##0"
    if b.get("target"):
        ws.cell(tot + 1, 2, "TARGET"); ws.cell(tot + 1, 7, b["target"])
        ws.cell(tot + 2, 2, "VARIANCE (Target − Total)"); ws.cell(tot + 2, 7, f"=G{tot+1}-G{tot}")
        for rr in (tot + 1, tot + 2):
            for i in (2, 7):
                c = ws.cell(rr, i); body(c, bold=True); c.number_format = "#,##0"
    widths(ws, {"A": 5, "B": 18, "C": 50, "D": 8, "E": 6, "F": 16, "G": 16, "H": 12})
    ws.freeze_panes = "C3"

# ---------------- Cost Sharing ----------------
def sheet_cost(wb, spec):
    cs = spec.get("cost_sharing")
    if not cs: return
    ws = wb.create_sheet("Cost Sharing"); ws.sheet_view.showGridLines = False
    m = cs["members"]; nm = len(m); nrows = cs.get("rows_count", 30)
    c0 = 5; cE = c0 + nm - 1                   # member columns
    cN, cC = cE + 1, cE + 2                   # participants, cost/person
    sP = cC + 2                               # settlement block start (one spacer col)
    left = ["#", "Activity", "PIC", "Cost"] + m + ["Number of\nParticipants", "Cost per\nPerson"]
    title_row(ws, "COSTS & EXPENSE SHARING", len(left))
    title_row(ws, "SETTLEMENT", 5, col=sP)
    header(ws, 2, left)
    header(ws, 2, ["Member", "Share to Pay", f"Prepay ({cs.get('currency','VND')})", "Balance", "Settlement"], col=sP)
    ws.merge_cells(start_row=3, start_column=1, end_row=3, end_column=len(left))
    note = ws.cell(3, 1, cs.get("hint", "Enter activities below. Yellow = manual input. Mark 'x' under each member who joined that activity."))
    note.font = Font(name="Arial", size=9, italic=True, color="404040")
    first, last = 4, 4 + nrows - 1
    rows = cs.get("rows", [])
    for k in range(nrows):
        r = first + k; it = rows[k] if k < len(rows) else {}
        ws.cell(r, 1, k + 1); ws.cell(r, 2, it.get("activity")); ws.cell(r, 3, it.get("pic")); ws.cell(r, 4, it.get("cost"))
        for j, name in enumerate(m):
            ws.cell(r, c0 + j, "x" if name in it.get("participants", []) else None)
        ws.cell(r, cN, f'=COUNTIF({L(c0)}{r}:{L(cE)}{r},"x")')
        ws.cell(r, cC, f"=IF({L(cN)}{r}=0,0,ROUND(D{r}/{L(cN)}{r},0))")
        for i in range(1, cC + 1):
            c = ws.cell(r, i); body(c, center=i != 2)
            if i in (4, cC): c.number_format = "#,##0"
    dv = DataValidation(type="list", formula1='"x"', allow_blank=True); ws.add_data_validation(dv)
    dv.add(f"{L(c0)}{first}:{L(cE)}{last}")
    dv2 = DataValidation(type="list", formula1='"' + ",".join(m) + '"', allow_blank=True); ws.add_data_validation(dv2)
    dv2.add(f"C{first}:C{last}")
    pre = cs.get("prepay", {})
    for j, name in enumerate(m):
        r = 3 + j; mc = L(c0 + j)
        P, Q, R, S, T = sP, sP + 1, sP + 2, sP + 3, sP + 4
        ws.cell(r, P, name)
        ws.cell(r, Q, f'=SUMIF({mc}{first}:{mc}{last},"x",{L(cC)}{first}:{L(cC)}{last})')
        ws.cell(r, R, pre.get(name, 0))
        ws.cell(r, S, f"={L(R)}{r}-{L(Q)}{r}")
        ws.cell(r, T, f'=IF({L(S)}{r}=0,"Settled",IF({L(S)}{r}<0,"Pay "&TEXT(-{L(S)}{r},"#,##0")&" more","Refund "&TEXT({L(S)}{r},"#,##0")))')
        for i in range(P, T + 1):
            c = ws.cell(r, i); body(c, center=i != T)
            if i in (Q, R, S): c.number_format = "#,##0"
        ws.cell(r, R).fill = fill(INPUT_FILL)
    tr = 3 + nm
    ws.cell(tr, sP, "Total")
    for off in (1, 2, 3):
        ws.cell(tr, sP + off, f"=SUM({L(sP+off)}3:{L(sP+off)}{tr-1})").number_format = "#,##0"
    for i in range(sP, sP + 5):
        body(ws.cell(tr, i), center=True, bold=True)
    # FUND block
    fr = tr + 3
    title_row(ws, "FUND", 2, size=11, row=fr, col=sP)
    header(ws, fr + 1, ["Member", "Fund"], col=sP)
    for j, name in enumerate(m):
        r = fr + 2 + j
        ws.cell(r, sP, name); ws.cell(r, sP + 1, cs.get("fund", {}).get(name, 0))
        body(ws.cell(r, sP), center=True); body(ws.cell(r, sP + 1), center=True)
        ws.cell(r, sP + 1).fill = fill(INPUT_FILL); ws.cell(r, sP + 1).number_format = "#,##0"
    r = fr + 2 + nm
    ws.cell(r, sP, "Total"); ws.cell(r, sP + 1, f"=SUM({L(sP+1)}{fr+2}:{L(sP+1)}{r-1})").number_format = "#,##0"
    body(ws.cell(r, sP), center=True, bold=True); body(ws.cell(r, sP + 1), center=True, bold=True)
    ws.column_dimensions["A"].width = 4; ws.column_dimensions["B"].width = 30
    ws.column_dimensions["C"].width = 10; ws.column_dimensions["D"].width = 14
    for j in range(nm): ws.column_dimensions[L(c0 + j)].width = 8
    ws.column_dimensions[L(cN)].width = 12; ws.column_dimensions[L(cC)].width = 12
    ws.column_dimensions[L(cC + 1)].width = 3
    for k, w in zip(range(sP, sP + 5), (12, 14, 14, 14, 32)): ws.column_dimensions[L(k)].width = w
    ws.freeze_panes = "B4"

def main():
    spec = json.load(open(sys.argv[1], encoding="utf-8"))
    wb = Workbook()
    sheet_schedule(wb, spec); sheet_checklists(wb, spec); sheet_food(wb, spec)
    sheet_attraction(wb, spec); sheet_budget(wb, spec); sheet_cost(wb, spec)
    wb.save(sys.argv[2]); print("saved", sys.argv[2])

if __name__ == "__main__":
    main()
