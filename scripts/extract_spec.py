#!/usr/bin/env python3
"""Reverse a Trip Plan workbook into a spec JSON (same schema as build_trip_plan.py).
Use it to (a) reuse an old trip as the starting point for a new one,
(b) round-trip check a workbook. Usage: python3 extract_spec.py plan.xlsx spec.json"""
import json, sys
import openpyxl

def rgb(c):
    try:
        return (c.fill.fgColor.rgb or "")[-6:].upper() if c.fill and c.fill.fill_type else ""
    except Exception:
        return ""

def s(v):
    if v is None: return None
    if isinstance(v, float) and v.is_integer(): return int(v)
    return v

def is_band(ws, r, ncol):
    a = ws.cell(r, 1).value
    return isinstance(a, str) and all(ws.cell(r, c).value is None for c in range(2, ncol + 1))

def schedule(ws):
    out = {"schedule_title": ws["A1"].value, "days": []}
    day = None
    for r in range(3, ws.max_row + 1):
        a = ws.cell(r, 1).value
        if a is None and ws.cell(r, 4).value is None: continue
        if is_band(ws, r, 9) and str(a).upper().startswith("DAY"):
            day = {"banner": a, "highlights": None, "rows": []}; out["days"].append(day); continue
        if is_band(ws, r, 9) and str(a).startswith("★"):
            day["highlights"] = a; continue
        keys = ("time", "duration", "type", "activity", "area", "transport", "notes", "backup1", "backup2")
        row = {k: s(ws.cell(r, i).value) for i, k in enumerate(keys, 1) if ws.cell(r, i).value is not None}
        if rgb(ws.cell(r, 1)) == "FF0000": row["critical"] = True
        if day is not None: day["rows"].append(row)
    return out

def checklists(ws):
    out, cur = {"before": [], "apps": [], "items": []}, None
    order = ["before", "apps", "items"]; idx = -1
    for r in range(3, ws.max_row + 1):
        a = ws.cell(r, 1).value
        if is_band(ws, r, 6) and isinstance(a, str):
            idx += 1; cur = order[min(idx, 2)]; out[cur + "_label"] = a; continue
        if cur and ws.cell(r, 2).value:
            keys = ("category", "detail", "when", "link", "status")
            out[cur].append({k: s(ws.cell(r, i).value) for i, k in enumerate(keys, 2) if ws.cell(r, i).value is not None})
    return out

def sectioned(ws, keys, ncol):
    secs, sec = [], None
    for r in range(3, ws.max_row + 1):
        a = ws.cell(r, 1).value
        if is_band(ws, r, ncol) and isinstance(a, str):
            sec = {"title": a, "rows": []}; secs.append(sec); continue
        if sec is None or ws.cell(r, 3).value is None: continue
        row = {k: s(ws.cell(r, i).value) for i, k in enumerate(keys, 2) if ws.cell(r, i).value is not None}
        if rgb(ws.cell(r, 3)) == "FFF2CC": sec["backup"] = True
        if ws.cell(r, 3).font and ws.cell(r, 3).font.b: row["must_go"] = True
        sec["rows"].append(row)
    return secs

def budget(ws):
    tiers = {"B6D7A8": "prebook", "FCE4D6": "advance"}
    b = {"title": ws["A1"].value, "rows": []}
    h = str(ws["F2"].value or "")
    if "(" in h: b["currency"] = h.split("(")[-1].rstrip(")")
    for r in range(3, ws.max_row + 1):
        cat = ws.cell(r, 2).value
        if cat == "GRAND TOTAL": continue
        if cat == "TARGET": b["target"] = s(ws.cell(r, 7).value); continue
        if not cat or str(cat).startswith("VARIANCE"): continue
        b["rows"].append({"category": cat, "detail": ws.cell(r, 3).value, "unit": ws.cell(r, 4).value,
                          "qty": s(ws.cell(r, 5).value), "price": s(ws.cell(r, 6).value),
                          "tier": tiers.get(rgb(ws.cell(r, 2)), "onsite"), "booked": ws.cell(r, 8).value or "—"})
    return b

def cost(ws):
    members, c = [], 5
    while ws.cell(2, c).value and not str(ws.cell(2, c).value).startswith("Number"):
        members.append(ws.cell(2, c).value); c += 1
    rows = []
    for r in range(4, ws.max_row + 1):
        if ws.cell(r, 2).value:
            rows.append({"activity": ws.cell(r, 2).value, "pic": ws.cell(r, 3).value, "cost": s(ws.cell(r, 4).value),
                         "participants": [m for j, m in enumerate(members) if ws.cell(r, 5 + j).value == "x"]})
    prepay, fund = {}, {}
    for col in range(1, ws.max_column + 1):
        if ws.cell(2, col).value == "Member":
            for r in range(3, ws.max_row + 1):
                n = ws.cell(r, col).value
                if n in members and n not in prepay: prepay[n] = s(ws.cell(r, col + 2).value) or 0
                if ws.cell(r, col + 1).value == "Fund":
                    for k in range(len(members)):
                        fund[ws.cell(r + 1 + k, col).value] = s(ws.cell(r + 1 + k, col + 1).value) or 0
                    break
    return {"members": members, "rows": rows, "prepay": prepay, "fund": fund}

def main():
    wb = openpyxl.load_workbook(sys.argv[1])
    spec = schedule(wb["Schedule"])
    spec["checklists"] = checklists(wb["Checklists"])
    spec["local_name_label"] = wb["Food & Drink"]["D2"].value
    spec["food"] = sectioned(wb["Food & Drink"], ("day", "name", "local", "area", "price", "booked", "notes", "address"), 9)
    spec["attraction_title"] = wb["Attraction"]["A1"].value
    spec["attractions"] = sectioned(wb["Attraction"], ("city", "name", "local", "category", "access", "ticket", "time", "notes"), 9)
    spec["budget"] = budget(wb["Budget"])
    if "Cost Sharing" in wb.sheetnames: spec["cost_sharing"] = cost(wb["Cost Sharing"])
    json.dump(spec, open(sys.argv[2], "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    print("days", len(spec["days"]), "| sched rows", sum(len(d["rows"]) for d in spec["days"]),
          "| food", sum(len(x["rows"]) for x in spec["food"]), "| attractions", sum(len(x["rows"]) for x in spec["attractions"]),
          "| budget", len(spec["budget"]["rows"]))

if __name__ == "__main__":
    main()
