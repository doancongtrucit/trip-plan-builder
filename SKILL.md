---
name: "trip-plan-builder"
description: Build or update a multi-day trip plan Excel workbook (Schedule, Checklists, Food & Drink, Attraction, Budget, Cost Sharing) in a standard group-trip format. Use for "lên trip plan", "build lịch trình", "trip plan Excel", or a new version of an existing plan.
---

# Trip Plan Builder

Builds a 6-sheet trip-plan workbook for group trips. Reference build: a 6-day Busan → Seoul group trip (see `assets/`).

Chat with the user in Vietnamese. Workbook: English headers/Type values, Vietnamese notes.

## Folder map

| Path | Use |
|---|---|
| `scripts/build_trip_plan.py` | spec JSON → formatted workbook (NEW mode) |
| `scripts/extract_spec.py` | existing workbook → spec JSON (reuse an old trip as the base for a new one; round-trip check) |
| `scripts/verify_workbook.py` | LibreOffice recalc + formula-error scan + totals + per-sheet diff vs previous version |
| `references/format_rules.md` | Format standard per sheet — **read before writing any cell** |
| `references/research_checklist.md` | Mandatory fact checks per scheduled row (closures, fixed departures, airport buffer…) |
| `references/spec_schema.md` | Spec JSON schema |
| `assets/sample_spec_korea_6n5d.json` | Full spec of the reference Korea plan (personal data anonymized) — copy its tone, density and row patterns |

`SKILL_DIR` below = this skill's folder.

## Modes

| Mode | When | Approach |
|---|---|---|
| NEW | New trip | Intake → research → write spec → `build_trip_plan.py` |
| NEW from old trip | "làm giống chuyến Korea cho chuyến X" | `extract_spec.py` on the old file → keep structure, replace all trip facts → build |
| UPDATE | Next version of an existing plan | Edit the latest file in place with openpyxl (copy style from a neighbouring row of the same Type); never rebuild — rebuilding drops the user's manual edits |

## Step 1 — Intake (never guess)

Ask for anything missing (AskUserQuestion or a table: field | why needed):
destinations + city order · dates · flights both ways (code, times, airports) · hotels (name + address) · group size + member names · budget target/person + currency · must-have attractions · items already booked/paid.

Unknown values → `[TBD: ...]`. Never invent prices, restaurants, times or names.

UPDATE mode: first check whether the user edited the latest `_vN` locally (stage it, run `verify_workbook.py <their file> <last Claude version>`). If anything changed, list the changes and confirm before building the next version.

## Step 2 — Research

Follow `references/research_checklist.md` for every scheduled row. A conflict (e.g. attraction closed on that weekday) is flagged in Notes and raised to the user with 2–3 options and a recommendation — never silently moved.

## Step 3 — Build

Read `references/format_rules.md`, then:

```bash
python3 "$SKILL_DIR/scripts/build_trip_plan.py" spec.json <Base_Name>_v1.xlsx
# reuse an old trip:
python3 "$SKILL_DIR/scripts/extract_spec.py" old_plan.xlsx spec.json
```

## Step 4 — Verify (mandatory)

```bash
python3 "$SKILL_DIR/scripts/verify_workbook.py" <new>.xlsx [<previous>.xlsx]
```

- Exit 1 / any formula error → fix before delivery.
- UPDATE: out-of-scope sheets must show **0 diffs**; report the counts to the user.
- Re-run research checks 1–5 on every changed row and on both days of any swap.
- Scan for stale references after structural changes (removed place still in Food/Attraction/Checklists/Budget, old area names, unrenumbered No. columns).

## Step 5 — Deliver

- Send the file; if the user's trip folder is connected, commit it there as the new `_vN`.
- Reply in Vietnamese: one line on what changed, then a table of open conflicts/risks (item | why | decision needed).
- Log the version's changes and unresolved items to project notes.
