# Format rules — trip plan standard

Source: reference Korea plan (v1 → v17 feedback). Apply to every new or updated plan.

## Workbook
- Sheet order fixed: Schedule, Checklists, Food & Drink, Attraction, Budget, Cost Sharing.
- File name `<Base_Name>_v<N>.xlsx`; N+1 on every save, never overwrite a previous version.
- No footer note blocks (data rows only). No emoji.
- Language: headers and Type values in English; Notes/descriptions in Vietnamese; local-language names in the Local Name column.

## Schedule
Columns: Time | Duration | Type | Activity / Location | Area | Transport | Notes | Backup Plan 1 | Backup Plan 2

| Element | Rule |
|---|---|
| Day banner | Red row, merged: `DAY n — T7 24/10 (THỨ BẢY) — <route summary>` |
| Highlights | Yellow row right below the banner: `★ HIGHLIGHTS: <must-haves of the day>` |
| Type | Only `TRANSPORT`, `MEAL`, `SIGHTSEEING`, `REST`, `MEETING` |
| Time | `9h30–11h`, rounded to :00/:30. Flight/train-driven days (Day 1, last day) keep exact times |
| Duration | `45p` under 1h; `1h`, `1h30`, `2h26` |
| Gather row | Day 2+ start with `MEETING` "Tập trung xuất phát" (default 7h00, first activity 7h30) — confirm with the user |
| Activity | Bare place/dish name. Meals: `BREAKFAST: <restaurant>`, `LUNCH: ...`, `DINNER: ...` |
| Notes | Terse, actionable warnings only (closures, booking, crowd timing). No trivia/narrative |
| Backup Plan 1/2 | Used on MEAL rows: `<restaurant> - <address> - ~<distance> / <walk time>` |
| Fills | MEAL activity = pink F4CCCC + bold; SIGHTSEEING activity = beige F2EFEA + bold; flight/train-critical time cell = red FF0000. No other fills |
| Conflicts | Write `⚠ CONFLICT: ...` / `CHƯA XỬ LÝ` in Notes and raise it to the user |

## Checklists
- Exactly 3 sections: A. Before the trip (sorted by deadline) · B. Apps to install · C. Essential items to bring.
- Include links. No clothing / miscellaneous items. One-line notes.
- Status `Done` = green; `When` overdue = red bold.

## Food & Drink
- Columns: No. | Day | Food & Drink | Local Name | Area | Price Range | Booked? (x) | Notes | Address.
- Sections: one per day → "MÓN NÊN THỬ" → "TIỆN LỢI" → optional yellow "BỔ SUNG" (`"backup": true`).
- Only places on the route. Real address for every scheduled meal.

## Attraction
- Columns: No. | City | Attraction | Local Name | Category | Area / Access | Ticket | Time Needed | Notes.
- Sections per city/day → rain backup (max 2) → yellow "BACKUP — ĐIỂM HOT THAY THẾ".
- Must-haves bold (`"must_go": true`). Only places in route cities. Renumber No. after removals.
- If a tier column is requested: labels exactly `Must Go` / `Have To Go` / `Should Go`.

## Budget (per person)
- Unit × Qty × Price = Total (formula). GRAND TOTAL, TARGET, VARIANCE (Target − Total).
- Tier fill: `prebook` green B6D7A8 (paid/booked) · `advance` orange FCE4D6 (book before trip, has deadline) · `onsite` no fill.
- Costs live only here — no cost column in Schedule.

## Cost Sharing
- Members as columns, "x" dropdown per activity, PIC dropdown, auto Number of Participants and Cost per Person.
- Settlement block: Share to Pay (SUMIF) / Prepay (yellow input) / Balance / settlement text. FUND block below.
- Yellow = manual input only.
