# Spec JSON schema (input of build_trip_plan.py / output of extract_spec.py)

Full real example: `assets/sample_spec_korea_6n5d.json`. Omit empty keys.

```json
{
  "schedule_title": "SCHEDULE — BUSAN → SEOUL | 6 DAYS / 5 NIGHTS | 24–29 OCT 2026",
  "local_name_label": "Korean Name",
  "attraction_title": "ATTRACTION — SEOUL · BUSAN",
  "days": [{"banner": "DAY 1 — T7 24/10 (THỨ BẢY) — ĐẾN BUSAN", "highlights": "★ HIGHLIGHTS: ...",
            "rows": [{"time": "9h45–10h30", "duration": "45p", "type": "MEAL", "activity": "BREAKFAST: ...",
                      "area": "", "transport": "", "notes": "", "backup1": "", "backup2": "", "critical": false}]}],
  "checklists": {"before": [{"category": "", "detail": "", "when": "", "link": "", "status": "Done"}],
                 "apps": [], "items": [],
                 "before_label": "optional custom section title"},
  "food": [{"title": "DAY 1 — BUSAN: NAMPO", "backup": false,
            "rows": [{"day": "D1", "name": "", "local": "", "area": "", "price": "", "booked": "", "notes": "", "address": ""}]}],
  "attractions": [{"title": "BUSAN (Day 1-3)", "backup": false,
                   "rows": [{"city": "", "name": "", "local": "", "category": "", "access": "", "ticket": "",
                             "time": "", "notes": "", "must_go": true}]}],
  "budget": {"title": "BUDGET — ... — 1 NGƯỜI", "currency": "VND", "target": 20000000,
             "rows": [{"category": "Flight", "detail": "", "unit": "Vé", "qty": 1, "price": 4400000,
                       "tier": "prebook|advance|onsite", "booked": "x"}]},
  "cost_sharing": {"members": ["A", "B"], "currency": "VND", "rows_count": 30, "hint": "optional",
                   "rows": [{"activity": "", "pic": "A", "cost": 0, "participants": ["A", "B"]}],
                   "prepay": {"A": 0}, "fund": {"A": 0}}
}
```

Notes
- `critical: true` → red Time cell (flight, train, fixed departure).
- `type` drives fills: MEAL pink, SIGHTSEEING beige.
- Section `backup: true` → whole section yellow (Food "BỔ SUNG", Attraction "BACKUP — ĐIỂM HOT").
- `cost_sharing` optional — omit for trips without group expense splitting.
