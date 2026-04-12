# Config Restructure Design

**Date:** 2026-04-11

## Overview

Five changes to the Budget Planner, primarily restructuring the Config sheet.

| # | Change | Scope |
|---|--------|-------|
| 1 | Per-WP rate in WP Budget table, default 6500. Replaces `fte_rate` in all ceiling/EUR formulas. | Config + Plan + Summary + Trends + Actuals |
| 2 | `git pull` with error handling at start of all .bat files | 4 .bat files |
| 3 | Multiple month ranges per WP via WP Schedule table (WP ID + Start + End + Type) | Config + Capacity + Plan + Actuals |
| 4 | Personnel Ceiling next to Rate, shorter headers | Config table layout |
| 5 | RATE BY GRADE moved to bottom of Config sheet | Config sheet builder |

## Change 1: Per-WP Rate

Each WP gets its own EUR/PM rate as a blue input cell in the Budget table. Default value: 6500.

- Personnel Ceiling formula changes from `=FTE × fte_rate` to `=FTE × WP_Rate` (same row)
- `fte_rate` named range remains (weighted average from grade rates) but is no longer referenced in any formula — purely informational
- All sheets that previously used `fte_rate` for PM→EUR conversion now reference the per-WP rate via direct cell reference: `=PM_value * Config!$D${wp_row}`

## Change 2: git pull in .bat Files

All four project .bat files (`setup.bat`, `init.bat`, `demo.bat`, `rebuild.bat`) get a `git pull` with error handling immediately after `@echo off`.

**setup.bat** — git pull only (no venv check, it creates the venv):

```bat
@echo off
git pull
if errorlevel 1 (
    echo WARNING: git pull failed. Continuing with local version.
    pause
)
...rest of file (python check, venv creation, etc.)...
```

**init.bat, demo.bat, rebuild.bat** — git pull then venv check:

```bat
@echo off
git pull
if errorlevel 1 (
    echo WARNING: git pull failed. Continuing with local version.
    pause
)
if not exist .venv\Scripts\activate.bat (
    echo ERROR: Run setup.bat first.
    pause
    exit /b 1
)
call .venv\Scripts\activate.bat
...rest of file...
```

Shows warning and pauses on pull failure, then continues. Aborts if venv is missing.

## Change 3: WP Schedule Table

Split the current single WP table into two separate tables on the Config sheet.

### WP Budget Table (one row per WP)

| Col | Header | Style | Content |
|-----|--------|-------|---------|
| A | WP ID | Blue input | "WP1", "WP2", ... |
| B | WP Name | Blue input | Free text |
| C | FTE Budget | Blue input | Person-months, format `0.00` |
| D | Rate (€/PM) | Blue input | EUR per PM, default 6500, format `€#,##0` |
| E | Ceiling (€) | Black formula | `=C{row}*D{row}` |

TOTAL row sums columns C and E.

### WP Schedule Table (multiple rows per WP allowed)

| Col | Header | Style | Content |
|-----|--------|-------|---------|
| A | WP ID | Blue input | "WP1", "WP4", ... (can repeat) |
| B | Start | Blue input | Month number, format `0` |
| C | End | Blue input | Month number, format `0` |
| D | Type | Blue input | Distribution curve: flat/front-loaded/bell/back-loaded |

No TOTAL row. Data validation dropdown on Type column. Active Months column is dropped.

### How Builders Determine Active Months

For each WP, collect all (start, end) pairs from the Schedule table. A month is active if it falls within any of those ranges. Builders filter the schedule list by WP ID.

## Change 4: Shorter Headers

Personnel Ceiling placed next to Rate in the Budget table. Headers shortened:
- "Rate (€/PM)" instead of "FTE Rate (EUR)"
- "Ceiling (€)" instead of "Personnel Ceiling (€)"

## Change 5: Config Sheet Section Order

New top-to-bottom layout:

1. **Project Parameters** (rows 1-5)
2. *(gap row)*
3. **WP BUDGET** section header + table
4. *(gap row)*
5. **WP SCHEDULE** section header + table
6. *(gap row)*
7. **TEAM ROSTER** section header + table (unchanged)
8. *(gap row)*
9. **REPORTING PERIODS** section header + table (unchanged)
10. *(gap row)*
11. **RATE BY GRADE** section header + table (moved from position 2 to bottom)

`fte_rate` formula in Project Parameters still references the grade table — just at higher row numbers. Named range `tbl_GradeRates` updates accordingly.

## Formula Impact Across Sheets

| Sheet | Current | New |
|-------|---------|-----|
| **Config** | Ceiling: `=C{row}*fte_rate` | `=C{row}*D{row}` |
| **Plan** | `partner_ceiling = SUM(Config!H)` | `partner_ceiling = SUM(Config!E)` |
| **Plan** | Header displays `fte_rate` ref | Removed |
| **Summary** | `Plan_FTE × fte_rate` per WP/year | `Plan_FTE × Config!$D${wp_row}` |
| **Trends** | `SUM(Plan_months) × fte_rate` | `SUMPRODUCT(Plan_months, WP_rates)` or per-WP terms |
| **Actuals** | `=partner_ceiling` | No change (still points to named range) |

## Reader & Rebuild Impact

`read_config_sheet()` changes to parse two tables:

1. **Find WP BUDGET section** — scan for header, read rows (ID, Name, FTE, Rate, Ceiling)
2. **Find WP SCHEDULE section** — scan for header, read rows (ID, Start, End, Type)

Returned data structure changes:

```python
# Current
{"id": "WP4", "name": "...", "fte": 15.0, "start": 2, "end": 11, "type": "bell"}

# New
budget: {"id": "WP4", "name": "...", "fte": 15.0, "rate": 10000}
schedule: [
    {"id": "WP4", "start": 2, "end": 11, "type": "bell"},
    {"id": "WP4", "start": 12, "end": 14, "type": "flat"},
]
```

Builders receive both lists. Each builder that needs active months filters the schedule list by WP ID.

## Seed Data Default

`config.py` changes:
- `DEFAULT_RATE` changes from 7500 to 6500
- `WORK_PACKAGES` list splits into `WORK_PACKAGES` (budget fields) and `WP_SCHEDULE` (schedule rows)
- See `docs/wp-schedule-type-suggestions.md` for per-WP distribution type recommendations
