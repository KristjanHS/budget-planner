# Improvements Round 2 — Design Document

**Date:** 2026-04-03
**Status:** Design complete

## Decisions Summary

| # | Item | Decision |
|---|------|----------|
| 1 | Password protection | Comment out temporarily |
| 2 | WP config in Plan | Remove from Plan, add derived columns to Config |
| 3 | Plan header rows | Flatten 4 rows into 1 row |
| 4 | Project start | 01.11.2026 |
| 5 | Project duration | 14 months (1y2m) |
| 6 | WP share seeds | Remove entirely |
| 7 | Team tab | Merge into Config (stacked), eliminate Team sheet |
| 8 | Column grouping | Calendar year boundaries |
| 9 | Year framing | Calendar years everywhere (all sheets, Summary) |
| 10 | Trends E14 TOTAL | Fix #VALUE! error |
| 11 | --skeleton flag | Remove entirely |
| 12 | Demo mode | `init --demo` flag, time-based capacity assignments |

---

## Detailed Design

### 1. Disable password protection (temporary)

Comment out the call to `_apply_protection()` in workbook.py. Leave the function intact for later re-enablement. Skip protection-related tests.

### 2. Config sheet — single source of WP configuration

**Config sheet layout (stacked):**
```
Row 1:  "Project Title"     | <value>           (blue input)
Row 2:  "Project Start"     | <date value>      (blue input)
Row 3:  "Project Months"    | <integer>         (blue input)
Row 4:  "FTE Rate (EUR)"    | <value>           (blue input)
Row 5:  (blank separator)
Row 6:  WORK PACKAGES section header
Row 7:  Headers: WP ID | WP Name | Total FTE | Start Mo | End Mo | Active Months | Personnel Ceiling (EUR)
Row 8+: WP data rows (blue input for cols A-E, formulas for F-G)
        Col F: =E-D+1 (Active Months, black formula)
        Col G: =C*fte_rate (Personnel Ceiling, black formula)
Row N:  TOTAL row (SUM on cols C, G)
Row N+1: (blank separator)
Row N+2: TEAM ROSTER section header
Row N+3: Headers: Person | Max FTE
Row N+4+: Team member rows (blue input)
```

- `tbl_WorkPackages` Excel table covers the WP data rows
- Named ranges: `project_start` = Config!$B$2, `fte_rate` = Config!$B$4
- No WP share seeds anywhere

### 3. Plan sheet — flattened header + FTE block only

**Plan sheet layout:**
```
Row 1:  "Title:" B1:=Config!B1 | "Start:" D1:=Config!B2 | "Rate:" F1:=Config!B4 | "Ceiling:" H1:=SUM(Config!G8:G<last>)
        All values: green cross-ref font
        H1 also serves as the `partner_ceiling` named range target
Row 2:  (blank separator)
Row 3:  Calendar year headers (merged by calendar year: "2025", "2026", ...)
Row 4:  Month headers (M1, M2, ..., M14 with calendar dates)
Row 5+: Planned FTE per WP
        Col A: WP ID (green cross-ref to Config)
        Col B: WP Name (green cross-ref to Config)
        Cols C+: SUM formulas from Capacity (one per month)
        Val cols: Sum, Budget (green cross-ref to Config ceiling), Variance
Row N:  TOTAL row (SUM per month)
```

- Config block completely removed from Plan
- Personnel Ceiling references Config WP table directly
- `partner_ceiling` named range repointed to new cell (H1)

### 4 & 5. Project parameters

```python
PROJECT_START = datetime(2026, 11, 1)
PROJECT_MONTHS = 14
```

### 6. WP share seeds — removed

- Remove `TEAM[*]["wps"]` dict from config.py
- Remove WP share seed matrix from team.py (which becomes dead code → delete)
- Remove `read_team_shares()` from reader.py
- Capacity builder no longer seeds from shares — all month cells start at 0 or empty
- Remove share-related tests

### 7. Team sheet eliminated

- team.py deleted (or gutted — roster-building logic moves to config_sheet.py)
- Team roster section added to config_sheet.py builder
- reader.py: `read_team_sheet()` replaced by `read_team_from_config()` reading the roster section in Config
- All cross-refs that pointed to Team sheet now point to Config sheet
- Capacity person headers: `=Config!A<roster_row>` instead of `=Team!A<row>`

### 8 & 9. Calendar year grouping and framing

**Year header calculation:**
```python
# For project starting Nov 2026, 14 months:
# M1=Nov'26, M2=Dec'26 | M3=Jan'27 ... M14=Dec'27
# Groups: [M1-M2], [M3-M14]
```

**Changes to `build_year_month_headers()`:**
- Group by calendar year boundaries instead of 12-month spans
- Year labels: just `"2025"`, `"2026"`, etc. (no "Year N" prefix)
- First group may be < 12 months (partial first year)
- Last group may be < 12 months (partial last year)

**Changes to `apply_year_groups()`:**
- Same calendar-year-boundary logic for column grouping
- Collapse/expand based on which calendar year contains `current_month`

**Changes to Summary sheet:**
- Row labels: `"2026"`, `"2027"` instead of `"Year 1"`, `"Year 2"`...
- Each calendar year row sums the months belonging to that calendar year
- Formulas adjust for partial years (first/last year may have fewer months)

**Changes to all sheets with month columns (Plan, Capacity, Actuals):**
- Year merged header row uses calendar years
- Column groups align to calendar year boundaries

### 10. Fix Trends E14 TOTAL (#VALUE!)

The SUMPRODUCT formula uses `IF(ISNUMBER(...))` inside SUMPRODUCT. This is an implicit array formula that can fail in LibreOffice and older Excel.

**Fix:** Wrap with `_xlfn._xlws.ISNUMBER` or restructure to avoid IF inside SUMPRODUCT:
```excel
=SUMPRODUCT(
    (E5:E13)*ISNUMBER(E5:E13),
    Actuals!$M$4:$M$12
)/partner_ceiling
```
Using `*ISNUMBER()` as a multiplier (1 for numbers, 0 for text) avoids the IF array issue.

Also verify array sizes match (Trends WP rows count == Actuals WP rows count).

### 11. Remove --skeleton flag

- Remove `--skeleton` from CLI argument parser in `__main__.py`
- Remove skeleton logic from `init_workbook()` and `rebuild_workbook()`
- `init` always produces empty Capacity
- `rebuild` always preserves existing Capacity/Actuals values
- Remove skeleton-related tests

### 12. Demo mode — `init --demo`

**CLI:** `python -m budget_planner init --demo`

**Demo capacity assignments in config.py:**
```python
DEMO_CAPACITY_ASSIGNMENTS = [
    {"person": "Alice", "wp": "WP1", "start": 1, "end": 14, "fte": 0.10},
    {"person": "Alice", "wp": "WP2", "start": 1, "end": 6, "fte": 0.50},
    {"person": "Alice", "wp": "WP4", "start": 6, "end": 14, "fte": 0.40},
    {"person": "Bob",   "wp": "WP5", "start": 1, "end": 30, "fte": 0.80},
    # ... etc
]
```

**Behavior:**
- `init --demo` calls `init_workbook()` then `_seed_demo_data(wb)`
- `_seed_demo_data()` fills Capacity cells from `DEMO_CAPACITY_ASSIGNMENTS`
- Actuals seeded from existing `DEMO_ACTUALS_MULTIPLIERS` (applied to planned EUR)
- Legacy `python -m budget_planner` (no subcommand) removed

---

## Dependency Graph

```
Step 1:  [1] Disable protection         (independent)
Step 2:  [4,5] Config values             (independent)
Step 3:  [6] Remove share seeds          (independent)
Step 4:  [7] Merge Team → Config         (depends on 3, 6)
Step 5:  [2] Remove Plan config block    (depends on 4 — needs Config WP table finalized)
Step 6:  [3] Flatten Plan header         (depends on 5 — needs Plan layout finalized)
Step 7:  [8,9] Calendar year grouping    (independent of 1-6, touches layout.py, styles.py, summary.py)
Step 8:  [10] Fix Trends E14             (independent)
Step 9:  [11] Remove --skeleton          (depends on 6)
Step 10: [12] Demo mode                  (depends on 3, 6, 9 — needs new seeding mechanism)
Step 11: Update _META, tests, reader     (depends on all above)
```

**Parallel groups:**
- Group A (independent): Steps 1, 2, 3, 7, 8
- Group B (after A): Steps 4, 9
- Group C (after B): Steps 5, 10
- Group D (after C): Steps 6
- Group E (after all): Step 11
