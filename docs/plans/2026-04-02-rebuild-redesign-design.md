# Rebuild Redesign — Integrated Solution for Top 5 Problems

**Date**: 2026-04-02
**Status**: Approved design, not yet implemented

## Goals

Solve five interconnected problems in one integrated redesign:

1. **Config-as-Python is a dead end** — non-technical PMs can't modify project structure
2. **No round-trip** — regeneration destroys manual Capacity/Actuals data
3. **60-column navigation** — horizontal scrolling nightmare
4. **No cell protection** — accidental formula overwrites
5. **No input validation** — bad config produces broken workbooks silently

## Design Decisions

| Question | Decision |
|---|---|
| Target user model | Hybrid: developer runs `init`, PM owns workbook day-to-day |
| How PM makes structural changes | Excel-as-config: edit Config sheet, run `rebuild` |
| What gets preserved on rebuild | Capacity FTE allocations + Actuals EUR values |
| Navigation solution | Column grouping by year, auto-collapse to current year |
| Protection model | All sheets protected, blue input cells unlocked, password "grant" |
| Validation approach | Hard errors at rebuild time + CF warnings in workbook |
| config.py fate | Seed for `init` only; Config sheet is living source of truth after that |
| Config sheet structure | Single sheet: labeled cells for params, Excel Tables for WP/Team/External |

## Chosen Approach: Thin Rebuild Layer

Add `rebuild` mode that reads the existing workbook, extracts preserved data, regenerates all sheets from the Config sheet, and re-injects preserved values. Minimal changes to existing builders (optional `preserved` parameter). Column grouping and cell protection applied as post-generation decorators.

**Why this over alternatives:**
- **Patch-in-place** (modify existing workbook surgically) rejected: openpyxl can't auto-adjust formula references on row insert, making this extremely fragile (~500 lines of formula-shifting code)
- **Two-file system** (separate input + output workbooks) rejected: worse PM UX, two files to manage

## CLI Modes

### `python -m budget_planner init`

- Creates a fresh workbook from `config.py` seed data
- Builds Config sheet with Excel Tables + all other sheets
- Only time `config.py` is read
- Replaces current `python -m budget_planner` for first-time generation

### `python -m budget_planner rebuild [path]`

- Reads existing workbook (default: `output/budget_plan.xlsx`)
- Prints timestamp: `Reading from: budget_plan.xlsx (modified 2026-03-15 14:30)`
- Extracts: Config sheet -> project structure, Capacity -> FTE values, Actuals -> EUR values
- Validates Config (hard errors abort, warnings logged)
- Regenerates all sheets from Config sheet data (not config.py)
- Re-injects preserved Capacity + Actuals values matched by WP ID + person name + month index
- Applies protection, column grouping, CF warnings
- Backs up original file with timestamp before overwriting
- Writes orphaned data to `_REBUILD_LOG` sheet

### Unchanged

- `python -m budget_planner --skeleton` — empty Capacity, reads config.py
- `--snapshot` / `--restore` — file copy as before

## Config Sheet Layout

Single sheet with three sections stacked vertically.

### Section A: Project Parameters (rows 1-5)

Simple labeled cells, not a Table. Blue input values.

```
Row 1:  Project Title    | Programme Project Budget Plan
Row 2:  Project Start    | 2026-07-01
Row 3:  Project Months   | 60
Row 4:  FTE Rate (EUR)   | 7500
Row 5:  (empty separator)
```

Named ranges `project_start` and `fte_rate` point here (moved from Plan sheet).

### Section B: Work Packages Table (~rows 6-15)

Excel Table named `tbl_WorkPackages`:

| WP ID | WP Name | FTE Budget | Start Month | End Month |
|---|---|---|---|---|
| WP1 | Project Management | 12.0 | 1 | 60 |
| WP2 | Requirements & Design | 8.0 | 1 | 18 |

PM adds a row to the Table, Excel auto-expands, run `rebuild`.

### Section C: Team Table (~rows 18-25)

Excel Table named `tbl_Team`:

| Person | Max FTE | WP1 | WP2 | WP3 | ... |
|---|---|---|---|---|---|
| Alice | 1.0 | 0.25 | 0.50 | 0.25 | ... |

WP columns match WP IDs from tbl_WorkPackages. When a WP is added, PM manually adds a column here.

### Section D: External Commitments Table (~rows 28-32)

Excel Table named `tbl_External`:

| Person | Project | Total FTE | Start Month | End Month |
|---|---|---|---|---|
| Alice | ProjectX | 3.0 | 1 | 11 |

### Reading Tables in Python

Read Table data by cell position, not Table object. Recreate Table objects on every generation — openpyxl has known quirks round-tripping Table objects.

## Round-Trip Data Flow

### Step 1: Extract from existing workbook

```
read_config_sheet()     -> project params + WP list + team list + external commitments
read_capacity_values()  -> {("Alice", "WP1"): [0.08, 0.08, ..., 0.0], ...}
read_actuals_values()   -> {("WP1",): [0, 0, ..., 500, ...], ...}
```

All data keyed by stable identifiers (WP ID + person name + month index), never cell addresses.

### Step 2: Validate extracted Config

**Hard errors (abort):** duplicate WP IDs, empty tables, zero/negative rate or months, external person not in team.

**Warnings (log):** WP start > end, FTE <= 0, shares don't sum to ~1.0, orphaned preserved data, truncated monthly data.

### Step 3: Generate fresh workbook

Calls existing builders with extracted config data instead of config.py imports:

```
build_config_sheet()
build_plan_sheet()
build_team_sheet()
build_capacity_sheet(..., preserved=capacity_data)
build_actuals_sheet(..., preserved=actuals_data)
build_summary_sheet()
build_trends_sheet()
build_meta_sheet()
```

### Step 4: Re-injection logic

Preserved data checked first in the monthly cell loop:

```python
if preserved and (name, wp_id) in preserved:
    seed_value = preserved[(name, wp_id)][month_index]
elif mode == "skeleton":
    seed_value = None
else:
    seed_value = round(share * wp["fte"] / active_months, 4)
```

### Step 5: Handle orphaned data

Orphaned keys (WP IDs or person names in preserved data but not in new Config) written to `_REBUILD_LOG` sheet with counts and context. Never silently dropped.

### Month count changes

- PROJECT_MONTHS grows (60 -> 72): months 1-60 preserved, 61-72 empty
- PROJECT_MONTHS shrinks (60 -> 48): months 1-48 preserved, 49-60 logged if non-zero

## Column Grouping

Applied to Capacity, Plan, and Actuals sheets after generation.

- Months grouped by year at outline level 1 (12 columns per group)
- On `rebuild`: group containing `current_month` expanded, all others collapsed
- On `init`: Year 1 expanded, rest collapsed
- Utility function `apply_year_groups(ws, current_month)` in `styles.py`
- Graceful degradation: if LibreOffice handles grouping differently, data is still correct

## Cell Protection

All sheets protected after generation. Only blue input cells unlocked.

| Sheet | Unlocked cells |
|---|---|
| Config | All input cells (params, Table data) |
| Plan | None (formula-driven) |
| Team | Roster values, share seed matrix |
| Capacity | Monthly FTE cells, external project cells |
| Actuals | Monthly EUR cells |
| Summary | None (formula-driven) |
| Trends | B1 (current_month) only |
| _META | None |
| _REBUILD_LOG | None |

Implementation:
- `style_cell()` gains `unlocked` parameter; set `cell.protection = Protection(locked=False)` for blue input cells
- After all builders complete, `workbook.py` loops sheets: `ws.protection.set_password("grant")`
- Password documented in _META sheet

## Validation & CF Warnings

### Rebuild-time validation (Python)

**Hard errors (abort):** duplicate WP ID, PROJECT_MONTHS <= 0 or > 120, FTE rate <= 0, external person not in team, empty WP or team tables.

**Warnings (generate but log):** WP start > end, FTE budget <= 0, shares don't sum to ~1.0, orphaned preserved data, truncated monthly data.

Warnings printed to terminal and written to `_REBUILD_LOG` sheet.

### CF warnings in workbook

Three new rules extending existing CF:

- **Capacity**: Yellow if Sum > Budget in per-person WP validation column
- **Config**: Yellow if WP start > end (`=D7>E7`)
- **Config**: Yellow if team share row doesn't sum to ~1.0 (helper SUM column)

All existing CF rules unchanged.

## Pre-Mortem Mitigations

| Risk | Mitigation |
|---|---|
| openpyxl corrupts Excel Tables on round-trip | Read cell data, recreate Table objects on every generation |
| PM renames WP ID, loses Capacity data | Rebuild emits WARNING, writes orphaned data to _REBUILD_LOG |
| reader.py misidentifies person headers after resave | Dual detection: style-based primary, pattern-based fallback (scan for TOTAL rows) |
| PM forgets to save Excel before rebuild | CLI prints source file last-modified timestamp |
| Formula references break on row count change | Match by WP ID + person name, never by row number |

## Impact on Existing Code

### New files

| File | Purpose | ~Lines |
|---|---|---|
| `budget_planner/config_sheet.py` | Build + read Config sheet (Tables + labeled cells) | ~150 |

### Modified files

| File | Change | ~Delta |
|---|---|---|
| `workbook.py` | `init_workbook()` + `rebuild_workbook()` orchestration | +80 |
| `__main__.py` | `init` and `rebuild` subcommands, timestamp print | +20 |
| `capacity.py` | `preserved` optional parameter, one `if` branch | +15 |
| `actuals.py` | `preserved` optional parameter, one `if` branch | +15 |
| `reader.py` | `read_config_sheet()`, `read_actuals_values()`, hardened person-header detection | +80 |
| `styles.py` | `apply_year_groups()`, `unlocked` parameter on `style_cell()` | +30 |
| `plan.py` | Named ranges point to Config sheet, globals become formula refs | +10 |
| `meta.py` | Document Config sheet, rebuild workflow, protection, _REBUILD_LOG | +20 |
| `team.py` | Read from passed-in config data instead of config.py import | ~5 |

### New tests (~40)

- Config sheet round-trip (write Tables, save, reopen, read back)
- Rebuild preservation (generate, modify values, rebuild, verify preserved)
- Orphaned data (remove WP, rebuild, check _REBUILD_LOG)
- Validation (bad Config -> hard error with message)
- Column grouping (groups exist, correct collapse state)
- Cell protection (formula cells locked, blue cells unlocked)
- CF warnings (overallocation highlighted on Config sheet)

### Unchanged

- `summary.py`, `trends.py`, `config.py` — no changes
- All 146 existing tests pass (builders get optional params with None defaults)

### Total new/changed code

~425 lines production + ~300 lines tests

## Out of Scope

- Multi-partner support (separate project)
- VBA macros
- Web dashboard or Google Sheets output
- Reading Actuals from external sources (CSV import)
- Scenario naming/diffing beyond snapshot/restore
