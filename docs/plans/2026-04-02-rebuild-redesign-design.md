# Rebuild Redesign — Integrated Solution for Top 5 Problems

**Date**: 2026-04-02
**Status**: Approved design, not yet implemented
**Revised**: 2026-04-02 (Kaizen Muda review — fixed 7 findings)

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
| config.py fate | Seed for `init` only; Config sheet + existing sheets are living source of truth |
| Config sheet structure | Single sheet: labeled cells for params, one Excel Table for work packages |

## Chosen Approach: Thin Rebuild Layer

Add `rebuild` mode that reads the existing workbook, extracts preserved data, regenerates all sheets from the Config sheet + existing Team/Capacity data, and re-injects preserved values. Minimal changes to existing builders (optional `preserved` parameter, config passed as parameters instead of module imports). Column grouping and cell protection applied as post-generation decorators.

**Why this over alternatives:**
- **Patch-in-place** (modify existing workbook surgically) rejected: openpyxl can't auto-adjust formula references on row insert, making this extremely fragile (~500 lines of formula-shifting code)
- **Two-file system** (separate input + output workbooks) rejected: worse PM UX, two files to manage

## CLI Modes

### `python -m budget_planner init`

- Creates a fresh workbook from `config.py` seed data
- Builds Config sheet with project params + WP Table, plus all other sheets
- Only time `config.py` is read
- Replaces current `python -m budget_planner` for first-time generation

### `python -m budget_planner rebuild [path]`

- Reads existing workbook (default: `output/budget_plan.xlsx`)
- Prints timestamp: `Reading from: budget_plan.xlsx (modified 2026-03-15 14:30)`
- Extracts: Config sheet -> project params + WPs, Team sheet -> roster + shares, Capacity sheet -> external commitments + FTE values, Actuals sheet -> EUR values
- Validates extracted data (hard errors abort, warnings logged)
- Regenerates all sheets from extracted data (not config.py)
- Re-injects preserved Capacity + Actuals values matched by WP ID + person name + month index
- Applies protection, column grouping, CF warnings
- Backs up original file with timestamp before overwriting
- Writes rebuild warnings to _META sheet

### `python -m budget_planner rebuild --skeleton [path]`

- Same as `rebuild` but leaves Capacity monthly cells empty for re-entry
- Preserves Actuals data but clears Capacity allocations

### `python -m budget_planner --skeleton` (legacy)

- Empty Capacity for manual entry, reads config.py
- Only useful before `init` has been run

### Unchanged

- `--snapshot` / `--restore` — file copy as before

## Where Each Data Source Lives

A key design invariant: on `rebuild`, **no data comes from config.py**. Each piece of project data has exactly one source.

| Data | Source on `init` | Source on `rebuild` |
|---|---|---|
| Project title, start, months, rate | config.py | Config sheet (labeled cells) |
| Work packages (ID, name, FTE, start, end) | config.py | Config sheet (tbl_WorkPackages) |
| Team roster (name, max FTE) | config.py | Team sheet (roster block, blue cells) |
| WP share seeds | config.py | Team sheet (share matrix, blue cells) |
| External commitments | config.py | Capacity sheet (config block, blue cells) |
| Monthly FTE allocations | Computed from seeds | Capacity sheet (blue cells, preserved) |
| Monthly EUR actuals | Empty (zero) | Actuals sheet (blue cells, preserved) |
| Layout constants (MONTH_COL_START) | config.py | Computed in Python from extracted params |

**Why team/external data is NOT on the Config sheet:**
- The Team sheet already has blue input cells for roster + share seeds (reader.py can parse them via `read_team_sheet()`)
- The Capacity sheet already has a config block for external commitments (reader.py can parse via `read_capacity_sheet()`)
- Duplicating this data on a Config sheet creates a "which is the source of truth?" problem
- Removing duplication also eliminates the "PM must manually add WP column to Team Table" gap — on rebuild, the Team sheet's share matrix is read as-is

## Config Sheet Layout

Single sheet with two sections. Simpler than original design — only project-level parameters and work package definitions.

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

### What's NOT on the Config sheet

- **Team roster + shares** — lives on Team sheet (blue input cells)
- **External commitments** — lives on Capacity sheet config block (blue input cells)

Both are read by reader.py during rebuild. The PM edits them in their natural location.

### Reading the Table in Python

Read Table data by cell position, not Table object. Recreate Table objects on every generation — openpyxl has known quirks round-tripping Table objects.

## Round-Trip Data Flow

### Step 1: Extract from existing workbook

```
read_config_sheet()       -> project params + WP list
read_team_sheet()         -> roster + share seeds (already exists in reader.py)
read_capacity_sheet()     -> external commitments + FTE values (already exists)
read_capacity_values()    -> {("Alice", "WP1"): [0.08, 0.08, ..., 0.0], ...}
read_actuals_values()     -> {("WP1",): [0, 0, ..., 500, ...], ...}  (new)
```

All data keyed by stable identifiers (WP ID + person name + month index), never cell addresses.

**Important:** reader.py functions must receive `project_months` and `month_col_start` as parameters, not import them from config.py. These values come from the Config sheet during rebuild.

### Step 2: Validate extracted data

**Hard errors (abort):**
- Duplicate WP IDs in tbl_WorkPackages
- PROJECT_MONTHS <= 0 or > 120
- FTE rate <= 0
- tbl_WorkPackages is empty
- Team roster is empty
- External commitment person not found in team roster

**Warnings (log):**
- WP start > end
- WP FTE budget <= 0
- Share seeds for a person don't sum to ~1.0 (tolerance +/-0.05)
- Orphaned preserved Capacity data (WP ID or person removed)
- Orphaned preserved Actuals data (WP ID removed)
- Truncated monthly data (PROJECT_MONTHS shrank, non-zero values dropped)

### Step 3: Generate fresh workbook

All builders receive config data as function parameters. No builder imports from config.py.

```
build_config_sheet(ws, project_params, work_packages)
build_plan_sheet(ws, wb, project_params, work_packages)
build_team_sheet(ws, wb, team, work_packages)
build_capacity_sheet(ws, wb, team_layout, project_params, work_packages,
                     team, external_commitments, preserved=capacity_data)
build_actuals_sheet(wb, layout, project_params, work_packages, preserved=actuals_data)
build_summary_sheet(wb, layout, actuals_layout, project_params, work_packages)
build_trends_sheet(ws, wb, layout, actuals_layout, capacity_layout, team_layout,
                   project_params, work_packages, team)
build_meta_sheet(ws, layout, project_params, work_packages, team, external_commitments)
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

Same pattern for actuals.py.

### Step 5: Handle orphaned data

Orphaned keys (WP IDs or person names in preserved data but not in new Config/Team) written to the _META sheet under a `═══ LAST REBUILD LOG ═══` section. Never silently dropped.

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
| Config | Project param values, tbl_WorkPackages data cells |
| Plan | None (formula-driven) |
| Team | Roster values (name, max FTE), share seed matrix |
| Capacity | Monthly FTE cells, external project cells, external config block |
| Actuals | Monthly EUR cells |
| Summary | None (formula-driven) |
| Trends | B1 (current_month) only |
| _META | None |

Implementation:
- `style_cell()` gains `unlocked` parameter; set `cell.protection = Protection(locked=False)` for blue input cells
- After all builders complete, `workbook.py` loops sheets: `ws.protection.set_password("grant")`
- Password documented in _META sheet

## Validation & CF Warnings

### Rebuild-time validation (Python)

**Hard errors (abort):** duplicate WP ID, PROJECT_MONTHS <= 0 or > 120, FTE rate <= 0, external person not in team, empty WP table or team roster.

**Warnings (generate but log):** WP start > end, FTE budget <= 0, shares don't sum to ~1.0, orphaned preserved data, truncated monthly data.

Warnings printed to terminal and written to _META sheet rebuild log section.

### CF warnings in workbook

Three new rules extending existing CF:

- **Capacity**: Yellow if Sum > Budget in per-person WP validation column
- **Config**: Yellow if WP start > end (`=D7>E7`)
- **Team**: Yellow if share row doesn't sum to ~1.0 (helper SUM column, already has this CF)

All existing CF rules unchanged.

## Pre-Mortem Mitigations

| Risk | Mitigation |
|---|---|
| openpyxl corrupts Excel Tables on round-trip | Read cell data, recreate Table objects on every generation |
| PM renames WP ID, loses Capacity data | Rebuild emits WARNING, writes orphaned data to _META rebuild log |
| reader.py misidentifies person headers after resave | Dual detection: style-based primary, pattern-based fallback (scan for TOTAL rows) |
| PM forgets to save Excel before rebuild | CLI prints source file last-modified timestamp |
| Formula references break on row count change | Match by WP ID + person name, never by row number |
| Builder still imports config.py | All builders receive config as parameters; no module-level config imports in rebuild path |
| reader.py uses config.py constants | All reader functions receive project_months and month_col_start as parameters |

## Impact on Existing Code

### New files

| File | Purpose | ~Lines |
|---|---|---|
| `budget_planner/config_sheet.py` | Build Config sheet (params + tbl_WorkPackages) + read Config sheet back | ~100 |

### Modified files

| File | Change | ~Delta |
|---|---|---|
| `workbook.py` | `init_workbook()` + `rebuild_workbook()` orchestration, config parameter passing | +100 |
| `__main__.py` | `init` and `rebuild` subcommands (with `--skeleton`), timestamp print | +25 |
| `capacity.py` | Accept config as params (not imports), `preserved` optional parameter | +25 |
| `actuals.py` | Accept config as params, `preserved` optional parameter | +20 |
| `plan.py` | Accept config as params, named ranges point to Config sheet | +20 |
| `team.py` | Accept config as params | +15 |
| `summary.py` | Accept config as params | +15 |
| `trends.py` | Accept config as params | +20 |
| `reader.py` | `read_config_sheet()`, `read_actuals_values()`, parameterize constants, hardened person-header detection | +100 |
| `styles.py` | `apply_year_groups()`, `unlocked` parameter on `style_cell()` | +30 |
| `meta.py` | Accept config as params, document Config sheet, rebuild log section | +30 |

### New tests (~40)

- Config sheet round-trip (write Table, save, reopen, read back)
- Rebuild preservation (generate, modify Capacity values, rebuild with new WP, old values preserved)
- Orphaned data (remove WP from Config, rebuild, check _META rebuild log)
- Validation (bad Config data -> hard error with clear message)
- Column grouping (groups exist, correct collapse state)
- Cell protection (formula cells locked, blue cells unlocked)
- CF warnings (overallocation highlighted on Config sheet)
- Builder parameterization (no builder imports config.py at module level in rebuild path)

### Unchanged

- `config.py` — stays as seed for `init`, untouched

### Total new/changed code

~600 lines production + ~300 lines tests

## Kaizen Review Findings (2026-04-02)

Muda analysis identified 7 issues in the original design. All addressed in this revision:

| # | Type | Finding | Resolution |
|---|---|---|---|
| 1 | Overproduction | Config tbl_Team/tbl_External duplicated Team/Capacity sheets | Removed from Config; read from existing sheets |
| 2 | Defect | summary.py/trends.py still imported config.py | All 7 builders decoupled; receive config as params |
| 3 | Defect | reader.py imported config.py constants | Reader functions parameterized |
| 4 | Over-processing | _REBUILD_LOG as separate sheet | Merged into _META as rebuild log section |
| 5 | Waiting | Manual WP column in Team Table | Eliminated by fix #1 (no tbl_Team) |
| 6 | Inventory | Stale --skeleton after init | Added `rebuild --skeleton` mode |
| 7 | Motion | LOC estimate too low (~425) | Updated to ~600 |

## Out of Scope

- Multi-partner support (separate project)
- VBA macros
- Web dashboard or Google Sheets output
- Reading Actuals from external sources (CSV import)
- Scenario naming/diffing beyond snapshot/restore
