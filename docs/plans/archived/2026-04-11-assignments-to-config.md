# Move Assignments & External Commitments to Config Sheet

**Date**: 2026-04-11
**Status**: Design validated

## Motivation

- **Single-source-of-truth UX**: all setup inputs on one sheet
- **Rebuild reliability**: Config reader is simpler/more robust than Capacity reader
- **Visual proximity**: WP-level PM budget next to person-level PM assignments enables at-a-glance reconciliation
- **Capacity simplification**: becomes purely a monthly planning grid, no config blocks

## Config Sheet Layout (new order)

```
Row 1-5:    Project Parameters (unchanged)

Section B:  WP BUDGET (rows 7+)
  Headers:  WP ID | WP Name | FTE Budget | Rate (EUR/PM) | Ceiling (EUR) | Assigned PM | Variance
  Data:     One row per WP (blue input cols A-D, formula cols E-G)
  TOTAL:    SUM row on cols C, E, F, G
  Table:    tbl_WorkPackages covers A:G (header + data, excludes TOTAL)

Section B2: PERSON-WP ASSIGNMENTS (new, below WP BUDGET)
  Headers:  Person | WP | Total FTE | Start Month | End Month | Sig Assigned | Sig External | Balance
  Data:     Blue input cols A-E, formula cols F-H
  All rows blue input for core data; formula columns show person-level aggregates

Section B3: EXTERNAL COMMITMENTS (new, below assignments)
  Headers:  Person | Project | Total FTE | Start Month | End Month | EU Grant
  Data:     All blue input

Section C:  WP SCHEDULE (unchanged, shifts down)
Section D:  TEAM ROSTER (unchanged, shifts down)
Section E:  REPORTING PERIODS (unchanged, shifts down)
Section F:  RATE BY GRADE (unchanged, stays at bottom)
```

## Validation Formulas

### WP Budget table (cols F-G)

| Column | Header | Formula | Font |
|--------|--------|---------|------|
| F | Assigned PM | `=SUMIFS(assign_ftes, assign_wps, A<row>)` | Black |
| G | Variance | `=C<row>-F<row>` | Black |

CF: yellow fill on col G when `ABS(G<row>) > 0.01`.

TOTAL row: SUM on cols F and G.

### Person-WP Assignments table (cols F-H)

| Column | Header | Formula | Font |
|--------|--------|---------|------|
| F | Sig Assigned | `=SUMIFS(assign_ftes, assign_persons, A<row>)` | Black |
| G | Sig External | `=SUMIFS(ext_ftes, ext_persons, A<row>)` | Black |
| H | Balance | `=VLOOKUP(A<row>, roster_name_fte_range, 2, FALSE) * project_months - F<row> - G<row>` | Black |

Values repeat for rows with the same person (intentional — in-place context).

CF on Balance: red if < 0 (over-committed), yellow if < 10% of available PM.

## Named Ranges (new)

| Named Range | Points To | Purpose |
|------------|-----------|---------|
| `assign_persons` | Config!$A$first:$A$last (assignment Person col) | SUMIFS lookup |
| `assign_wps` | Config!$B$first:$B$last (assignment WP col) | WP Budget Assigned PM |
| `assign_ftes` | Config!$C$first:$C$last (assignment Total FTE col) | Both tables |
| `ext_persons` | Config!$A$first:$A$last (ext commitment Person col) | Assignment Sig External |
| `ext_ftes` | Config!$C$first:$C$last (ext commitment Total FTE col) | Assignment Sig External |

## Capacity Sheet Changes

Remove the two collapsed config blocks at the top. Sheet starts directly with year/month headers (row 1-3), per-person grids (row 4+).

The builder interface is unchanged — still receives `assignments` and `external_commitments` as parameters.

## Code Changes

### Phase 0: Preparatory Refactoring

Decompose the most complex functions *before* adding new functionality.
See `docs/analysis/2026-04-11-simplifier-report.md` for full analysis.

#### 0a. Column constants (`config_sheet.py`, `reader.py`)

Define named column constants per section so builders and readers share one source of truth.
Existing magic numbers (3, 4, 5) become `COL_WP_FTE`, `COL_WP_RATE`, `COL_WP_CEILING`, etc.
New sections (assignments, external) use constants from day one.

#### 0b. Extract section builders from `build_config_sheet` (`config_sheet.py`)

Current: single 283-line function, CC=20. Split into:
- `_write_budget_section(ws, row, ...) -> next_row`
- `_write_schedule_section(ws, row, ...) -> next_row`
- `_write_roster_section(ws, row, ...) -> next_row`
- `_write_reporting_section(ws, row, ...) -> next_row`
- `_write_rate_by_grade_section(ws, row, ...) -> next_row`

Top-level function becomes a coordinator. New sections added in Phase 1 follow the same pattern.

#### 0c. Split `read_capacity_sheet` (`reader.py`)

Current: CC=40, CogC=73, nesting=6. Extract per-section readers:
- `_read_assignments_from_capacity(ws, ...) -> list[dict]`
- `_read_externals_from_capacity(ws, ...) -> list[dict]`
- `_read_person_grids(ws, ...) -> dict`

Top-level detects section headers and delegates. Backward compat (Phase 1) becomes a clean dispatch, not more branches.

#### 0d. Extract section builders from `build_capacity_sheet` (`capacity.py`)

Current: CC=50, CogC=76, 356 lines, 11 params (2 unused). Split into:
- `_write_capacity_config_blocks(ws, row, ...) -> next_row` (removed in Phase 1)
- `_write_person_grid(ws, row, member, ...) -> next_row`
- `_write_capacity_totals(ws, row, ...) -> next_row`

Remove unused params `preserved` and `wb`. Clean up manual index tracking → `enumerate()`.

### Phase 1: Feature Implementation

#### `config_sheet.py`
- Accept `assignments` and `external_commitments` parameters
- Add `_write_assignments_section()` and `_write_external_section()` (same pattern as 0b)
- Add validation formula columns to WP Budget table (expand to A:G)
- Add validation formula columns to assignments table (cols F-H)
- Create new named ranges
- Return additional layout keys: `assign_header_row`, `assign_first_row`, `assign_last_row`, `ext_header_row`, `ext_first_row`, `ext_last_row`

#### `capacity.py`
- Remove `_write_capacity_config_blocks()` call (extracted in 0d, now deleted)
- Per-person grid building unchanged

#### `reader.py`
- Add Config-based reader for assignments and external commitments (detect by section header text)
- Backward compatibility: if Config has no "PERSON-WP ASSIGNMENTS" header, fall back to capacity sub-readers (extracted in 0c)

#### `workbook.py`
- `init_workbook()`: pass assignments/commitments to `build_config_sheet()` (in addition to capacity builder)
- `rebuild_workbook()`: read assignments/commitments from Config reader (not Capacity), pass to both builders

#### `meta.py`
- Update Config section documentation (new sections, expanded WP Budget columns)
- Remove Capacity config block documentation
- Update data flow description

## Migration / Edge Cases

- **Old workbooks**: reader checks Config for section header first, falls back to Capacity reader. Rebuild works on both formats.
- **Empty tables**: sections written with headers, no data rows. Named ranges point to single empty row. SUMIFS return 0.
- **Table sizing**: reader scans header down to next blank row. Named ranges recalculated on every build.
- **Column width**: `auto_fit_columns` handles wider Config sheet. Section header merges expand.
- **CLI**: no breaking changes to `init`, `init --demo`, `rebuild`.
