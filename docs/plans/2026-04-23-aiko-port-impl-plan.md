# Aiko-Port Implementation Plan

**Status**: DRAFT — ready for `/impag`.
**Created**: 2026-04-23.
**Source brief**: `docs/plans/2026-04-23-aiko-port-decisions.md` (locked Q-A…Q-F answers; rationale + alternatives live there — do not re-litigate).
**Depends on**: solver freeze-point (shipped); Stage B of indirect-view-restructure is orthogonal (different sheets — Actuals/Plan-Forecast/Capacity/Actuals-Hours; no file overlap with this plan's edits to `config_sheet.py` + `timelog_plan.py`).

## Scope

Port four Aiko-workbook features forward into the builder so round-trip rebuild preserves them:

- **T1 (Q-F)** — `Config!F13:F20` become structured-ref formulas (from literal seeds).
- **T2 (Q-B)** — Timelog Plan WP-summary `B6:B13` direct-xref Config `F{row}`; `C14 = SUM(C6:C13)` raw; `C15 = SUM(B6:B13)`.
- **T3 (Q-A)** — Extend per-year band from 4 cols to 5 cols (`J:N` = Y1…Y5); add `I15 = "Yearly Budget +i"` row with schedule-derived proportional formula.
- **T4 (Q-D)** — Add row-30 per-month person-cost row above PER-PERSON SUMMARY (formula refs `H33:H40`).

Out of scope (per brainstorm §"Session-2 answers"): Q-C (no block shift / no title rename), Q-E (manual data edits — user handles live).

## Tasks

### T1 — Config F13:F20 structured-ref formula

**File**: `budget_planner/config_sheet.py` (~lines 190–200).
**Edit**: Replace literal write at `COL_WP_BUDGET_PLUS_I` with structured-ref formula:
```
=tbl_WorkPackages[[#This Row],[FTE Budget]]*tbl_WorkPackages[[#This Row],[Rate (€/PM) -i]]*1.25
```
`tbl_WorkPackages` already registered at `config_sheet.py:287` with those named columns — formula will resolve on recalc.
**Tests**: `tests/test_config_sheet.py` — update any assertion expecting a literal in `F13:F20`; add one asserting structured-ref substring. Run recalc (`scripts/recalc.py`) and confirm `data_only=True` read still returns the same numeric.

### T2 — WP-summary B-col direct Config xref + TOTAL-row reshape

**File**: `budget_planner/timelog_plan.py` — `_write_wp_summary_row` (~1151), `_write_monthly_total_row` / TOTAL-row writer for C14/C15.
**Edit**:
- `B6:B13` per-WP row: write `=Config!$F{row}` (where `{row}` = 13 + WP index, relative so WP-row maintenance flows). Header B5 = "Budget +i" (keep current wording; confirm styling matches `+i` fill per `budget_planner/styles.py`).
- `C14`: `=SUM(C6:C13)` raw (drop any inline `*1.25` on this aggregate).
- `C15`: `=SUM(B6:B13)` (B-col is now Budget +i direct; total is straight sum).
- Row-14/15 labels at B: `B14 = "TOTAL Assigned +i"`, `B15 = "TOTAL Budget +i"` (verify current text; patch if drifted).

**Tests**: `tests/test_timelog_plan.py` — update B6:B13 formula assertions from prior seed/formula shape to `=Config!$F{n}`; update C14/C15 assertions to raw SUMs. Run recalc; confirm totals numerically match prior with-indirect values.

### T3 — Extend yearly band J:N (5 years) + "Yearly Budget +i" row

**Files**: `budget_planner/layout.py`, `budget_planner/timelog_plan.py` (`_year_col_specs` ~855, `_write_yearly_rows` ~982, `_write_wp_summary_row` ~1151).

**Edit**:
- Confirm year count: current band currently spans 4 cols (I:L per brainstorm); expected 2 for Y2026–Y2027. Audit `_year_col_specs` — if it computes dynamic year count from project start/end, the 5-col band may already emerge when project spans 5 years. If dynamic width is correct but anchor shifts from `I` → `J`, adjust `TIMELOG_YEARLY_LABEL_COL` accordingly. Decide based on audit.
- `I14 = "Yearly Assigned +i"` (label col), `J14:N14 = =SUM(J6:J13) … =SUM(N6:N13)`.
- `I15 = "Yearly Budget +i"` (new row), `J15:N15 = schedule-derived proportional Budget +i`. Formula shape per brainstorm §"Q-A coefficient-shape sub-answer": fraction = `(months of WPn overlapping yearY) / (total active months of WPn)`, counts derived from WP BUDGET `H/I` start/end (cols in Config WP table). Per WP, sum across rows: `Σ fraction × Config!$F{wp_row}`.
  - Implementation sketch: `=SUMPRODUCT( ((Config!$H$13:$H$20<=<year_end>)*(Config!$I$13:$I$20>=<year_start>)) * <overlap_months_per_wp> / <total_months_per_wp> * Config!$F$13:$F$20 )` — resolve exact month arithmetic during implementation; fallback is a row-explicit per-WP sum.
- Per-WP year cols `J6:N13`: shape `=ROUND(SUMPRODUCT((<year_cols>="WPx")*H33:H40)*1.25,0)` per brainstorm. `<year_cols>` = monthly grid cols B20:AS20 filtered to that year. Current `_year_col_specs` likely already emits this against the old anchor — just verify the col-letter wiring after T3's width change.

**Tests**: `tests/test_timelog_plan.py` — add/update yearly-band width assertions; add "Yearly Budget +i" row-15 label + formula-shape assertion; verify J14:N14 / J15:N15 numerical results via recalc.

### T4 — Row-30 per-month person-cost row

**File**: `budget_planner/timelog_plan.py` (write-step above PER-PERSON SUMMARY header; PER-PERSON SUMMARY header is at row 32 per Q-C-skip decision — row 30 lands in the gap row 29–31 currently empty).

**Edit**:
- Insert writer for row 30: cols `B:AS`, formula `=ROUND(SUMPRODUCT((<col>21:<col>28<>"")*H33:H40)*1.25,0)` where `<col>` = each month column letter (B..AS).
- Row-label in col A (e.g. `A30 = "Monthly person cost +i"` or similar — pick label consistent with other row labels; note in commit).
- Styling: match monthly-total +i row conventions (`WITH_INDIRECT_HEADER_FILL` per T7.3 pattern) — green/orange per `styles.py`.

**Tests**: `tests/test_timelog_plan.py` — add assertion for row-30 label + formula shape across B:AS; recalc & confirm numerics match aggregated PER-PERSON SUMMARY totals when summed across months.

## Dependency order

1. **T1** is independent (Config-only). Ship first.
2. **T2** (Timelog Plan B-col) is independent of T3/T4 — ship second.
3. **T3** (yearly band extension) changes col-letter anchors; **ship before T4** so T4's `<col>21:<col>28` letter sweep doesn't get re-spelled.
4. **T4** last.

All four tasks together = 4 implementation tasks → single `/impag` session per skill sizing rule. No stage split.

## Test & recalc gate

After each task: targeted `pytest tests/test_timelog_plan.py -q` (or `test_config_sheet.py` for T1) + `pyright`.
Before final commit: full `pytest tests/ -q` + `python scripts/recalc.py output/budget_plan.xlsx` + inspect `F13:F20`, `B6:B13`, `J:N` band, row 30 in the recalc'd workbook via `mcp__openpyxl-read__get_cell_details` on a handful of spot cells.

## Invariants to preserve

- Named ranges `project_start`, `fte_rate` untouched.
- Font color convention: B6:B13 (`=Config!$F{n}`) = **green** cross-sheet ref; J15:N15 schedule-derived = **black** formula; F13:F20 structured-ref = **black** formula.
- Inactive month cells on row 30 still greyed/empty (inherit grey-out mask from existing monthly-total row if current code does so; else apply same mask).
- `docs/design-invariants.md` — check whether the shift from "Budget +i seeded literal" to "Budget +i formula" needs a new invariant entry (trust-boundary / derived-metric rule). Likely yes for T1.

## Post-ship

- Update `memory/project_state.md`: mark Aiko-port shipped with commit range; drop "impl plan not yet written" note.
- `git mv` this file + `docs/plans/2026-04-23-aiko-port-decisions.md` → `docs/plans/archived/` after ship (plan-hygiene rule: `git add` working-tree amendments BEFORE `git mv`).
- Any load-bearing §Design content not in a living doc: migrate to `docs/internals.md` or `docs/design-invariants.md` before archive.
