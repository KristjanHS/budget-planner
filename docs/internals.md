# Internals

Technical reference for the workbook generation engine: builder pipeline, reader/rebuild, layout infrastructure, conditional formatting, and cross-sheet references.

---

## Sheet Builders

Each sheet is produced by a stateless builder function. Builders receive data as parameters and return a layout dict of row/column positions consumed by downstream builders.

### Builder Interface

```python
def build_<sheet>_sheet(ws, wb, ..., *, project_params=None, work_packages=None, ...):
    # ws: openpyxl Worksheet
    # wb: parent Workbook (named ranges, cross-sheet access)
    # ...: upstream layout dicts
    # Returns: layout dict with row/column numbers for downstream builders
```

### Builder Dependency Chain

| Builder | Needs from upstream |
|---------|-------------------|
| config_sheet | (none) |
| plan | config_layout.wp_first_row |
| capacity | config_layout.person_rows, plan_layout |
| plan (patch) | capacity_layout.person_groups.wp_rows |
| timelog_plan | config_layout + Trends-stub `current_month` named range + `team_priorities` / `roster_lookup` named ranges |
| actuals | plan_layout |
| summary | plan_layout + actuals_layout |
| trends | all upstream layouts |

### Per-Sheet Notes

**Config** (`config_sheet.py`): Stacked sections — blue-input project-parameters block (Title, Start, Months, FTE Rate, EU Ceiling, Hours/FTE Month, Vacation Month), black-formula builder-written scalars (`Num Vacation Months`, `Version Stamp` — Rule 3 carve-out, see `.claude/rules/dev-workflow.md`), then the WP table (`tbl_WorkPackages`), then Team Roster. Source of truth for every named range listed in `docs/design-invariants.md` § Named-range contract except the two through-month names (hosted on Actuals / Actuals-Hours). Returns `config_layout` with `wp_first_row`, `wp_last_row`, `person_rows`.

**Plan** (`plan.py`): Two-phase build. Phase 1 creates structure with empty FTE cells. Phase 2 (`_apply_planned_formulas()`) fills cells with `=SUM(Capacity!...)` after Capacity is built. Has validation columns: Sum, Budget FTE, Variance.

**Capacity** (`capacity.py`): Per-person groups with WP rows (blue input, distribution-curve static values), TOTAL, REMAINING (Max - TOTAL), and **EU TOTAL** (programme WPs + EU-flagged externals — tracks the 215-day annual ceiling separately from TOTAL). Non-obvious conventions: per-person WP rows cover **only the assigned WPs** (from Config PERSON-WP ASSIGNMENTS), not every WP; distribution curve type comes from the Type col on the merged WP BUDGET row; the right-side raw-monthly band sits at `outline_level=1` (default collapsed); the always-visible TO-DATE `Variance PM` column is populated only on TOTAL + EU TOTAL rows. Returns `capacity_layout` with `person_groups` + `todate_col_actual` / `todate_col_var`. See `capacity.py` for the formulas and CF rules.

**Timelog Plan** (`timelog_plan.py`): Prescriptive whole-month 1.0-FTE WP allocation grid — separate allocation surface from Plan/Capacity. See the `timelog_plan.py` module docstring + `docs/plans/archived/2026-04-18-timelog-plan-design.md` (v1), `2026-04-19-timelog-plan-v2-design.md` (merged grid + toggle), `2026-04-19-timelog-plan-enhancements-design.md` (phases 3–4 totals + vacation). Cross-sheet invariants in `docs/design-invariants.md`.

**Actuals** (`actuals.py`) + **Actuals-Hours** (`actuals_hours.py`): EUR and hours input surfaces (two cadences). Builder internals live in module docstrings; cross-sheet trust boundaries and through-month invariants are canonical in `docs/design-invariants.md` § Two-cadence actuals.

**Plan-Forecast** (`plan_forecast.py`): WP x month forecast-EUR grid — pure aggregation, read-only (every cell is a formula). Shape mirrors Plan. Each monthly cell `=SUM(<Actuals-Hours forecast-EUR leaves for this WP x month across all assignments>)`. No user inputs.

**Summary** (`summary.py`): Calendar-year column groups (5 sub-columns/year: Plan FTE, Plan EUR, Actual EUR, Forecast EUR, Variance EUR) + PROJECT TOTAL. Budget Overview applies the **25% indirect uplift** to Total Claimable. Reporting Period Summary emits **Form C preparation rows** grouped by Config reporting periods — compliance-load-bearing, not free to restructure. Per-row specifics in `summary.py`.

**Trends** (`trends.py`): Dashboard — WP Burn Rate & Exhaustion table, cumulative EUR LineChart (Planned / Actual / Forecast / Ceiling / RP End; Actual + Forecast series clip via `NA()` past their respective through-month), **Reconciliation: Forecast vs Actual EUR** section, per-WP Gap table, a gap-over-time LineChart, and a Person Load table. **Non-obvious invariant:** the **Reconciled through** value (`=MIN(eur_actuals_through_month, hours_actuals_through_month)`) is a **local helper cell on Trends** at `$B$<recon_section_row+4>` — it is **not a named range**; per-WP Gap SUMs and the gap chart both read it via absolute same-sheet ref. CF thresholds: Gap % red >|10%| / amber >|5%|; Variance to-date % red >|20%| / amber >|10%|. See `trends.py` for chart wiring, helper-row block layout, and formula bodies.

### Sheet Role Contract

The authoring UX assigns each sheet a single role — derived views must not write to input surfaces and vice versa.

| Sheet | Role |
|-------|------|
| Config (PERSON-WP ASSIGNMENTS, EXTERNAL COMMITMENTS, WP tables, roster, grade rates, reporting periods) | **Primary planning surface** — all assignment/roster edits land here |
| Capacity monthly cells | Auto-filled by distribution curves on rebuild; users can override for fine-tuning (overrides lost on next rebuild) |
| Plan monthly cells | Formula-driven (`=SUM(Capacity!...)`) — **do not edit directly** |
| Timelog Plan merged grid | Literal WP codes — solver-written on init; rebuild behaviour gated by integer `timelog_rebuild_from_month` (Config!$B$9): months `< N` frozen verbatim, months `≥ N` re-solved against ceilings + pools narrowed by pre-N consumption |
| Actuals monthly cells | Input surface for EUR spend (accountant figures) |
| Actuals-Hours monthly cells | Input surface for hour timelog (primary hours surface) |
| Plan-Forecast | Read-only — every cell is a formula aggregating Actuals-Hours leaves |
| Summary, Trends | Read-only; Trends `current_month` (B1) is the only user input |

### Build Order and Tab Order

Build order (enforced by `test_build_order_is_stable`) differs from user-facing tab order — some builders are ordered by dependency rather than display position.

**Build order:** Config -> Capacity -> Plan -> patch Plan (fill FTE cells with Capacity SUMs) -> Trends stub (creates `current_month` named range before Actuals reads it) -> **Timelog Plan** (consumes config_layout + `current_month` + `team_priorities` + `roster_lookup`) -> Actuals-Hours -> Plan-Forecast -> Actuals -> Summary -> Trends.

**Tab order** (set after build via `wb.move_sheet`): **Config, Timelog Plan, Actuals-Hours**, Capacity, Plan, Actuals, Plan-Forecast, Trends, Summary. (Positions 1-3 pulled forward 2026-04-19 so the frequently-edited tabs sit one reach from Config.)

Key build-order rationale: Timelog Plan is independent of the forecast chain — it only needs the Config layout and the `current_month` stub, so it slots right after Trends-stub; Actuals-Hours must precede Plan-Forecast (forecast-EUR leaves exist before they are SUMmed); Plan-Forecast must precede Actuals (forecast monitoring block green-cross-refs Plan-Forecast); Summary must follow Plan-Forecast (pulls Forecast EUR/year); Trends is populated last so every upstream layout dict (including `helper_row_forecast`) is known.

---

## Reader and Rebuild

`reader.py` extracts structured data from an existing workbook so `rebuild_workbook()` can regenerate all sheets while preserving user-entered values. This is the most fragile module.

### Formula Chain Resolution

The reader must follow cross-sheet formula references to resolve actual values:

```
Actuals cell =Plan!A6 -> Plan!A6 =Config!A8 -> Config!A8 = "WP1"
```

`_resolve_plan_ref()` handles this, following up to 2 hops.

### Extraction Functions

| Function | Extracts | Key technique |
|----------|----------|---------------|
| `read_config_sheet(ws)` | Project params (blue-input scalars + Vacation Month; derived `num_vacation_months` + `version_stamp` via named ranges) and WP table. Through-month values are read from the Actuals / Actuals-Hours sheets, not Config. | Direct cell reads + named-range map |
| `read_actuals_override(ws)` | Optional integer blue-input override on the Row 1 through-month control strip (shared signature — caller passes either the Actuals or Actuals-Hours worksheet); returns `None` when blank/non-numeric so the effective cell falls through to the auto-derived scan | Typed cell read |
| `read_team_from_config(ws)` | Team roster (name + max_fte) | Scans from "TEAM ROSTER" header |
| `read_capacity_sheet(ws, ...)` | External commitments, per-person monthly FTE | Person header detection by bold + fill `2F5496` |
| `read_actuals_values(ws, ...)` | EUR per WP per month | Formula chain resolution |
| `read_timelog_plan(wb, months)` | Merged-grid literal WP codes per (person, month) + `rebuild_from_month` integer from Config!$B$9 (named range `timelog_rebuild_from_month`, defaults `1` / solve-all when missing or unparseable). Returns `{"grid": dict, "rebuild_from_month": int}`. | Col-A person-row scan + `TIMELOG_MONTH_COL_START` offset; named-range lookup for the freeze-point |

### What Rebuild Preserves

All sheets are regenerated from scratch. Only these user-entered values survive:

| Data | Source in old workbook |
|------|----------------------|
| Project params, WP definitions, team roster, grade rates, reporting periods | Config sheet sections |
| Person-WP assignments, external commitments | Config sheet PERSON-WP ASSIGNMENTS (B2) + EXTERNAL COMMITMENTS (B3) sections |
| Monthly EUR values per WP | Actuals data cells |
| Hour timelog per person × WP × month | Actuals-Hours data cells |
| Timelog Plan merged-grid cells (WP codes) | Timelog Plan merged grid; preserved verbatim for months `< timelog_rebuild_from_month`; months `≥ N` are re-solved (any preserved cells there are discarded). Default `N = 1` re-solves everything. |

Everything else (formulas, styling, CF, charts, Plan/Summary/Trends) is overwritten.

### Preserved Data Structures

```python
# Capacity FTE: (person_name, wp_id) -> [monthly values]
preserved_capacity: dict[tuple[str, str], list[float | None]]

# Actuals EUR: wp_id -> [monthly EUR values]
preserved_actuals: dict[str, list[float | None]]
```

### Orphan Detection

- Person removed from team: Capacity entries dropped (warning)
- WP removed from config: Capacity + Actuals entries dropped (warning)
- External commitment references missing team member: hard error

All orphaned data is printed to stderr at rebuild time. Hard errors abort rebuild; warnings do not.

### Row-Span Resolution (named-range-primary)

Reader functions for Config tables (`read_team_from_config`,
`read_assignments_from_config`, `read_externals_from_config`) are **named-range
primary**: they resolve `roster_ftes` / `assign_persons` / `ext_persons` via
`wb.defined_names` as the authoritative row span, then iterate within and
tolerate internal blank rows. Excel auto-extends these named ranges when the
user inserts rows inside them, so the reader's bounds stay in sync with live
data after hand-edits. Only when a named range is missing (legacy workbooks)
does the reader fall back to sentinel-based scanning that stops at the first
blank. See `_iter_named_range_rows` in `reader.py` and shape-drift tests in
`tests/test_rebuild_shape_drift.py` / `tests/test_rebuild_shapes.py`.

### Known Fragility

1. **Person header detection** relies on specific font/fill styling. Manual restyling breaks detection.
2. **Formula chain resolution** assumes patterns like `=Config!A8`. Custom formulas are not followed.
3. **Sentinel-scan sections** (WP BUDGET, RATE BY GRADE, reporting periods) still stop at the first blank in column A. Use filled rows when hand-editing these, or add named ranges to extend the same treatment as roster/assignments/externals.

---

## Solvers

Pure algorithmic modules — no openpyxl, no workbook I/O. Called by a builder, return plain-Python data structures the builder writes into cells. Keep solvers strictly deterministic so tests can pin behaviour.

### timelog_solver (`budget_planner/timelog_solver.py`)

Two-pass greedy allocator driving the Timelog Plan merged grid on init (and on rebuild for months `≥ timelog_rebuild_from_month`, against ceilings + pools narrowed by frozen pre-N consumption). See `docs/plans/archived/2026-04-18-timelog-plan-design.md` §4 for the algorithm spec; `docs/plans/archived/2026-04-19-timelog-plan-v2-design.md` §4 for the merged-grid rewrite; `docs/plans/archived/2026-04-23-solver-freeze-point-design.md` for the freeze-point semantics + `Person.frozen_months` solver guard.

- `greedy_pass(persons, wps, person_wp_windows, project_months)` — fills WPs in earliest-close order; within a WP, priority-1 persons before priority-2 before priority-3; continue-a-run preferred over starting a new run. Check-after-placing allows exactly one month of overshoot (the "buffer for auditor haircut" requirement).
- `switch_reduction_pass(assigned, ...)` — adjacent-pair local search: swaps `(p1, m, m+1) ↔ (p2, m, m+1)` whenever the swap is legal and strictly reduces the total switch count. **Known shape limitation:** only pair-wise adjacent swaps are considered. Rearrangements that require a 3+ month block swap or a same-month-window 3-person rotation are not found. In practice the greedy pass leaves few such patterns, and the algorithm trades optimality for determinism + bounded runtime.
- `solve(...)` — composed entry point: `greedy_pass` then `switch_reduction_pass`. Deterministic → reproducible tests in `tests/test_timelog_solver.py`.

No other solvers exist today. If a future builder needs algorithmic assignment (e.g. automatic Capacity distribution beyond the current curve presets), it belongs here as a sibling module, not inside the builder.

---

## Layout and Styling

> Layout constants, row/column positions, validation-column offsets, freeze panes, and alert thresholds live in `budget_planner/layout.py`. Not duplicated here to avoid drift. Styling primitives (fonts, fills, `BLACK_FORMULA` / `BLUE_INPUT`, `VACATION_FILL`) live in `budget_planner/styles.py`.

### Styling conventions

- **`style_cell()` auto-unlock.** When font color is blue (`0000FF`), `style_cell()` automatically sets `Protection(locked=False)` so blue inputs stay editable when sheet protection is re-enabled.
- **Column grouping.** `apply_year_groups()` creates collapsible column groups per calendar year. The group containing `current_month` is expanded; others collapse.
- **Tab color is a semantic role marker** — don't change without updating the mapping:

| Sheet | Tab color | Why |
|-------|-----------|-----|
| Config | Blue (4472C4) | Single source of truth for all parameters |
| Capacity | Purple | Primary planning surface (derived from Config assignments) |
| Timelog Plan | Orange (FF8C00) | Prescriptive 1.0-FTE WP allocation — distinct from Plan/Capacity |
| Actuals | Blue (0000FF) | EUR input from accountant |
| Actuals-Hours | Red (C00000) | Hours input — primary hours surface |
| Plan-Forecast | Green (00B050) | Read-only aggregation (green = formula-derived) |
| Summary | Green (008000) | Read-only aggregation |
| Trends | Gold (FFD700) | PM dashboard |

Row-order contracts and the inactive-month rule are cross-sheet invariants — see `docs/design-invariants.md` § Row-order contracts / § Inactive-month rule.

---

## Conditional Formatting Rules

### Config Sheet
- WP start/end columns: **yellow** when start month > end month

### Plan Sheet
- Planned variance column: **yellow** when `ABS(value) > 0.01`
- Actual remaining column: **red** when `value < 0`

### Team Sheet
- WP share SUM row: **yellow** when `ABS(SUM - 1.0) > 0.001`

### Capacity Sheet
- TOTAL row per person: **amber** when `TOTAL > 0.8 * Max FTE`, **red** when `TOTAL > Max FTE`
- REMAINING row per person: **red** when `REMAINING < 0`
- EU TOTAL per calendar year: **amber** when annual sum approaches `eu_annual_ceiling_pm`, **red** when exceeds (215-day ceiling)
- Raw-monthly Variance-FTE band: **red** when `|variance| > CAPACITY_VARIANCE_FTE_RED`
- Raw-monthly Actual-TOTAL utilisation vs Config Max FTE: amber/red same thresholds as planned TOTAL
- To-date block `Variance PM` (TOTAL + EU TOTAL only): **red** when `< 0`. Per `docs/plans/archived/2026-04-14-t5-simplification-design.md` the block was reduced to 2 cols — the prior per-WP Variance-% and to-date util/ceil CF rules were dropped as duplicates of planned-side signals.

### Actuals Sheet
- Negative variance: **red**
- % Ceiling: **amber** > 85%, **red** > 100%
- Breach In: **amber** < 12 months, **red** < 6 months
- **Reconciliation grey (T8):** Actual-EUR monthly cells past `eur_actuals_through_month` are greyed (not yet finance-closed). Mirror rule on Actuals-Hours forecast band greys the *opposite* side (months ≤ `eur_actuals_through_month`) so users see at a glance which side is authoritative per month.
- Parallel forecast-monitoring block: same amber/red thresholds as actual-EUR block on Cum Forecast Variance / % Ceiling / Breach In.

### Actuals-Hours Sheet
- Annual EU hours ceiling: **red** when annual hour sum `> eu_annual_ceiling_pm * hours_per_fte_month`, **amber** within one `hours_per_fte_month` of the ceiling

### Plan-Forecast Sheet
- Variance (Ceiling - Sum): **red** when negative (forecast exceeds ceiling)

### Summary Sheet
- Remaining EUR column: **red** when `value < 0`
- Variance EUR (per year + project total): **red** when negative
- Forecast EUR: **red** when `Forecast < 0.8 * Plan EUR` AND year is fully covered (`hours_actuals_through_month >= year_end_pm`) — gated so partial-coverage years don't trigger false-alarm underspend alerts; those are instead shaded grey (see next point)
- Per-year partial-coverage grey: Act € shaded by `eur_actuals_through_month < year_end_pm`; Fcst € by `hours_actuals_through_month < year_end_pm`; Var € by the stricter MIN. PROJECT TOTAL columns have no grey CF (span all years).
- Budget Overview Remaining: **amber** when < 10% of ceiling, **red** when negative

---

## Cross-Sheet Reference Map

| Target Cell | References | Via |
|-------------|-----------|-----|
| **Config -> Plan** | | |
| `project_start` | Config!$B$2 | Named range |
| `fte_rate` | Config!$B$4 | Named range |
| `timelog_rebuild_from_month` | Config!$B$9 | Named range |
| `version_stamp` | Config!$B$10 | Named range |
| **Plan -> Config** | | |
| Plan global cells (start, rate) | Config params | Green formula |
| **Plan -> Capacity** | | |
| Plan Planned FTE (monthly) | Capacity per-person WP rows | `=SUM(Capacity!{cells})` green formula |
| **Capacity -> Plan** | | |
| Capacity WP ID/Name | Plan config rows | `=Plan!A{row}` / `=Plan!B{row}` green formula |
| **Capacity -> Config** | | |
| Person header | Config roster name | `=Config!A{row}` formula |
| REMAINING | Config Max FTE | `=Config!B{row}-{TOTAL}` formula |
| **Summary -> Plan** | | |
| Summary WP ID/Name | Plan config rows | Green formula |
| Year N Plan/Act FTE | Plan monthly ranges | `=SUM(Plan!{range})` |
| Remaining EUR | Plan ceiling - Actual total | Formula |
| **Trends -> multiple** | | |
| WP start/end | Config refs | Formula |
| Cumulative/ceiling | Actuals refs | Formula |
| TOTAL rows | Capacity refs | Formula |
| Planned FTE (chart) | Plan refs | Formula |
