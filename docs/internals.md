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

**Config** (`config_sheet.py`): Rows 1-8 blue-input project params (Title, Start, Months, FTE Rate, EU Ceiling, Hours/FTE Month, EUR Actuals Through Month, Hours Actuals Through Month), then row 9 black-formula-styled `Version Stamp` (code-written audit cell overwritten on every init/rebuild via `_resolve_version_stamp()` in `workbook.py` — `git describe --always --dirty`, fallback `"unknown"`), then rows 10-12 derived/black-formula params (Months Elapsed Hours, Months Elapsed EUR, Reconciled Through Month), then WP table (`tbl_WorkPackages`), then Team Roster section. Creates named ranges `project_start` (B2), `fte_rate` (B4), `eur_actuals_through_month` (B7), `hours_actuals_through_month` (B8), `version_stamp` (B9), `months_elapsed_hours` (B10), `months_elapsed_eur` (B11), `reconciled_through_month` (B12). Returns `config_layout` with `wp_first_row`, `wp_last_row`, `person_rows`.

**Plan** (`plan.py`): Two-phase build. Phase 1 creates structure with empty FTE cells. Phase 2 (`_apply_planned_formulas()`) fills cells with `=SUM(Capacity!...)` after Capacity is built. Has validation columns: Sum, Budget FTE, Variance.

**Capacity** (`capacity.py`): Most complex builder. Creates per-person groups: WP rows (blue input, static values from distribution curves), external project rows, TOTAL (SUM), REMAINING (Max FTE - TOTAL), EU TOTAL (programme WP rows + EU-flagged externals — tracks total EU commitment against the 215-day annual ceiling, distinct from TOTAL). Per-person WP rows are **only the WPs assigned to that person** in Config PERSON-WP ASSIGNMENTS — not every WP. Distribution curve per WP (flat/front-loaded/bell/back-loaded) comes from Config WP SCHEDULE.Type. Right-side raw-monthly band (outline_level=1, default collapsed): Actual-FTE (green cross-ref to Actuals-Hours hour / hours_per_fte_month) + Variance-FTE (Planned - Actual). Always-visible 2-col **TO-DATE** block beyond that: `Actual PM` (WP rows: green ref to Actuals-Hours `cum_hours_col` / hours_per_fte_month; TOTAL/EU TOTAL: SUM of WP Actual cells) + `Variance PM` (only populated on TOTAL + EU TOTAL: windowed plan SUM on same-row planned cells − Actual; red CF < 0). Returns `capacity_layout` with `person_groups` + `todate_col_actual` / `todate_col_var`.

**Timelog Plan** (`timelog_plan.py`): Prescriptive whole-month 1.0-FTE WP allocation grid — separate allocation surface from Plan/Capacity (see `docs/design-invariants.md` § Plan vs Timelog Plan). Twin grids: a **Proposed** grid (left, locked, pastel-filled by CF per WP) written by the solver (`budget_planner/timelog_solver.py`) and an **Assigned** grid (right) that starts as `=<Proposed cell>` formulas and accepts blue-input overrides. Row-2 traffic lights + per-person summary block (Assigned count, Pool, Remaining, Over-pool?, Switches, First log, Last log) + per-WP summary block (Ceiling €, Assigned €, Overshoot €, Undershoot €, Assigned months, Window months remaining, Any out-of-window?). Per-WP `Assigned €` uses a SUMPRODUCT `(assigned_grid="WPk")*person_rate_col` pattern — verified in LO; if Excel returns #VALUE! in a real-file check, swap to per-row helper-column pattern. Reader: `read_timelog_plan` extracts `assigned_overrides` (cells the user typed literal values into, not `=Proposed!...` formulas) so rebuild preserves them.

**Actuals** (`actuals.py`): Similar to Plan but with EUR input cells. Row 1 is an "As of" header pair (label + `=EDATE(project_start, eur_actuals_through_month-1)` green date for the Actual block; same pattern for the Forecast block at the forecast-column offset). Section header on row 2. Extra validation columns: Cum. EUR (SUMPRODUCT up to `current_month`), % Ceiling, Breach In — formula `(ceiling − cum EUR) / (cum EUR / months_elapsed_eur)` (cumulative-slope from project start through `eur_actuals_through_month`; not a per-month slope). Parallel forecast-monitoring block (4 cols, outline_level=1, collapsible, separated from actual-EUR block by a 1-col gutter): Cum Forecast EUR, Forecast Variance, Forecast % Ceiling, Forecast Breach In (denominator `months_elapsed_hours`) — green cross-refs into Plan-Forecast WP rows. CF rules mirror the actual-EUR block. Return dict key `header_row` is `1` (the "As of" row); section header is on row 2.

**Actuals-Hours** (`actuals_hours.py`): Per-person monthly hour timelog — primary hours input surface. Per-person block order: header, WP rows, external rows, **two summary rows** ("Cumulative hours to date" + "Actual PM to date" — col A blank so the reader skips them; formulas reference `cum_hours_col`; outline_level=1 so they collapse with the block), TOTAL, EU TOTAL. The TOTAL per-month SUM spans only the data rows (WP + external), not the summary rows. Annual ceiling is enforced by CF against `eu_annual_ceiling_pm * hours_per_fte_month`. Person header A-column name **must be a literal string**, not a cross-ref formula — it is the VLOOKUP key for the forecast-EUR leaves and must evaluate under `data_only=True`. Sibling forecast-EUR band (outline_level=1, 1-col gutter right of Sum): each leaf `=IF(h="","",ROUND(h/hours_per_fte_month*VLOOKUP(person,roster_lookup,4,FALSE),0))`. ROUND wraps each leaf to whole EUR so hundreds-of-leaves SUMs on Plan-Forecast don't drift. Forecast-EUR leaves here are the **sole authoritative source** for Plan-Forecast and Actuals forecast monitoring. One gutter right of the forecast band, an always-visible helper col **`Cum hours (to date)`** carries a per-row windowed `SUM+INDEX` cumulation through `hours_actuals_through_month` — Capacity's to-date Actual PM reads this single cell per WP row (see `cum_hours_col` in the returned layout). Two grey CF rules on the forecast band: `month <= eur_actuals_through_month` (finance-closed) and `month > hours_actuals_through_month` (no hours yet).

**Plan-Forecast** (`plan_forecast.py`): WP x month forecast-EUR grid — pure aggregation, read-only (every cell is a formula). Shape mirrors Plan. Each monthly cell `=SUM(<Actuals-Hours forecast-EUR leaves for this WP x month across all assignments>)`. No user inputs.

**Summary** (`summary.py`): Calendar-year groups with Plan FTE, Plan EUR, Actual EUR, Forecast EUR, Variance EUR sub-columns (5 per year) + PROJECT TOTAL group. Budget Overview section: Personnel Actual, Total Claimable (+25% indirect), Personnel Ceiling, Remaining, % Consumed. Reporting Period Summary section (grouped by Config reporting periods) with Form C preparation rows (Personnel EUR, Indirect 25%, Total Eligible, Cum. Eligible, CFS Required?).

**Trends** (`trends.py`): Dashboard with WP Burn Rate & Exhaustion table; cumulative EUR LineChart (Planned / Actual / Cumulative Forecast / Ceiling / RP End) where Actual and Forecast series emit `NA()` past `eur_actuals_through_month` / `hours_actuals_through_month` respectively so the lines clip cleanly; a **Reconciliation: Forecast vs Actual EUR** section (between the cumulative chart and Person Load) with three health cells (Actuals EUR through / Hours through / Reconciled through) and a per-WP summary table (WP, Cum Forecast EUR, Cum Actual EUR, Gap EUR, Gap %) with TOTAL row, all cumulative-through-`reconciled_through_month`, Gap % CF red >|10%| and amber >|5%|; a second LineChart (gap-over-time, anchored right below the reconciliation TOTAL) — single purple `cum Actual − cum Forecast` series clipped via `NA()` past `reconciled_through_month`, dashed grey zero reference line; a 4-column **Person Load** table (Person, Avg FTE to date from Actuals-Hours cumulative / `months_elapsed_hours`, Plan Avg FTE to date, Variance to-date %; CF red >|20%| / amber >|10%|). Hidden helper-row block below Person Load grew from 6 to 8 rows (added `helper_row_gap`, `helper_row_zero`) and feeds both charts' data series.

### Sheet Role Contract

The authoring UX assigns each sheet a single role — derived views must not write to input surfaces and vice versa.

| Sheet | Role |
|-------|------|
| Config (PERSON-WP ASSIGNMENTS, EXTERNAL COMMITMENTS, WP tables, roster, grade rates, reporting periods) | **Primary planning surface** — all assignment/roster edits land here |
| Capacity monthly cells | Auto-filled by distribution curves on rebuild; users can override for fine-tuning (overrides lost on next rebuild) |
| Plan monthly cells | Formula-driven (`=SUM(Capacity!...)`) — **do not edit directly** |
| Timelog Plan Proposed grid | Solver-written, locked — regenerated every rebuild |
| Timelog Plan Assigned grid | Starts as `=Proposed`; user overrides become blue-input literals, preserved across rebuild |
| Actuals monthly cells | Input surface for EUR spend (accountant figures) |
| Actuals-Hours monthly cells | Input surface for hour timelog (primary hours surface) |
| Plan-Forecast | Read-only — every cell is a formula aggregating Actuals-Hours leaves |
| Summary, Trends | Read-only; Trends `current_month` (B1) is the only user input |

### Build Order and Tab Order

Build order (enforced by `test_build_order_is_stable`) differs from user-facing tab order — some builders are ordered by dependency rather than display position.

**Build order:** Config -> Capacity -> Plan -> patch Plan (fill FTE cells with Capacity SUMs) -> Trends stub (creates `current_month` named range before Actuals reads it) -> **Timelog Plan** (consumes config_layout + `current_month` + `team_priorities` + `roster_lookup`) -> Actuals-Hours -> Plan-Forecast -> Actuals -> Summary -> Trends.

**Tab order** (set after build via `wb.move_sheet`): Config, Capacity, Plan, **Timelog Plan**, Actuals, Actuals-Hours, Plan-Forecast, Trends, Summary.

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
| `read_config_sheet(ws)` | Project params (rows 1-8 blue inputs + row 9 version stamp + rows 10-12 derived), WP table (rows 13+) | Direct cell reads + named-range map |
| `read_team_from_config(ws)` | Team roster (name + max_fte) | Scans from "TEAM ROSTER" header |
| `read_capacity_sheet(ws, ...)` | External commitments, per-person monthly FTE | Person header detection by bold + fill `2F5496` |
| `read_actuals_values(ws, ...)` | EUR per WP per month | Formula chain resolution |

### What Rebuild Preserves

All sheets are regenerated from scratch. Only these user-entered values survive:

| Data | Source in old workbook |
|------|----------------------|
| Project params, WP definitions, team roster, grade rates, reporting periods | Config sheet sections |
| Person-WP assignments, external commitments | Config sheet PERSON-WP ASSIGNMENTS (B2) + EXTERNAL COMMITMENTS (B3) sections |
| Monthly EUR values per WP | Actuals data cells |

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
3. **Sentinel-scan sections** (WP BUDGET, WP SCHEDULE, RATE BY GRADE, reporting periods) still stop at the first blank in column A. Use filled rows when hand-editing these, or add named ranges to extend the same treatment as roster/assignments/externals.

---

## Solvers

Pure algorithmic modules — no openpyxl, no workbook I/O. Called by a builder, return plain-Python data structures the builder writes into cells. Keep solvers strictly deterministic so tests can pin behaviour.

### timelog_solver (`budget_planner/timelog_solver.py`)

Two-pass greedy allocator driving the Timelog Plan Proposed grid. See `docs/plans/archived/2026-04-18-timelog-plan-design.md` §4 for the algorithm spec.

- `greedy_pass(persons, wps, person_wp_windows, project_months)` — fills WPs in earliest-close order; within a WP, priority-1 persons before priority-2 before priority-3; continue-a-run preferred over starting a new run. Check-after-placing allows exactly one month of overshoot (the "buffer for auditor haircut" requirement).
- `switch_reduction_pass(assigned, ...)` — adjacent-pair local search: swaps `(p1, m, m+1) ↔ (p2, m, m+1)` whenever the swap is legal and strictly reduces the total switch count. **Known shape limitation:** only pair-wise adjacent swaps are considered. Rearrangements that require a 3+ month block swap or a same-month-window 3-person rotation are not found. In practice the greedy pass leaves few such patterns, and the algorithm trades optimality for determinism + bounded runtime.
- `solve(...)` — composed entry point: `greedy_pass` then `switch_reduction_pass`. Deterministic → reproducible tests in `tests/test_timelog_solver.py`.

No other solvers exist today. If a future builder needs algorithmic assignment (e.g. automatic Capacity distribution beyond the current curve presets), it belongs here as a sibling module, not inside the builder.

---

## Layout and Styling

`layout.py` and `styles.py` centralize constants that every builder imports.

### Month Column Positioning

All sheets with monthly data share the same column structure:

```
Col A (1) = WP ID / Person name
Col B (2) = Label
Col C (3) = Month 1              (MONTH_COL_START = 3)
...
Col N     = Month N              (month_col_end(N))
Col N+1   = Sum                  (val_columns(N)["sum"])
Col N+2   = Budget               (val_columns(N)["budget"])
Col N+3   = Variance             (val_columns(N)["diff"])
Col N+4   = Cum. EUR             (Actuals only)
Col N+5   = % Ceiling            (Actuals only)
Col N+6   = Breach In            (Actuals only)
```

Validation columns are offsets from the last month column, so changing project duration auto-repositions them.

### Alert Thresholds (layout.py)

| Constant | Value | Used by |
|----------|-------|---------|
| `CAPACITY_AMBER_LOAD` | 0.80 | Capacity + Trends person load |
| `ACTUALS_AMBER_PCT_CEIL` | 0.85 | Actuals % ceiling |
| `ACTUALS_AMBER_BREACH_MONTHS` | 12 | Actuals breach column |
| `ACTUALS_RED_BREACH_MONTHS` | 6 | Actuals breach column |
| `TREND_OK_BAND` | 0.05 | Trends burn delta |
| `TREND_OVER_THRESHOLD` | 0.15 | Trends burn delta |
| `TREND_SLOW_THRESHOLD` | -0.10 | Trends burn delta |

### style_cell() Auto-Unlock

When font color is blue (`0000FF`), `style_cell()` automatically sets `Protection(locked=False)`. When sheet protection is re-enabled, all blue input cells become editable with no extra code.

### Column Grouping

`apply_year_groups()` creates collapsible column groups by calendar year. The group containing `current_month` is expanded; all others are collapsed.

### Tab Colors

Each sheet's tab color is a semantic role marker — don't change without updating this table.

| Sheet | Tab color | Why |
|-------|-----------|-----|
| Config | Blue (4472C4) | Single source of truth for all parameters |
| Capacity | Purple | Primary planning surface (derived from Config assignments) |
| Timelog Plan | Orange (FF8C00) | Prescriptive 1.0-FTE WP allocation — distinct surface from Plan/Capacity |
| Actuals | Blue (0000FF) | EUR input from accountant |
| Actuals-Hours | Red (C00000) | Hours input — distinct to flag as the primary hours surface |
| Plan-Forecast | Green (00B050) | Read-only aggregation (green = formula-derived) |
| Summary | Green (008000) | Read-only aggregation |
| Trends | Gold (FFD700) | PM dashboard |

### Freeze Panes

Every sheet with monthly data freezes at column C (row varies) so columns A (ID) and B (Name/Label) stay visible during horizontal scroll. Row varies by sheet — use the builder's live coordinates, not a hardcoded row number:

| Sheet | Freeze |
|-------|--------|
| Config | `A{wp_section_row}` (freezes project-params block above WP BUDGET) |
| Plan | `C{plan_first_wp}` |
| Actuals | `C4` |
| Actuals-Hours | `C{first_person_header}` |
| Plan-Forecast | `C{pf_first_wp_row}` |
| Summary | `C{sum_first_wp}` |

### Row-Order Contracts

Cross-sheet formulas (Plan's `=SUM(Capacity!...)`, Trends person-load refs, Summary aggregations) index by row position, so row order is a contract not an accident.

- **Capacity person order matches Config TEAM ROSTER order.** Adding a roster row shifts every downstream Capacity person group.
- **Plan WP row order matches Config WP BUDGET order.** Same for Actuals, Plan-Forecast, Summary.
- **Capacity per-person WP rows are ordered by Config PERSON-WP ASSIGNMENTS** for that person (intersection filter — not all WPs).
- **Actuals-Hours person + WP row order mirrors Capacity** — the T9 Actual-FTE band on Capacity sibling-refs Actuals-Hours by matched row position.

### Inactive-Month Rule

On every sheet with monthly data (Capacity, Plan, Actuals, Actuals-Hours, Plan-Forecast), months outside the intersection of the assignment window `[Start, End]` and WP SCHEDULE active months are **greyed and empty**. Distribution curves must never leak FTE into gap months. WPs with multiple SCHEDULE entries (gapped active periods) grey out the gap months even when inside the assignment window.

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
- To-date block `Variance PM` (TOTAL + EU TOTAL only): **red** when `< 0`. Per `docs/plans/2026-04-14-t5-simplification-design.md` the block was reduced to 2 cols — the prior per-WP Variance-% and to-date util/ceil CF rules were dropped as duplicates of planned-side signals.

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
