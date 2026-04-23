# Design Invariants

Cross-sheet invariants, trust boundaries, and derived-metric rules that must hold across code, tests, and docs. When a new structural invariant lands, record it here — future agents need it in loaded context, not buried in a sheet builder.

## Two-cadence actuals (hours vs EUR)

The workbook records actuals on two separate cadences. Any derived view that reads actual data must pick the correct trust boundary — mixing them produces wrong answers that look right.

### Hours cadence

Per-person × per-WP × per-month cells on `Actuals-Hours`.

- **Granularity:** monthly, per person per WP.
- **Entry convention:** skip-and-cumulate is valid. A user may leave intermediate months blank and enter the cumulative hours-to-date in a later month. This means a single-month cell is **NOT** authoritative for that month — only the running cumulative total is.
- **Trust boundary:** `hours_actuals_through_month` — the named range now points at `'Actuals-Hours'!$D$1` (black-formula effective cell; sheet name single-quoted because of the hyphen). D1 is `=IF(ISNUMBER(B1),B1,C1)` where B1 is an optional blue-input user override and C1 is an auto-derived scan of the programme-WP hour data (largest 1-based month-index with any numeric value, external/subtotal rows excluded). Months `1..hours_actuals_through_month` are closed for hours; later months are forecast-only territory.

### EUR cadence

WP-level totals on the `Actuals` sheet.

- **Granularity:** monthly, per WP (no per-person EUR on `Actuals`).
- **Entry convention:** driven by accountant/finance close, which runs with a 2–3 month lag behind the hours timelog.
- **Trust boundary:** `eur_actuals_through_month` — the named range now points at `Actuals!$D$1` (black-formula effective cell). D1 is `=IF(ISNUMBER(B1),B1,C1)` where B1 is an optional blue-input user override and C1 is an auto-derived scan of the WP EUR data (largest 1-based month-index with any numeric value, TOTAL row excluded). Months past this index are not yet reconciled.

### Derived boundary

- "Reconciled through month" = `MIN(eur_actuals_through_month, hours_actuals_through_month)`. It is no longer a named range — Trends holds the single live helper cell (`$B$<recon_section_row+4>`, black formula) and consumers on the same sheet read it via absolute ref; cross-sheet consumers inline the `MIN(...)` expression.
- This is the authoritative boundary for any **head-to-head comparison** between Actual EUR and Forecast EUR (hours-derived). Past this month at least one side is stale or missing.

### Design principle (enforce on every new actual-side readout)

No derived view reads a single-month cell as authoritative for that month. Every actual-side readout is **either**:

- **(a)** a cumulative-through-a-through-month total, or
- **(b)** a to-date **rate** (cumulative / `MAX(1, <through_month>)`, inlined — the former `months_elapsed_*` derived names were removed in the 2026-04-19 Config cleanup).

Raw monthly actual cells surface only in collapsed / clearly-labelled bands ("reliable only when hours were entered per-month"), never as the input to a derived metric.

### Rollout status (hours-timelog v2)

Stage 1 wired the Config inputs + named ranges only; no consumer code read them yet — so if you grep for `hours_actuals_through_month` in an early-rollout checkout and find only Config + docs, that was expected mid-rollout, not drift. Later (2026-04-19 config-through-month cleanup) the two through-month named ranges moved off Config onto `Actuals!$D$1` / `'Actuals-Hours'!$D$1`; downstream consumers were not touched — they still resolve via the named range.

Stage 2 added a to-date column block on Capacity, switched Breach-In to a cumulative-slope formula, drove per-column partial-year shading on Summary by the stricter of the two through-months, and added cosmetic grey shading past the through-month on `Actuals-Hours` + `Plan-Forecast`.

Stage 3 added a Reconciliation section on Trends: three health cells (EUR through / Hours through / Reconciled through — since 2026-04-19 the Reconciled-through cell is a local Trends helper, not a named range), a per-WP summary table with cumulative Forecast vs Actual EUR + gap %, and a gap-over-time LineChart clipped at the Reconciled-through helper.

Stage 4 (T14–T15) will harden rebuild preservation + the shape fingerprint.

## Plan vs Timelog Plan — two allocation surfaces

The workbook hosts two distinct allocation surfaces, both editable, and they may legitimately diverge:

- **Plan / Capacity** records **actual effort** — fractional monthly FTE per (person, WP) reflecting what people *do*. Driven by Config PERSON-WP ASSIGNMENTS + distribution curves.
- **Timelog Plan** records **timesheet-compliant allocation** — prescriptive whole-month 1.0-FTE WP labels per (person, month) telling the person where to *log*. Single merged editable grid (v2, 2026-04-19) holding literal WP codes; no formulas per cell. Initial content comes from the solver (`budget_planner/timelog_solver.py`); rebuild behaviour is governed by the `timelog_rebuild_from_month` freeze-point (see below).

A person splitting effort 0.6/0.4 across two WPs in a given month may legitimately log 1.0 to one of them that month and 1.0 to the other a different month — the timesheet rule is whole-month, the underlying effort is not. **Do not write cross-sheet formulas that reconcile Plan monthly cells against Timelog Plan monthly cells as if they must agree.** The only shared constraints are the Config inputs both surfaces consume (WP windows, PERSON-WP ASSIGNMENTS windows, Max FTE pool, Ceiling €).

See `docs/plans/archived/2026-04-18-timelog-plan-design.md` §1 for v1 scope and `docs/plans/archived/2026-04-19-timelog-plan-v2-design.md` for the merged-grid + toggle redesign.

### Timelog Plan merged grid = literal values; `timelog_rebuild_from_month` governs re-solve

The Timelog Plan merged grid holds **literal WP codes** (or empty cells), never formulas — user edits and solver output share one surface. The integer cell Config!$B$9 (named range `timelog_rebuild_from_month`) controls rebuild:

- Let `N = max(1, int(rebuild_from_month))`. Months `< N` are **frozen**: the solver does not touch them, and any literal WP code preserved from the prior workbook is written back verbatim.
- Months `≥ N` are **solved**: the solver re-runs for that span against per-WP `ceiling_eur` and per-Person `pool` already narrowed by the frozen span's consumption (builder-side subtraction — solver purity preserved). Any literal WP code preserved in months `≥ N` is discarded.
- Endpoints collapse the old binary toggle: `N ≤ 1` (or blank → defaults to `1`) ≡ "solve everything" (old `resolve_flag=TRUE`); `N > project_months` ≡ "preserve everything" (old `resolve_flag=FALSE`, remaining solver span empty).
- Overshoot-as-state: if the frozen span over-consumes a WP's `ceiling_eur`, the solver does not raise. The per-WP Buffer +i cell renders the negative residual against the **original** (pre-narrowed) Budget — matches existing out-of-window overshoot semantics.

No backwards-compatibility shim: the old `timelog_resolve_flag` named range is unregistered on first rebuild (one-shot migration); the derelict toggle cell at Timelog Plan row 2 col 10 is left untouched. Readers encountering a pre-v2 workbook (no `timelog_rebuild_from_month` name) default to `1` / solve-all.

### Banned: `MATCH(1, arr, 0)` / `LOOKUP(2, 1/arr, …)` — use `SUMPRODUCT(MIN/MAX(…))`

Formulas that pick the first/last non-empty index via `MATCH(1, (arr<>"")*1, 0)` or `LOOKUP(2, 1/(arr<>""), headers)` depend on CSE array coercion that **openpyxl-written workbooks do not carry**, so LibreOffice recalc returns `""` or the wrong extreme silently (v1 Timelog Plan's First log / Last log shipped with this bug — diagnosed and fixed in v2, commit 7d8c399). Use `SUMPRODUCT(MIN((arr<>"")*COL(...)))` / `SUMPRODUCT(MAX(...))` instead — SUMPRODUCT forces array context natively and survives the openpyxl→LO round-trip.

## Future invariants (reserved)

Placeholders for invariants that have a home elsewhere today but belong here when the next cross-sheet rule lands. Write them up when the next invariant accretes — don't flesh them out speculatively.

- **Font color convention** (Critical Rule 3 pointer): Blue = input, Black = formula, Green = cross-sheet ref.
- **Named-range contract**: who writes vs reads each named range.

  | Named range | Written by | Read by |
  |---|---|---|
  | `project_start` | `config_sheet.py` (Config!$B$2) | nearly every builder (date math, headers) |
  | `fte_rate` | `config_sheet.py` (Config!$B$4) | `plan.py` leaves (currently via direct Config ref; pending consolidation) |
  | `hours_actuals_through_month` | `actuals_hours.py` (`'Actuals-Hours'!$D$1`, black-formula effective `=IF(ISNUMBER(B1),B1,C1)`; B1 blue-input override, C1 auto-derived scan) | `capacity.py`, `actuals_hours.py` (greying), `plan_forecast.py` (greying), `summary.py` (partial-year shading), `trends.py` (reconciliation, NA-clipping). Also inlined as `MAX(1,hours_actuals_through_month)` at every prior `months_elapsed_hours` callsite. |
  | `eur_actuals_through_month` | `actuals.py` (`Actuals!$D$1`, black-formula effective `=IF(ISNUMBER(B1),B1,C1)`; B1 blue-input override, C1 auto-derived scan) | `actuals.py` (Breach-In cumulative-slope, As-of), `summary.py`, `trends.py`. Also inlined as `MAX(1,eur_actuals_through_month)` at every prior `months_elapsed_eur` callsite. |
  | `roster_lookup` | `config_sheet.py` (Team Roster cols A-D: name, Max FTE, grade, rate EUR/PM; Priority column E intentionally excluded so existing VLOOKUP indices stay stable) | `actuals_hours.py` (forecast-EUR leaves VLOOKUP col 4), `config_sheet.py` (Assignments Balance VLOOKUP col 2), `timelog_plan.py` (per-person rate helper VLOOKUP col 4) |
  | `team_priorities` | `config_sheet.py` (Team Roster Priority col E, blue input, default 2) | `timelog_plan.py` solver fill order — priority flows in via `reader.read_team_from_config` (header scan of the Priority column, not via name resolution). The named range is reserved for future Excel-formula consumers. |
  | `current_month` | `workbook.py` Stage 5 Trends stub (Trends!$B$1, **black formula** — `=MAX(1,MIN(Config!B3, ...))`; user may overwrite with a static number to freeze). Stub written before Actuals builder runs. | `actuals.py` (Cum. EUR / Breach-In cutoff), `timelog_plan.py` (per-WP `Window months remaining`) |
  | `timelog_rebuild_from_month` | `config_sheet.py` (Config!$B$9, blue-input integer, default `1` on init, `min=1 / max=999` data validation) | `reader.read_timelog_plan` returns it as `rebuild_from_month: int`; `timelog_plan.py` builder partitions `preserved_grid` into frozen (months `< N`) and narrows solver-input `ceiling_eur` + `pool` by frozen consumption (see "Timelog Plan merged grid" above) |
  | `vacation_month` | `config_sheet.py` (Config!$B$7, blue input string, default `"Jul"`) | `reader.read_config_sheet` → orchestrator + `timelog_plan.py` solver (vacation-month exclusion). Fallback to `"Jul"` when missing (one-shot migration for pre-enhancement workbooks). |
  | `num_vacation_months` | `config_sheet.py` (Config!$B$8, BLACK_FORMULA builder-written literal — Rule 3 carve-out, same pattern as `version_stamp`) | `timelog_plan.py` PER-PERSON SUMMARY Pool formula: `=ROUND(max_fte × (project_months − num_vacation_months), 0)`. Recomputed on rebuild from `vacation_month` + `project_start` + `Project Months`. |
- **Reader/builder round-trip rule**: readers use named ranges (not coordinates); tolerate internal blanks; rebuild must preserve user-authored values on blue input cells.

## Timelog Plan — auditor-side invariants

Cross-sheet invariants only. Grid/summary/totals-row implementation details (Pool formula, Monthly total € row, Yearly total/budget rows, build-time proportions) live next to the code — see the `timelog_plan.py` module docstring and the `_write_*` helper docstrings.

### Vacation cells are always locked and empty

Every TIMELOG GRID cell whose 1-indexed project month matches the calendar `vacation_month` must carry `VACATION_FILL` (pale yellow `#FFF2CC`), `Protection(locked=True)`, and an empty string value across every person row — NOT BLUE_INPUT (cell is not user-editable). The solver must never place a WP code there; on rebuild, the builder drops vacation-month entries from `frozen_placements` before the overlay step, so a preserved WP code in a vacation cell is discarded regardless of where the vacation month falls relative to `timelog_rebuild_from_month`. No CF rule targets vacation cells; the direct yellow fill is authoritative. Violating any leg of this invariant corrupts the programme auditor story (time logged during vacation is non-allowable) or breaks the solver's `(grid<>"") × rate` SUMPRODUCT inputs.

## WP schedule — single active range

Every WP has exactly one `(start, end)` pair (cols H/I of WP BUDGET). The reader hard-fails on multi-row per WP or legacy two-section Config layout with a `ValueError` carrying a migration recipe (see `budget_planner/reader.py` — "Legacy Config layout detected" / "Duplicate WP id"). Multi-segment schedules are rejected at rebuild.

## Row-order contracts (cross-sheet)

Cross-sheet formulas (Plan's `=SUM(Capacity!...)`, Trends person-load refs, Summary aggregations) index by row position, so row order is a contract not an accident.

- **Capacity person order matches Config TEAM ROSTER order.** Adding a roster row shifts every downstream Capacity person group.
- **Plan WP row order matches Config WP BUDGET order.** Same for Actuals, Plan-Forecast, Summary.
- **Capacity per-person WP rows are ordered by Config PERSON-WP ASSIGNMENTS** for that person (intersection filter — not all WPs).
- **Actuals-Hours person + WP row order mirrors Capacity** — the Actual-FTE band on Capacity sibling-refs Actuals-Hours by matched row position.

## Inactive-month rule

On every sheet with monthly data (Capacity, Plan, Actuals, Actuals-Hours, Plan-Forecast), months outside the intersection of the assignment window `[Start, End]` and the WP's single-range active months are **greyed and empty**. Distribution curves must never leak FTE into out-of-range months. Auditors reject PM logged outside a WP's active range.

## Config `Budget +i` (col F) is a derived black-formula, not a blue input

`Config!F13:F20` holds a structured-reference formula into `tbl_WorkPackages` that resolves to `FTE Budget × Rate (€/PM) -i × 1.25` (shipped 2026-04-23 as Aiko-port T1). The authoritative blue-input fields for a WP's budget are **FTE Budget (col C)** and **Rate (col D)**; col F is a display-layer derivation.

- **Rule 3 carve-out**: the cell uses `BLACK_FORMULA` font even though Budget +i is semantically the grant-agreement-stated figure — signals "do not hand-edit; edit FTE or Rate instead."
- **Reader round-trip**: `reader.py::read_config_sheet` loads Config under `data_only=False`, so `ws.cell(r, COL_WP_BUDGET_PLUS_I).value` returns the formula string; `_safe_float` falls back to the same `FTE × Rate × 1.25` product the formula computes, preserving `budget_plus_i` numerically on rebuild. Any user hand-edit to `F{n}` is silently reset by the next rebuild — this is intentional (Aiko-port Q-F).
- **Consequences for downstream sheets**: `wp_budget_with_indirect` named range still targets `Config!$F${first}:$F${last}`, so Actuals / Plan-Forecast `=INDEX(wp_budget_with_indirect, i+1)` formulas keep resolving to the same numeric after recalc.

## Timelog Plan PER-WP SUMMARY col B is `Config!F{wp_row}` (Budget +i)

Shipped 2026-04-23 as Aiko-port T2. `_write_wp_summary_row` col B xrefs `Config!$F{cfg_wp_row}` (via `COL_WP_BUDGET_PLUS_I`) — direct-read of Budget +i. Previously xrefed col E (Budget -i).

Knock-on effects recorded here so future row-15 / cross-check work stays consistent:
- `_write_yearly_rows` col C cross-checks are raw SUMs (no inline ×1.25). `yearly_total_row` col C = `SUM(C{first}:C{last})` is -i by construction (PER-WP SUMMARY col C / Assigned € stays -i). `yearly_budget_row` col C = `SUM(B{first}:B{last})` is +i (col B is now +i direct).
- Drift-detection between `yearly_total_row` col C and the per-year cells on the same row requires a ×1.25 scale (col C is -i, per-year cells are +i via ×1.25-at-leaf). See `tests/test_timelog_plan.py::test_yearly_total_col_c_crosscheck`.
- `yearly_budget_row` col C exactly matches the per-year sum within float64 noise (both +i from Config!F).
