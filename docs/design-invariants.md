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
- **Timelog Plan** records **timesheet-compliant allocation** — prescriptive whole-month 1.0-FTE WP labels per (person, month) telling the person where to *log*. Driven by the solver (`budget_planner/timelog_solver.py`) and overridable on the Assigned grid.

A person splitting effort 0.6/0.4 across two WPs in a given month may legitimately log 1.0 to one of them that month and 1.0 to the other a different month — the timesheet rule is whole-month, the underlying effort is not. **Do not write cross-sheet formulas that reconcile Plan monthly cells against Timelog Plan monthly cells as if they must agree.** The only shared constraints are the Config inputs both surfaces consume (WP windows, PERSON-WP ASSIGNMENTS windows, Max FTE pool, Ceiling €).

See `docs/plans/archived/2026-04-18-timelog-plan-design.md` §1 for scope.

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
- **Reader/builder round-trip rule**: readers use named ranges (not coordinates); tolerate internal blanks; rebuild must preserve user-authored values on blue input cells.
