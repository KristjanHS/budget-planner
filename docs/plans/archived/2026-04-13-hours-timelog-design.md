# Monthly hours timelog — design

**Date:** 2026-04-13
**Status:** Design (pre-implementation-plan)
**Audience:** PI / project lead; secondarily auditor

## Problem

Real EUR actuals land ~2 months after month-end (finance-close lag). Monthly hours per person per WP are available at month-end. Today the workbook has no way to use the fast signal: Plan → Actual EUR variance monitoring runs only on the slow signal, so problems compound for two months before they show in any alarm.

## Goal

Produce a layered month-end ritual driven by logged hours:

1. **Forecast EUR** from hours at month-end (shadow actuals).
2. **Early variance alarm** on WPs and people, parallel to the existing EUR-based alarms.
3. **Reallocation support** — where free capacity exists, where people are over cap.
4. **Delivery-risk readout** — WP progress against PM budget.

The existing EUR-based monitoring (Actuals, Summary, Trends) is preserved. The new layer runs in parallel, becomes authoritative for months the EUR hasn't landed yet, and acts as a cross-check for months the EUR has landed.

## Non-goals for v1

- Workload smoothing recommendations (solver work).
- Skill/grade coverage analysis (new taxonomy).
- Dedicated paste-in CSV import (manual entry into the same grid suffices).
- Personal per-team-member views.

See [Roadmap](#roadmap) for v2 / v3 / v4.

## Prerequisites (already met)

- Per-person rate is resolvable: roster holds grade per person; `GRADE_RATES` table maps grade → EUR/PM; `roster_lookup` named range already encodes this chain (`budget_planner/config_sheet.py:488,512`). Forecast-EUR leaves use the same `VLOOKUP(person, roster_lookup, …)` pattern.

## v1 — scope

### 1. Data model

**New sheet: `Actuals-Hours`** — mirrors the `Capacity` sheet 1:1.

- Same person sections, same WP rows per person, same month columns. Reuse `layout.py` constants; do not duplicate.
- Monthly cells are **hours** (blue input), not FTE. Human-natural unit at month-end.
- Inactive-month cells greyed out and empty (Critical Rule 4).
- Per-person summary rows: TOTAL hours (programme WPs), EU TOTAL hours.
- No Budget / Variance / Remaining columns — those live on derived views.

**New Config constant:** `hours_per_fte_month` — default **143** (1720h/year ÷ 12). Placed in the next free Config cell, named range `hours_per_fte_month`. Single source.

**Builder:** `budget_planner/actuals_hours.py`, stateless builder contract mirroring `capacity.py`. Share person-section geometry helpers with `capacity.py` (extract if needed; do not copy-paste).

**_META** gets one new sheet entry. Reader and rebuild logic covers the new sheet.

### 2. Forecast EUR chain

Per person-WP-month leaf cell:

```
forecast_FTE  = Actuals-Hours[person, wp, month] / hours_per_fte_month
forecast_EUR  = forecast_FTE × person_rate
person_rate   = VLOOKUP(person, roster_lookup, <grade-rate col>, FALSE)
```

**Leaf EUR lives on `Actuals-Hours`** as a **collapsible (outline-grouped) column band**, visible by default, collapsible to hide when not needed. Each person-WP-month cell has a sibling EUR cell using the formula above.

**Aggregations `SUM` from the leaves** (approach chosen over SUMPRODUCT to keep every EUR traceable to a single cell):

- **Plan-Forecast sheet (new)** — WP × month shape matching `Plan`. Per-WP-month cell sums the relevant leaf EUR cells on `Actuals-Hours`.
- **Summary** — new column per year: **Forecast EUR (hours-based)** next to Plan EUR and Actual EUR.
- **Actuals sheet** — new column block per WP (see §3) reading from Plan-Forecast.

### 3. Forecast-driven variance & alarm layer

Everything Actuals EUR produces today, Forecast EUR produces in parallel — same columns, same formulas, same CF rules.

**On `Actuals` sheet**, per WP row, duplicate the monitoring block:

| Existing (Actual EUR)   | New parallel (Forecast EUR)  |
|-------------------------|------------------------------|
| Ceiling                 | (shared)                     |
| Cumulative EUR          | Cumulative Forecast EUR      |
| Variance vs. Ceiling    | Forecast Variance vs. Ceiling|
| % Ceiling               | Forecast % Ceiling           |
| Breach In (months)      | Forecast Breach In           |

CF thresholds match exactly: red on breach, amber on ≥85%, amber on <6 months to breach. Forecast columns sit in a grouped block to the right of the Actual EUR block with a clear header band.

**Reconciliation via named range `eur_actuals_through_month`:**

- Months ≤ `eur_actuals_through_month`: Actual EUR is authoritative; Forecast is a cross-check.
- Months > `eur_actuals_through_month`: Forecast is authoritative; Actual cells blank/shaded.

A single named range drives the visual emphasis. `current_month` (already exists) keeps its role for the reporting-period boundary.

**On `Trends` sheet:**

- **LineChart** adds series: **Cumulative Forecast EUR** alongside Cumulative Actual EUR, Cumulative Planned, Ceiling. Forecast line typically runs ~2 months ahead of Actual — by design.
- **Person Load table** gains parallel columns: peak / average **actual load** per person (from hours), and a per-person **planned − actual load** variance.

**On `Summary`:**

- Forecast EUR column per year gains the same variance-vs-Plan CF as Actual EUR has today.

### 4. Capacity-side variance (plan FTE vs. actual FTE)

Separate from EUR — hours-vs-plan in FTE space, per person per month.

**On `Capacity` sheet** (or a twin view on `Actuals-Hours`):

- **Actual FTE** per monthly cell: `= Actuals-Hours[cell] / hours_per_fte_month` (green cross-sheet).
- **Variance FTE** per cell: planned − actual. Red CF if `|variance| > 0.2` FTE; threshold in `layout.py`.
- **Utilisation %** per person per month: actual TOTAL / Max FTE. Red ≥100%, amber ≥80%. Mirrors the planned-side CF.

**Per-person summary** gains parallel **actual TOTAL** and **actual EU TOTAL** rows. EU annual ceiling validation runs against **actual** hours too — an auditor-relevant early warning ahead of EUR confirmation.

### 5. Colour convention

- **Blue** = input (hours cells on `Actuals-Hours`).
- **Black** = formula (derived FTE/EUR leaves).
- **Green** = cross-sheet ref (Plan-Forecast aggregations, Actuals-side forecast monitoring columns, Capacity actual-FTE cross-refs).

No new colour convention. Critical Rule 3 preserved.

### 6. Deliverables (v1)

1. `budget_planner/actuals_hours.py` — new sheet builder.
2. Collapsible forecast-EUR column band on `Actuals-Hours` with per-leaf EUR formulas.
3. `hours_per_fte_month` Config constant + named range.
4. `budget_planner/plan_forecast.py` — new sheet (Plan-Forecast) aggregating leaves.
5. `Actuals` sheet extension: parallel forecast monitoring columns with matching CF.
6. `Summary` extension: forecast EUR column per year + CF.
7. `Trends` extension: Cumulative Forecast EUR line; Person Load actual-side columns.
8. `Capacity` extension: actual FTE + variance FTE + utilisation % + actual EU TOTAL row with ceiling check.
9. `_META` entry for `Actuals-Hours` and `Plan-Forecast`; reader and rebuild coverage.
10. Tests covering: builder, leaf EUR chain, aggregation sums, CF rules, reader round-trip, EU-ceiling check on actuals.

### 7. Open questions

- Exact placement of `hours_per_fte_month` in Config (next free cell vs. near `fte_rate`).
- Whether `Actuals-Hours` should live **next to** `Capacity` in tab order, or **next to** `Actuals` to group monitoring sheets. Recommend: next to `Actuals` — audit narrative flows Plan → Capacity → Actuals-Hours → Actuals → Summary.
- Forecast Breach-In formula wording when the forecast line is still climbing past its last logged month — treat projection beyond last month as extrapolation of the current slope, or stop at last logged month? Recommend: stop at last logged month; audit-safe, no speculative alarms.

## Roadmap

### v2 — ergonomics & adoption

Make v1 less painful at month-end.

- **Paste-in CSV import flow.** Named input region on `Actuals-Hours`; formulas pick up pasted blocks. Removes per-cell manual entry for teams with a timesheet tool.
- **Personal view per team member.** One auto-generated sheet per person, filtered view of their WPs × months with plan/actual/variance. Read-only; print-clean.
- **Forecast vs. Actual EUR reconciliation table on Trends.** Per-WP, months where both streams exist, show the gap. Surfaces systematic rate/indirect-cost drift single-signal views hide. Needs ~3 months of dual-signal history.

### v3 — analytical depth (once ~1 year of data exists)

- **Person-level Gantt / timeline view.** Load-density heatmap per person across all WPs and months, WP active bands overlaid. Supports quarterly staffing conversations.
- **"What if person X leaves / joins late" scenarios.** Structured delta on top of snapshot/restore: pick a person, zero them out from month M, compute the shortfall per WP.

### v4 — advanced / research (speculative; require new inputs)

- **Skill/grade coverage analysis.** Requires a skill taxonomy or expanded grade semantics. Answers "if X leaves, who else can absorb their WP2 allocation?" Build once taxonomy stabilises.
- **Workload smoothing recommendations.** Solver-style: given ceilings, WP budgets, and active-month bounds, propose FTE shifts flattening over-cap months. Build only if v1 alarms prove that manual re-plans are too slow.
