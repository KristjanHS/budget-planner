# Monthly hours timelog v2 — design

**Date:** 2026-04-14
**Status:** Design (pre-implementation-plan)
**Audience:** PI / project lead; secondarily auditor
**Supersedes for scope:** Roadmap v2 section of [`2026-04-13-hours-timelog-design.md`](2026-04-13-hours-timelog-design.md)

## Problem

v1 landed (T1–T16 + review + Tier A simplification). Three gaps surfaced in review:

1. Dense per-person layout on `Capacity` / `Actuals-Hours` with no fast way to hide other people.
2. Forecast and actual EUR never compared head-to-head — systematic drift (rate mismatch, indirect-cost factor shift, hours-per-FTE mis-calibration) is invisible until large.
3. **Every per-month derived readout on the actual side misreads under the PI's stated entry pattern.** Hours are entered with months skipped and a cumulative figure dropped into a later month's cell. Actual EUR arrives at WP level (not per-person) every 2–3 months. Per-month Actual FTE / Variance FTE / Utilisation, peak FTE, recent-month-slope Breach-In all produce garbage under this cadence.

## Goal

Three workstreams on top of v1:

- **W1: Row grouping** on `Capacity` + `Actuals-Hours`. One outline level per person; person-header stays visible, block collapses. Mirrors column outline grouping already present.
- **W2: Actual-side view redesign.** Input-cell semantics unchanged (monthly, blue, "hours in M", blank = not yet entered). All derived readouts that depended on monthly-granularity entry are replaced with **cumulative-to-date** or **PM-to-date** equivalents gated by through-month markers. Deprecates monthly Actual FTE / Variance FTE / Utilisation (as the "live" view), peak FTE, recent-month-slope Breach-In.
- **W3: Forecast-vs-Actual EUR reconciliation section** on Trends — per-WP cumulative-vs-cumulative summary table with CF, aggregate gap-over-time line chart, three through-month health cells.

## Non-goals for v2

- Paste-in CSV import (deferred to v3 — removes skip-and-cumulate pain at source).
- Personal per-team-member sheets (dropped; row grouping delivers the same view).
- Reinterpreting input cells as cumulative-through-month semantics (rejected — keeps auditor mental model intact).
- Workload solver.

See [Roadmap](#roadmap).

## Data-entry reality

Two cadences, both skip-and-cumulate:

- **Hours** — entered on `Actuals-Hours` per person × WP × month. The PI will leave intermediate months blank and drop cumulative totals into a later month's cell to save time. Blank cells inside the through-range therefore cannot be distinguished from legitimate zero months.
- **Actual EUR** — provided at **WP level only** (no per-person breakdown), on a **2–3 month cadence** (finance-close lag). Cells on `Actuals` are blank between report deliveries; a consolidated figure lands at the report month.

Design principle derived from this reality: **no derived view reads a single-month cell as authoritative for that month.** All actual-side readouts are cumulative-through-a-through-month or a to-date rate.

## Data model — new named ranges (Config sheet)

| Name | Cell | Meaning | User-entered? |
|---|---|---|---|
| `hours_actuals_through_month` | next free Config input, adjacent to `eur_actuals_through_month` | Latest month whose `Actuals-Hours` totals the PI trusts as complete | Yes (blue) |
| `months_elapsed_hours` | derived | `DATEDIF(project_start, hours_actuals_through_month, "M") + 1` | Black |
| `months_elapsed_eur` | derived | Analogous using `eur_actuals_through_month` | Black |
| `reconciled_through_month` | derived | `MIN(eur_actuals_through_month, hours_actuals_through_month)`; normally = `eur_actuals_through_month` | Black |

`eur_actuals_through_month`, `hours_per_fte_month`, `project_start`, `fte_rate`, `roster_lookup`, `eu_annual_ceiling_pm` already exist from v1.

**Docs update (post-_META removal 2026-04-14):** document the two-cadence reality as a design-time assumption in `docs/design-invariants.md`; future readers (AI or human) understand why two independent through-months exist, not one.

## W1 — Row grouping

- On `Capacity` and `Actuals-Hours`, each person's block (WP rows + TOTAL + EU TOTAL) gets `outline_level = 1`. Person-header row stays at level 0.
- `sheet_properties.outlinePr.summaryBelow = False` so the outline button sits next to the person header, not below.
- Implementation lives in the person-section geometry helper already shared between `capacity.py` and `actuals_hours.py` (single source). One outline level; no two-level grouping.

**Not in scope:** row grouping on Summary / Actuals / Trends (already compact).

**Tests:**

- Per sheet: every data row of every person carries `outline_level == 1`; person header at `0`.
- `summaryBelow is False`.
- Round-trip: outline levels survive read/rebuild.

## W2 — Actual-side view redesign

### W2.1 — Capacity sheet

**Remove from the live band:** per-WP-month Actual FTE, per-WP-month Variance FTE, per-month Utilisation % strip (if present).

**Retain but default-collapsed:** the raw monthly Actual FTE + Variance FTE bands are wrapped in their own `outline_level = 1` column group, default collapsed. Band header labeled "Raw monthly — reliable only when hours were entered per-month". No formula changes; pure layout + default outline state. One click expands when the PI did enter monthly and wants a monthly heat-map.

**New to-date column block** on the far right (own outline group, collapsible):

| Column | Scope | Formula |
|---|---|---|
| Plan PM to date | per WP row | `SUMIFS(plan FTE cells, month ≤ hours_actuals_through_month)` |
| Actual PM to date | per WP row | `SUM(Actuals-Hours per WP cells, month ≤ hours_actuals_through_month) / hours_per_fte_month` |
| Variance PM to date | per WP row | `Plan PM − Actual PM` |
| Variance % | per WP row | `(Plan − Actual) / Plan`. Red CF ≥ ±20%; both directions flag. |

Per-person summary rows gain:

| Row | Formula |
|---|---|
| Actual PM to date | SUM across WPs |
| Plan PM to date | SUM across WPs |
| Utilisation to date | `Actual PM to date / (Max FTE × months_elapsed_hours)`. Red ≥ 100%, amber ≥ 80% — mirrors v1 per-month CF, different denominator |
| EU Actual PM to date | sum across programme WPs only |
| EU Ceiling pro-rated | `eu_annual_ceiling_pm × (months_elapsed_hours / 12)` clipped to current calendar year |

**EU ceiling check** becomes **pro-rated-to-date**: red CF on `EU Actual PM to date > EU Ceiling pro-rated`. Fires earlier than year-end; honest under partial-year data.

**Keep:** all Plan FTE monthly cells (planning is forward-looking; monthly granularity meaningful). Planned TOTAL / EU TOTAL per-person per month.

**Colour:** new columns all green (cross-sheet refs to `Actuals-Hours`) or black (same-sheet). No blue. Rule 3 preserved.

### W2.2 — Actuals sheet

| Column | Action |
|---|---|
| Ceiling | Keep |
| Cumulative EUR (Actual) | Keep — correct at through-month edge |
| Variance vs Ceiling | Keep |
| % Ceiling | Keep |
| Breach In (months) | **Change formula.** Old: recent-month-slope extrapolation. New: `(ceiling − cum EUR) / (cum EUR / months_elapsed_eur)` — cumulative-slope from project start through `eur_actuals_through_month`. Auditor-safe, no per-month slope magic. |
| Forecast parallel block (v1) | Same treatment — forecast Breach In uses cumulative slope through `hours_actuals_through_month` |
| "As-of" header cells | **New.** Above the Actual block: `Actual EUR as of <eur_actuals_through_month>`. Above the Forecast block: `Forecast EUR as of <hours_actuals_through_month>`. Two small green cross-ref cells. |

### W2.3 — Trends sheet

- **LineChart series clipping.** Cumulative Actual EUR points past `eur_actuals_through_month` return `NA()`; Cumulative Forecast EUR past `hours_actuals_through_month` returns `NA()`. Lines end cleanly at their respective through-month edges. Cumulative Planned + Ceiling run through project end as today.
- **Person Load table.** Remove peak FTE column (meaningless under skip-and-cumulate). Replace avg FTE with **Avg FTE to date** = `actual PM to date / months_elapsed_hours`. Add **Plan avg FTE to date**. Add **Variance to-date %** with same CF as Capacity's Variance %.

### W2.4 — Summary sheet

- Annual columns: **per-column partial-year shading.** Forecast column shades by `hours_actuals_through_month`; Actual column shades by `eur_actuals_through_month`; Variance cell shades by the stricter `MIN`. `year_coverage_pct = months_of_year_covered / 12`.
- Variance vs Plan CF: applies only to fully-covered years. Partial years show the number with no CF (no false alarms on incomplete data).

### W2.5 — Actuals-Hours + Plan-Forecast

- **Actuals-Hours.** Per-leaf forecast-EUR cells cosmetically shaded past `hours_actuals_through_month` (grey, matching inactive-month convention). Formulas unchanged. New summary rows above the existing TOTAL row: "Cumulative hours to date" and "Actual PM to date", both cumulative-through-`hours_actuals_through_month`.
- **Plan-Forecast.** Cosmetic shading past `hours_actuals_through_month`. No structural change.

## W3 — Forecast-vs-Actual EUR reconciliation section

### Placement

New section on `Trends`, between the existing LineChart and the Person Load table. Clear section header band.

### Health cells (three stacked rows)

| Label | Value |
|---|---|
| Actuals EUR through | `eur_actuals_through_month` (green) |
| Hours through | `hours_actuals_through_month` (green) |
| Reconciled through | `reconciled_through_month` = MIN of the two (black; derived) |

Answer "should I trust this?" at a glance. Expected normal state: `Reconciled through` ≈ `Actuals EUR through` (EUR lags hours by 2–3 months).

### Summary table (per-WP, one row per programme WP)

| Column | Formula | CF |
|---|---|---|
| WP | static label | — |
| Cum Forecast EUR | `SUMIFS(Plan-Forecast[WP, months ≤ reconciled_through_month])` | — |
| Cum Actual EUR | `SUMIFS(Actuals[WP, months ≤ reconciled_through_month])` | — |
| Gap EUR | `Cum Actual − Cum Forecast` | — |
| Gap % | `Gap EUR / Cum Forecast` | red ≥ ±10%, amber ≥ ±5% (absolute) |

Bottom totals row: sum each column across WPs; Gap % of totals.

### Gap-over-time chart

Single LineChart, placed beside or below summary table.

- X axis: months from `project_start` to `reconciled_through_month`.
- Single series: `cumulative Actual EUR (all WPs) − cumulative Forecast EUR (all WPs)` per month.
- Series returns `NA()` for months > `reconciled_through_month` → line ends cleanly at reconciled edge.
- Gridlines, zero reference line (dashed), no secondary axis. Drift is slope.

**Cumulative-only math, no per-month differencing** — immune to skip-and-cumulate on either side.

### Empty-state behaviour

First month of project, no overlap yet:

- `reconciled_through_month` evaluates to `project_start` or `#N/A`.
- Summary table cells and gap chart series all return blank / NA.
- Health cells still show the two independent through-months — user sees "Reconciled through: —" and knows not to trust the table yet. No special hide logic.

## Colour convention (all three workstreams)

- **Blue:** only `hours_actuals_through_month` (single new input cell).
- **Green:** all cross-sheet refs (to-date block, reconciliation section, as-of headers).
- **Black:** same-sheet derived (months-elapsed cells, reconciled-through-month).

No new convention. Critical Rule 3 preserved.

## Critical Rules check

- **Rule 1** (formulas, not values): every new cell is a formula.
- **Rule 3** (colour): blue/green/black mapping unchanged.
- **Rule 4** (inactive-month greying + empty): new shading rule (`> through-month`) is separate and non-conflicting.
- **Rule 5** (`project_start` + `fte_rate` named ranges): untouched.
- ~~**Rule 6** (_META reflects current structure)~~ — OBSOLETE: `_META` sheet removed 2026-04-14; invariants now live in `docs/design-invariants.md`.

## Deliverables

1. New named ranges: `hours_actuals_through_month`, `months_elapsed_hours`, `months_elapsed_eur`, `reconciled_through_month`.
2. `budget_planner/capacity.py`: wrap raw monthly Actual/Variance bands in default-collapsed outline group; add to-date column block; add per-person utilisation-to-date + EU-ceiling-pro-rated rows.
3. `budget_planner/actuals.py`: update Breach-In formula to cumulative-slope; add "As-of" header cells above Actual and Forecast blocks.
4. `budget_planner/trends.py`: LineChart series clipping via `NA()`; Person Load table updates; new Reconciliation section (health cells, summary table, gap chart).
5. `budget_planner/summary.py`: per-column partial-year shading; CF gated on fully-covered years.
6. `budget_planner/actuals_hours.py`: cosmetic shading past `hours_actuals_through_month`; new summary rows (Cumulative hours to date, Actual PM to date). Shared geometry helper gains row-group application for W1.
7. `budget_planner/plan_forecast.py`: cosmetic shading past `hours_actuals_through_month`.
8. ~~`budget_planner/meta.py`~~ — OBSOLETE (2026-04-14 `_META` removal). Instead: update `docs/design-invariants.md` + `docs/internals.md` with new named ranges, two-cadence reality note, Reconciliation section, to-date block, collapsed raw-monthly band.
9. `budget_planner/reader.py` + rebuild: new named ranges round-trip; to-date block is derived (no input), skip on read; row outline levels preserved on rebuild.

## Test coverage (additions)

- **W1.** Outline levels correct on both sheets; header at 0, body at 1; `summaryBelow is False`. Round-trip preserves levels.
- **To-date block.** Cumulative math correct under: (a) monthly-only entry, (b) skip-and-cumulate entry, (c) mixed entry. Variance % CF fires at ±20%.
- **Breach-In.** Cumulative-slope extrapolation vs. ceiling; stable under skip-and-cumulate.
- **Summary partial-year shading.** Per-column gating; CF suppressed on partial years.
- **Reconciliation.** Summary table math with two-cadence reality (EUR lagging hours); gap-chart `NA()` past reconciled-through; health cells show three correct through-months; empty-state returns blank.
- **LineChart series clipping.** Chart builds cleanly; no stray trailing-zero points past through-months.
- **Reader.** All four new named ranges survive round-trip; row outline levels survive rebuild.

## Open questions (non-blocking)

- Exact Config layout for `hours_actuals_through_month` — adjacent to `eur_actuals_through_month` recommended.
- Summary partial-year shading colour — reuse existing inactive-month grey recommended (matches convention).
- Reconciliation section vertical placement on Trends — between LineChart and Person Load recommended; narrative flows Plan/Actual/Forecast (chart) → Reconciliation (drift) → Person Load (capacity).

## Roadmap

### v3 — data-entry ergonomics (removes skip-and-cumulate pain at source)

- Paste-in CSV import flow on `Actuals-Hours`. Named input region; formulas pick up pasted blocks. When this lands, blank-means-unentered ambiguity dissolves and per-month views could potentially be revived.
- Per-WP small-multiples on the reconciliation section (if aggregate gap chart proves inadequate for spotting WP-specific drift patterns the summary table misses).
- Person-level Gantt / timeline view.

### v4 — analytical depth (once ~1 year of data exists)

- "What if person X leaves / joins late" scenarios — structured delta on snapshot/restore.
- Skill/grade coverage analysis (requires new taxonomy).
- Workload smoothing solver.
