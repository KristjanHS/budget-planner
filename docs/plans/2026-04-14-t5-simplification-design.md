# T5 Simplification — Capacity To-Date Block (Kaizen Redesign)

**Date:** 2026-04-14
**Supersedes:** Task 5 in `2026-04-14-hours-timelog-v2-plan.md`
**Status:** Design approved (Option B + A borrowing). Not yet implemented.

## Context

The currently-landed T5 shipped a 4-column to-date block on Capacity
(`Plan PM`, `Actual PM`, `Variance PM`, `Variance %`) with:

- Two-pass build (`build_capacity_sheet` reserves columns →
  Actuals-Hours builds → `apply_capacity_todate_formulas` back-patches).
- `SUM + INDEX` windowed-cumulation formulas on every WP row, twice
  (plan side and actuals side).
- `person_roster_rows` mapping in `capacity_layout` to dodge the
  VLOOKUP-on-formula-cells limitation.
- Variance-% Conditional Formatting on every WP row + TOTAL + EU TOTAL
  (≥ `WP_rows + 2·TEAM` rules).
- T6 extended this with `Utilisation` + `EU Ceiling PM` columns and
  three extra CF rules per person — duplicating signals already present
  on the planned side.

The user flagged this as "too complicated". The kaizen review agreed:
the block crossed the YAGNI line at column 3.

## Simplification (chosen: Option B + A-borrow)

**Reduce to 2 columns:** `Actual PM`, `Variance PM`.

**Move cumulation to Actuals-Hours** (borrowed from T9's planned
"Cumulative hours to date" column): Capacity's `Actual PM` becomes a
single green cross-sheet reference divided by `hours_per_fte_month`,
with no SUM+INDEX on the Capacity side.

### Per-row fill rules

| Row type        | `Actual PM`                                           | `Variance PM`                    |
| --------------- | ----------------------------------------------------- | -------------------------------- |
| WP row          | `=Actuals-Hours!$<cum_col>$<ah_row>/hours_per_fte_month` | blank + grey                     |
| External row    | blank + grey                                          | blank + grey                     |
| TOTAL row       | `=SUM(<actual_col>_cells_of_this_person_wp_rows)`     | `=<plan_sum_to_date> - <actual>` |
| EU TOTAL row    | `=SUM(<actual_col>_cells_of_this_person_wp_rows)`     | `=<plan_sum_to_date> - <actual>` |
| REMAINING row   | blank + grey                                          | blank + grey                     |

`<plan_sum_to_date>` on the TOTAL row =
`SUM($C$TOTAL:INDEX($C$TOTAL:$AU$TOTAL, 1, hours_actuals_through_month))`.
The TOTAL row's planned monthly cells already sum all WPs per month,
so a single windowed SUM+INDEX suffices. Same pattern for EU TOTAL
(whose planned monthly cells already sum EU-flagged WPs per month).

### Conditional formatting

**One rule** per person: red fill on `Variance PM < 0` on TOTAL and
EU TOTAL rows. Total = `2·TEAM` rules on the to-date block.

No variance-% CF. No utilisation CF. No EU-ceiling CF (the planned
side already has EU-ceiling CF; that's the authoritative check).

### Actuals-Hours addition (pulled forward from T9)

Add one new column to the Actuals-Hours sheet — `Cum hours (to
`hours_actuals_through_month`)`. One formula per WP row + external
row + TOTAL + EU TOTAL:

```
=SUM($C<row>:INDEX($C<row>:$AU<row>, 1, hours_actuals_through_month))
```

This is the windowed cumulative hours total. Capacity reads from this
cell divided by `hours_per_fte_month` — no cross-sheet SUM+INDEX.

Return `cum_hours_col` in `actuals_hours_layout` so `apply_capacity_todate_formulas`
can reference it.

## What drops

| Dropped                                  | Reason                                           |
| ---------------------------------------- | ------------------------------------------------ |
| `Plan PM` dedicated column               | Absorbed inline into Variance formula on TOTAL rows. |
| `Variance %` column                      | No use case beyond visual — user eyeballs absolute gap. |
| Variance-% red CF (≥ ±20%)               | Redundant with Variance PM < 0 on TOTAL rows.    |
| `Utilisation` column + CF                | Planned-side already has util CF (`>0.8·Max FTE`). |
| `EU Ceiling PM` column + CF              | Planned-side already has EU-ceiling CF.          |
| `person_roster_rows` in capacity_layout  | No longer needed — util formula is gone.         |
| Per-WP Plan PM SUM+INDEX formulas        | No per-WP plan column; plan lives only at TOTAL rows. |
| Per-WP Actual PM SUM+INDEX formulas      | Replaced by single cross-sheet cell ref.         |

## What stays

- Two-pass build (still needed so Capacity can reference Actuals-Hours
  rows), but `apply_capacity_todate_formulas` becomes ~half the size.
- T6's raw-monthly default-collapse on Capacity (`outline_level = 1`
  + `hidden = True`) and warning band labels — these are orthogonal
  and still correct.
- `hours_actuals_through_month`, `hours_per_fte_month`, and
  `months_elapsed_hours` named ranges. `eu_annual_ceiling_pm` stays
  for the planned-side CF that already uses it.

## Implementation scope (when picked up)

1. **`budget_planner/actuals_hours.py`** — add `cum_hours_col` at
   `forecast_col_end + 2`. Write SUM+INDEX formula per WP / external /
   TOTAL / EU TOTAL row. Header label, green fill, return col index
   in layout.

2. **`budget_planner/capacity.py`**:
   - `build_capacity_sheet`: reduce to-date columns to 2
     (`todate_col_actual`, `todate_col_var`). Drop `todate_col_plan`,
     `todate_col_pct`, `todate_col_util`, `todate_col_ceil` from
     layout. Drop `person_roster_rows` from returned layout. Update
     band label comment.
   - `apply_capacity_todate_formulas`: rewrite to 2-col logic.
     Variance formula on TOTAL + EU TOTAL rows references
     the same-row planned monthly cells via SUM+INDEX (no cross-sheet).
     Actual formula on WP rows references Actuals-Hours `cum_hours_col`
     divided by `hours_per_fte_month`. TOTAL `Actual PM` sums WP
     actual cells. CF: red on Variance PM < 0 at TOTAL + EU TOTAL only.

3. **`budget_planner/workbook.py`** — `apply_capacity_todate_formulas`
   call signature may change if we pass `cum_hours_col` from the
   Actuals-Hours layout; update accordingly.

4. **Tests:**
   - `tests/test_capacity_todate.py` — rewrite entirely:
     - Drop `plan`, `pct`, `util`, `ceil` column assertions.
     - Add test: WP Actual PM formula is cross-sheet ref to
       Actuals-Hours cum_hours_col (no SUM+INDEX).
     - Add test: TOTAL Variance PM formula = `(plan_windowed_sum) -
       (actual_cell)`; plan sum comes from same-row planned TOTAL
       cells (no cross-sheet reference in plan side).
     - Demo-value round-trip: TOTAL Variance ≈ expected.
   - `tests/test_conditional_formatting.py` — update
     `test_capacity_rule_count` expected formula: drop Variance-%
     CF per WP row, drop util amber+red, drop EU-ceiling red on
     to-date. Add single `2·TEAM` term for Variance PM red on TOTAL +
     EU TOTAL.
   - Add small test: Actuals-Hours sheet has `Cum hours (to date)`
     header and per-WP-row SUM+INDEX formula.

5. **`budget_planner/meta.py`** — update _META Capacity block to
   describe the 2-col to-date layout; update Actuals-Hours block to
   mention the new cum-hours column.

## Rollout

Single commit `refactor(v2-T5): kaizen 2-col to-date + cum-hours on Actuals-Hours`.
Reverts the T5 4-col commit's Capacity layout, keeps T6's raw-monthly
collapse. Net effect: fewer cells, fewer formulas, fewer CF rules,
simpler test suite.

## Deferred

- T7, T8, T9 (less T9's cum-hours row, now absorbed into this design).
- T6 util/ceil extensions are dropped entirely — planned-side CF
  already covers both signals.
