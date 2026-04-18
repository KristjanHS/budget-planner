# Directory Map

```
budget_planner/
  __init__.py          # Package init
  __main__.py          # CLI: init, init --demo, rebuild subcommands + snapshot/restore
  config.py            # Project-specific data (WPs, team, external commitments, demo seeds) — seed for init only
  config_sheet.py      # Config sheet builder (project params + tbl_WorkPackages + team roster)
  styles.py            # Font/fill/border constants, style_cell() helper, apply_year_groups()
  plan.py              # Plan sheet: compact header + planned FTE block
  capacity.py          # Capacity sheet: per-person monthly FTE allocation grid
  summary.py           # Summary sheet: budget overview + calendar-year aggregation
  actuals.py           # Actuals sheet: EUR per WP per month from accountant
  trends.py            # Trends sheet: burn rate, cumulative chart, person load
  timelog_plan.py      # Timelog Plan sheet: vertical-stacked bands (WP summary / Proposed / Assigned / Person summary)
  timelog_solver.py    # Pure two-pass greedy solver driving the Timelog Plan Proposed grid
  reader.py            # Layout detection + Config/Capacity/Actuals/Timelog Plan readers
  workbook.py          # Orchestrator: init_workbook() + rebuild_workbook() + seed_demo_data() + snapshot/restore
  layout.py            # Centralized layout constants, column structure, calendar year helpers
tests/
  conftest.py          # Fixtures: fresh wb paths
  test_workbook.py     # Structural invariants (sheets, params, inactive months)
  test_team.py         # Config roster section tests
  test_capacity.py     # Capacity sheet per-person groups, blue input, inactive months
  test_plan_v2.py      # Plan <-> Capacity cross-sheet formula linkage + font color
  test_reader_v2.py    # Reader: config, capacity extraction, discover_layout
  test_reader_new.py   # Config sheet reader, Actuals reader, parameterized constants
  test_cli.py          # CLI flags (--demo, --snapshot, --restore)
  test_config.py       # Config referential integrity (WP refs, person refs)
  test_config_sheet.py # Config sheet builder + round-trip
  test_integration.py  # Cross-sheet data flow + round-trip preservation
  test_rebuild.py      # Rebuild preservation, orphaned data, month count changes
  test_grouping.py     # Column grouping: groups exist, correct collapse state
  test_parameterization.py  # Builder decoupling: no config.py imports in rebuild path
  test_styles.py       # style_cell() unlocked param (for when protection is re-enabled)
  test_trends.py       # Trends sheet: burn rate, chart, person load, CF
  test_e2e.py          # End-to-end generation + LibreOffice recalc
scripts/
  recalc.py            # LibreOffice recalculation script
  office/              # LibreOffice helpers
```

## Sheet responsibilities

One bullet per worksheet — what it is, what feeds it, what it feeds. Build order matches `_build_all_sheets` in `workbook.py`. See `docs/architecture.md` for the full cross-sheet formula graph.

- **Config** — Single source of truth: project parameters, WP budget table, person-WP assignments, external commitments, WP schedule, team roster, reporting periods, grade rate table.
  - IN ← user edits only (blue input cells).
  - OUT → every other sheet (params, named ranges `project_start` / `fte_rate`, WP list, roster).

- **Capacity** — Primary planning surface: per-person monthly FTE allocation across programme WPs and external projects. Auto-distributed from Config assignments on init/rebuild; user tunes month-by-month.
  - IN ← Config (assignments, WP list, roster, months); Actuals-Hours (Actual-FTE to-date band).
  - OUT → Plan (planned FTE formulas); Summary (plan FTE per year).

- **Plan** — Compact 1-row header + Planned FTE block. Pure formula mirror of Capacity shaped as WP × month.
  - IN ← Capacity (SUM per WP × month).
  - OUT → Summary (Plan EUR = Plan FTE × rate); Trends (planned cumulative curve).

- **Timelog Plan** — Prescriptive whole-month 1.0-FTE WP allocation grid. Separate surface from Plan (see `docs/design-invariants.md` § Plan vs Timelog Plan). Vertical-stacked band layout (PER-WP SUMMARY / PROPOSED / ASSIGNED / PER-PERSON SUMMARY — all share the month-column range D..D+months-1, each band collapsible via `outline_level=1`). Solver writes the Proposed band (locked); the Assigned band starts as `=Proposed` and accepts blue-input overrides preserved across rebuild.
  - IN ← Config (WP budget + WP schedule + PERSON-WP ASSIGNMENTS + Team Roster including `Priority` + grade rates); `current_month` (per-WP window-months-remaining).
  - OUT → (terminal surface; no downstream sheet consumes it).

- **Actuals** — EUR actuals per WP per month (accountant input) plus a parallel forecast-monitoring block reading Plan-Forecast. CF greys monthly Actual EUR cells past `eur_actuals_through_month` (not-yet-closed months). Tab color: Blue.
  - IN ← user EUR input; Plan-Forecast (forecast block).
  - OUT → Summary (Actual EUR per year); Trends (cumulative actual curve).

- **Actuals-Hours** — Per-person monthly hour timelog plus a sibling forecast-EUR band (outline_level=1, collapsible). Sole source of forecast-EUR leaves. Tab color: Red (C00000). Consumes named ranges `hours_per_fte_month`, `eu_annual_ceiling_pm`, `roster_lookup`, `eur_actuals_through_month`.
  - IN ← user hour input (primary input surface for hours; two-cadence model).
  - OUT → Plan-Forecast (forecast-EUR leaves SUMmed per WP × month); Capacity (Actual-FTE to-date band).

- **Plan-Forecast** — WP × month forecast-EUR grid aggregating the leaf formulas on Actuals-Hours. Mirrors Plan shape. Tab color: Green (00B050).
  - IN ← Actuals-Hours (forecast-EUR leaves, SUM per WP × month).
  - OUT → Actuals (forecast monitoring block); Summary (Forecast EUR per year); Trends (Cumulative Forecast helper row).

- **Summary** — Calendar-year aggregation: Plan FTE, Plan EUR, Actual EUR, Forecast EUR, Variance EUR per WP per year (5 cols/year) + project total group. Budget overview + Form C reporting-period block.
  - IN ← Plan, Actuals, Plan-Forecast, Config (reporting periods).
  - OUT → (terminal read-only view; no downstream sheet consumes it).

- **Trends** — WP burn rate & exhaustion, cumulative EUR chart (Planned + Actual + Cumulative Forecast + Ceiling), person load (planned + actual-side cols). Tab color: Gold (FFD700).
  - IN ← Plan, Actuals, Plan-Forecast, Capacity.
  - OUT → (terminal read-only view + embedded chart).
