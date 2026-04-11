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
  meta.py              # _META sheet: AI-agent-readable documentation
  actuals.py           # Actuals sheet: EUR per WP per month from accountant
  trends.py            # Trends sheet: burn rate, cumulative chart, person load
  reader.py            # Layout detection + Config/Capacity/Actuals readers
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
