---
paths: "**/*.py"
---
# Architecture

```
budget_planner/
  __init__.py          # Package init
  __main__.py          # CLI: init, rebuild, rebuild --skeleton subcommands (legacy flags still work)
  config.py            # Project-specific data (WPs, team, external commitments, constants) — seed for init only
  config_sheet.py      # Config sheet builder + reader (project params + tbl_WorkPackages)
  styles.py            # Font/fill/border constants, style_cell() helper, apply_year_groups()
  plan.py              # Plan sheet: globals + config table + planned FTE block
  team.py              # Team sheet: roster + WP share seed matrix
  capacity.py          # Capacity sheet: per-person monthly FTE allocation grid
  summary.py           # Summary sheet: budget overview + annual aggregation
  meta.py              # _META sheet: AI-agent-readable documentation
  actuals.py           # Actuals sheet: EUR per WP per month from accountant
  trends.py            # Trends sheet: burn rate, cumulative chart, person load
  reader.py            # Layout detection + Config/Team/Capacity/Actuals readers, parameterized constants
  workbook.py          # Orchestrator: init_workbook() + rebuild_workbook() + snapshot/restore
tests/
  conftest.py          # Fixtures: fresh wb paths
  test_workbook.py     # Structural invariants (sheets, params, inactive months)
  test_team.py         # Team sheet roster + share matrix
  test_capacity.py     # Capacity sheet per-person groups, blue input, inactive months
  test_plan_v2.py      # Plan <-> Capacity cross-sheet formula linkage + font color
  test_reader_v2.py    # Reader v2: config, team, capacity extraction
  test_reader_new.py   # Config sheet reader, Actuals reader, parameterized constants
  test_cli.py          # CLI flags (--skeleton, --snapshot, --restore)
  test_config.py       # Config referential integrity (WP refs, person refs)
  test_config_sheet.py # Config sheet builder + round-trip
  test_integration.py  # Cross-sheet data flow + round-trip preservation
  test_rebuild.py      # Rebuild preservation, orphaned data, month count changes
  test_protection.py   # Cell protection: formula cells locked, blue cells unlocked
  test_grouping.py     # Column grouping: groups exist, correct collapse state
  test_parameterization.py  # Builder decoupling: no config.py imports in rebuild path
  test_styles.py       # Style helpers, apply_year_groups()
  test_trends.py       # Trends sheet: burn rate, chart, person load, CF
  test_e2e.py          # End-to-end generation + LibreOffice recalc
scripts/
  recalc.py            # LibreOffice recalculation script
  office/              # LibreOffice helpers
```

