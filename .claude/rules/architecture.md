---
paths: "**/*.py"
---
# Architecture

```
budget_planner/
  __init__.py          # Package init
  __main__.py          # CLI: python -m budget_planner [--skeleton|--snapshot|--restore]
  config.py            # Project-specific data (WPs, team, external commitments, constants)
  styles.py            # Font/fill/border constants, style_cell() helper
  plan.py              # Plan sheet: globals + config table + planned FTE block
  team.py              # Team sheet: roster + WP share seed matrix
  capacity.py          # Capacity sheet: per-person monthly FTE allocation grid
  summary.py           # Summary sheet: budget overview + annual aggregation
  meta.py              # _META sheet: AI-agent-readable documentation
  actuals.py           # Actuals sheet: EUR per WP per month from accountant
  trends.py            # Trends sheet: burn rate, cumulative chart, person load
  reader.py            # discover_layout() for layout detection (v1 + v2)
  workbook.py          # Orchestrator: create + update + snapshot/restore modes
tests/
  conftest.py          # Fixtures: fresh wb, recalc
  test_workbook.py     # Structural invariants (sheets, params, inactive months)
  test_recalc.py       # LibreOffice formula verification
  test_team.py         # Team sheet roster + share matrix
  test_capacity.py     # Capacity sheet per-person groups, blue input, inactive months
  test_plan_v2.py      # Plan <-> Capacity cross-sheet formula linkage + font color
  test_reader_v2.py    # Reader v2: config, team, capacity extraction
  test_cli.py          # CLI flags (--skeleton, --snapshot, --restore)
  test_config.py       # Config referential integrity (WP refs, person refs)
  test_integration.py  # Cross-sheet data flow + round-trip preservation
  test_trends.py       # Trends sheet: burn rate, chart, person load, CF
  test_e2e.py          # End-to-end generation + LibreOffice recalc
scripts/
  recalc.py            # LibreOffice recalculation script
  office/              # LibreOffice helpers
```

