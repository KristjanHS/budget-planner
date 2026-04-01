---
paths: "**/*.py"
---
# Architecture

```
budget_planner/
  __init__.py          # Package init
  __main__.py          # CLI: python -m budget_planner [--update|--skeleton|--snapshot|--restore]
  config.py            # Project-specific data (WPs, team, external commitments, constants)
  styles.py            # Font/fill/border constants, style_cell() helper
  plan.py              # Plan sheet: globals + config table + both FTE blocks
  team.py              # Team sheet: roster + WP share seed matrix
  capacity.py          # Capacity sheet: per-person monthly FTE allocation grid
  summary.py           # Summary sheet: annual aggregation
  meta.py              # _META sheet: AI-agent-readable documentation
  reader.py            # discover_layout() + read_workbook() for round-trip (v1 + v2)
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
  test_e2e.py          # End-to-end generation + LibreOffice recalc
scripts/
  recalc.py            # LibreOffice recalculation script
  office/              # LibreOffice helpers
```

## References

- `SPEC.md` — Full specification with business context, design decisions, and known limitations.
- `CONVENTIONS.md` — Spreadsheet formatting and structural conventions.

## Test Coverage

1. Formula recalculation (0 errors via LibreOffice)
2. Structural invariants: sheets, global params, named ranges, config table
3. Inactive month cells empty (Plan + Capacity, with grey fill check)
4. Actual FTE zeros in fresh workbook
5. Font color: black for formulas (Plan), blue for input (Capacity, Team)
6. Plan <-> Capacity cross-sheet SUM formula linkage
7. Capacity: per-person groups, TOTAL/REMAINING formulas, external rows
8. Team: roster, share matrix, validation sums
9. Config referential integrity (WP refs, person refs)
10. CLI: --skeleton, --snapshot/--restore
11. Reader v2: config, team, capacity extraction
12. Integration: round-trip preserves actuals, capacity edits, team roster
13. E2E: demo + skeleton generation with LibreOffice recalc
