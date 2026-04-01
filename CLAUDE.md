# CLAUDE.md — Project Instructions for Claude Code

## What This Is

An openpyxl-based Excel workbook generator for EU project FTE budget planning. The user will iterate on this with you.

## Key Commands

```bash
# Set up dev environment
uv venv && uv pip install -e ".[dev]"

# Generate the workbook (demo mode — pre-filled with seed allocations)
python -m budget_planner

# Generate skeleton workbook (empty Capacity for manual entry)
python -m budget_planner --skeleton

# Update existing workbook (preserves user-entered data across all sheets)
python -m budget_planner --update output/budget_plan.xlsx

# Snapshot / restore for scenario comparison
python -m budget_planner --snapshot output/budget_plan.xlsx
python -m budget_planner --restore output/budget_plan.xlsx

# Recalculate formulas (MANDATORY after any generation)
python scripts/recalc.py output/budget_plan.xlsx

# Run full test suite (62 tests including recalc + structural checks)
pytest tests/ -v
```

## Critical Rules

1. **Always use Excel formulas, never hardcode calculated values in Python.** The spreadsheet must remain dynamic.
2. **After every generation, run `pytest tests/ -v`** and verify all pass before presenting to user.
3. **Font color convention is sacred**: Blue = input, Black = formula, Green = cross-sheet ref. Never mix these.
4. **Inactive month cells must be greyed out and empty.** Auditors flag time logged outside WP active periods.
5. **Named ranges `project_start` and `fte_rate`** must always exist and point to Plan!$B$2 and Plan!$E$2.
6. **_META sheet must be updated** whenever you change the workbook structure. It's the primary documentation for future AI agents.

## Architecture Overview

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
  test_plan_v2.py      # Plan ↔ Capacity cross-sheet formula linkage + font color
  test_reader_v2.py    # Reader v2: config, team, capacity extraction
  test_cli.py          # CLI flags (--skeleton, --snapshot, --restore)
  test_config.py       # Config referential integrity (WP refs, person refs)
  test_integration.py  # Cross-sheet data flow + round-trip preservation
  test_e2e.py          # End-to-end generation + LibreOffice recalc
scripts/
  recalc.py            # LibreOffice recalculation script
  office/              # LibreOffice helpers (unchanged)
```

- `SPEC.md` — Full specification with business context, design decisions, and known limitations.
- `CONVENTIONS.md` — Spreadsheet formatting and structural conventions.

## Common Tasks

### "Replace example WPs with real data"
→ Edit `WORK_PACKAGES` in `budget_planner/config.py`. Run `python -m budget_planner` + pytest.

### "Replace example team with real data"
→ Edit `TEAM` and `EXTERNAL_COMMITMENTS` in `budget_planner/config.py`. Run `python -m budget_planner` + pytest.

### "Add a new column/metric to Summary"
→ Modify `budget_planner/summary.py`. Update `budget_planner/meta.py` text. Run + pytest.

### "Adjust team member WP shares"
→ Edit in the Team sheet directly (blue input cells), or update `TEAM` in `config.py` and regenerate.

### "Import actuals from timesheet"
→ Not yet implemented. Would be a separate script that reads CSV/XLSX and uses `generate_workbook(--update)`.

### "Compare budget scenarios"
→ Use `--snapshot` to save current state, make changes, compare. Use `--restore` to revert.

## Testing Approach

Run `pytest tests/ -v` after every change. The 62 tests cover:

1. Formula recalculation (0 errors via LibreOffice)
2. Structural invariants: sheets, global params, named ranges, config table
3. Inactive month cells empty (Plan + Capacity, with grey fill check)
4. Actual FTE zeros in fresh workbook
5. Font color: black for formulas (Plan), blue for input (Capacity, Team)
6. Plan ↔ Capacity cross-sheet SUM formula linkage
7. Capacity: per-person groups, TOTAL/REMAINING formulas, external rows
8. Team: roster, share matrix, validation sums
9. Config referential integrity (WP refs, person refs)
10. CLI: --skeleton, --snapshot/--restore
11. Reader v2: config, team, capacity extraction
12. Integration: round-trip preserves actuals, capacity edits, team roster
13. E2E: demo + skeleton generation with LibreOffice recalc
