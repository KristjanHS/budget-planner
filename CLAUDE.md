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

# Run full test suite (165 tests including recalc + structural checks)
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
  conftest.py          # Fixtures: fresh wb, recalc, wb_with_actuals
  test_workbook.py     # Structural checks (parametrized per WP)
  test_roundtrip.py    # Reader + update preserves actuals
  test_recalc.py       # LibreOffice formula verification
  test_team.py         # Team sheet roster + share matrix
  test_capacity.py     # Capacity sheet per-person groups + conditional formatting
  test_plan_v2.py      # Plan ↔ Capacity cross-sheet formulas
  test_reader_v2.py    # Reader v2 format (Team, Capacity, config)
  test_cli.py          # CLI flags (--skeleton, --snapshot, --restore)
  test_config.py       # Config validation (TEAM, EXTERNAL_COMMITMENTS)
  test_integration.py  # Multi-sheet round-trip integration
  test_e2e.py          # End-to-end generation + recalc
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

Run `pytest tests/ -v` after every change. The 165 tests cover:

1. Formula recalculation (0 errors via LibreOffice)
2. Sheet existence (_META, Plan, Team, Capacity, Summary)
3. Global parameters (project start date, FTE rate)
4. Named ranges (project_start, fte_rate)
5. Config table structure and data validity (parametrized per WP)
6. Planned FTE: cross-sheet SUM formulas from Capacity sheet
7. Actual FTE block (zeros in fresh workbook)
8. Inactive month cells (empty outside WP active periods)
9. Summary sheet WP count matches config
10. _META sheet has documentation content
11. Team sheet: roster structure, WP share matrix, validation sums
12. Capacity sheet: per-person groups, external commitments, conditional formatting
13. Config validation: TEAM references valid WPs, EXTERNAL_COMMITMENTS valid persons
14. CLI: --skeleton mode (empty Capacity), --snapshot/--restore
15. Reader v2: discovers Team + Capacity layout, extracts values
16. Integration: multi-sheet round-trip preserves all user data
17. E2E: full generation + LibreOffice recalc verification
