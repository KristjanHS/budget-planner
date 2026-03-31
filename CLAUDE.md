# CLAUDE.md — Project Instructions for Claude Code

## What This Is

An openpyxl-based Excel workbook generator for EU project FTE budget planning. The user will iterate on this with you.

## Key Commands

```bash
# Set up dev environment
uv venv && uv pip install -e ".[dev]"

# Generate the workbook
python -m budget_planner

# Update existing workbook (preserves user-entered actuals and planned edits)
python -m budget_planner --update output/budget_plan.xlsx

# Recalculate formulas (MANDATORY after any generation)
python scripts/recalc.py output/budget_plan.xlsx

# Run full test suite (58 tests including recalc + structural checks)
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
  __main__.py          # CLI: python -m budget_planner [--update path]
  config.py            # Project-specific data (WP list, constants, duration)
  styles.py            # Font/fill/border constants, style_cell() helper
  plan.py              # Plan sheet: globals + config table + both FTE blocks
  summary.py           # Summary sheet: annual aggregation
  meta.py              # _META sheet: AI-agent-readable documentation
  reader.py            # discover_layout() + read_workbook() for round-trip
  workbook.py          # Orchestrator: create + update modes
tests/
  conftest.py          # Fixtures: fresh wb, recalc, wb_with_actuals
  test_workbook.py     # 13 structural checks (parametrized per WP)
  test_roundtrip.py    # Reader + update preserves actuals
  test_recalc.py       # LibreOffice formula verification
scripts/
  recalc.py            # LibreOffice recalculation script
  office/              # LibreOffice helpers (unchanged)
```

- `SPEC.md` — Full specification with business context, design decisions, and known limitations.
- `CONVENTIONS.md` — Spreadsheet formatting and structural conventions.

## Common Tasks

### "Replace example WPs with real data"
→ Edit `WORK_PACKAGES` in `budget_planner/config.py`. Run `python -m budget_planner` + pytest.

### "Add a new column/metric to Summary"
→ Modify `budget_planner/summary.py`. Update `budget_planner/meta.py` text. Run + pytest.

### "Add per-person breakdown"
→ This is listed as future work in SPEC.md. Would require a new block or sheet. Discuss design with user before implementing.

### "Import actuals from timesheet"
→ Not yet implemented. Would be a separate script that reads CSV/XLSX and uses `generate_workbook(--update)`.

## Testing Approach

Run `pytest tests/ -v` after every change. The 58 tests cover:

1. Formula recalculation (0 errors via LibreOffice)
2. Sheet existence (_META, Plan, Summary)
3. Global parameters (project start date, FTE rate)
4. Named ranges (project_start, fte_rate)
5. Config table structure and data validity (parametrized per WP)
6. Planned FTE formulas exist in validation columns
7. Actual FTE block (zeros in fresh workbook)
8. Inactive month cells (empty outside WP active periods)
9. Summary sheet WP count matches config
10. _META sheet has documentation content
11. Round-trip: reader discovers layout, extracts values
12. Round-trip: update preserves actuals and manual planned edits
