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

# Run full test suite
pytest tests/ -v
```

## Critical Rules

1. **Always use Excel formulas, never hardcode calculated values in Python.** The spreadsheet must remain dynamic.
2. **After every generation, run `pytest tests/ -v`** and verify all pass before presenting to user.
3. **Font color convention is sacred**: Blue = input, Black = formula, Green = cross-sheet ref. Never mix these.
4. **Inactive month cells must be greyed out and empty.** Auditors flag time logged outside WP active periods.
5. **Named ranges `project_start` and `fte_rate`** must always exist and point to Plan!$B$2 and Plan!$E$2.
6. **_META sheet must be updated** whenever you change the workbook structure. It's the primary documentation for future AI agents.
