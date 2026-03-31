# CLAUDE.md — Project Instructions for Claude Code

## What This Is

An openpyxl-based Excel workbook generator for EU project FTE budget planning. The user will iterate on this with you.

## Key Commands

```bash
# Generate the workbook
python build_budget.py

# Recalculate formulas (MANDATORY after any generation)
python scripts/recalc.py output/budget_plan.xlsx

# Run full verification suite (recalc + 13 structural checks)
python verify.py output/budget_plan.xlsx

# Recalc output must show: "status": "success", "total_errors": 0
# Verify output must show: RESULT: PASS
```

## Critical Rules

1. **Always use Excel formulas, never hardcode calculated values in Python.** The spreadsheet must remain dynamic.
2. **After every generation, run recalc.py and verify zero errors** before presenting to user.
3. **Font color convention is sacred**: Blue = input, Black = formula, Green = cross-sheet ref. Never mix these.
4. **Inactive month cells must be greyed out and empty.** Auditors flag time logged outside WP active periods.
5. **Named ranges `project_start` and `fte_rate`** must always exist and point to Plan!$B$2 and Plan!$E$2.
6. **_META sheet must be updated** whenever you change the workbook structure. It's the primary documentation for future AI agents.

## Architecture Overview

- `config.py` — All project-specific data (WP list, constants, project duration). Edit this to change the project, not the builder.
- `build_budget.py` — Workbook generator. Three sheets: _META, Plan, Summary. The Plan sheet has three blocks: Config table, Planned FTE grid, Actual FTE grid.
- `SPEC.md` — Full specification with business context, design decisions, and known limitations.
- `CONVENTIONS.md` — Spreadsheet formatting and structural conventions.

## Pending Restructure

Approved architecture plan at `docs/plans/architecture-restructure.md` (v3, 2026-03-31). Migrates to `budget_planner/` package with pytest, reader for round-trip updates, and flat module structure. Not yet implemented.

## Common Tasks

### "Replace example WPs with real data"
→ Edit `EXAMPLE_WPS` in `config.py`. Run `build_budget.py` + recalc.

### "Add a new column/metric to Summary"
→ Modify the Summary section in `build_budget.py`. Update `_META` sheet text. Run + recalc.

### "Add per-person breakdown"
→ This is listed as future work in SPEC.md. Would require a new block or sheet. Discuss design with user before implementing.

### "Import actuals from timesheet"
→ Not yet implemented. Would be a separate script that reads CSV/XLSX and writes into the Actual block.

## Testing Approach

Run `python verify.py output/budget_plan.xlsx` after every change. It checks:

1. Formula recalculation (0 errors via LibreOffice)
2. Sheet existence (_META, Plan, Summary)
3. Global parameters (project start date, FTE rate)
4. Named ranges (project_start, fte_rate)
5. Config table structure and data validity
6. Planned FTE validation (sum matches budget per WP)
7. Actual FTE block (zeros in fresh workbook)
8. Inactive month cells (empty outside WP active periods)
9. Summary sheet WP count matches config
10. _META sheet has documentation content

Exit code 0 = PASS, 1 = FAIL. Warnings (e.g., rounding variance) don't cause failure.
