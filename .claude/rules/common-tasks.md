---
paths: "**/*.py"
---
# Common Tasks

## "Replace example WPs with real data"
Edit the `tbl_WorkPackages` table on the Config sheet in Excel. Run `python -m budget_planner rebuild` + pytest.

## "Replace example team with real data"
Edit the TEAM ROSTER section on the Config sheet (Person + Max FTE blue cells). Run `python -m budget_planner rebuild` + pytest.

## "Add a new work package"
Add a row to `tbl_WorkPackages` on the Config sheet in Excel. Run `python -m budget_planner rebuild` + pytest.

## "Add a team member"
Add a row to the TEAM ROSTER section on the Config sheet. Run `python -m budget_planner rebuild` + pytest.

## "Add a new column/metric to Summary"
Modify `budget_planner/summary.py`. Update `budget_planner/meta.py` text. Run + pytest.

## "Import actuals from timesheet"
Not applicable — the Actuals sheet tracks EUR costs from the accountant, not FTE from timesheets.

## "Add actual costs"
Enter EUR values in the Actuals sheet blue input cells (one row per WP, monthly columns).

## "Verify workbook after generation"
Use `openpyxl-read` MCP to spot-check: `get_sheet_names` (all 6 sheets present), `get_cell_details` on named ranges (`project_start`, `fte_rate`), `get_cell_range` on a few formula cells to confirm structure. Then run pytest for full validation.

## "Compare budget scenarios"
Use `--snapshot` to save current state, make changes, compare. Use `--restore` to revert.
