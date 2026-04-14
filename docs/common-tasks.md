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
Modify `budget_planner/summary.py`. If the change introduces a new cross-sheet invariant or trust boundary, also update `docs/design-invariants.md`; if it shifts row/column layout, reflect that in `budget_planner/layout.py` + `docs/internals.md`. Run + pytest.

## "Import actuals from timesheet"
Not applicable — the Actuals sheet tracks EUR costs from the accountant, not FTE from timesheets.

## "Add actual costs"
Enter EUR values in the Actuals sheet blue input cells (one row per WP, monthly columns).

## "Verify workbook after generation"
Run pytest — it covers structural validation (sheets, named ranges, formulas, fonts). Use `openpyxl-read` MCP only for targeted follow-up questions (e.g., `get_cell_details` to check a specific formula or font color, `get_list_of_sheets` to confirm dimensions). Don't MCP-scan the whole workbook as a routine step.

## "Compare budget scenarios"
Use `--snapshot` to save current state, make changes, compare. Use `--restore` to revert.
