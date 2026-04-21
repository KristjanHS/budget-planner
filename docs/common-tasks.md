# Common Tasks

## "Replace example WPs with real data"
Edit the `tbl_WorkPackages` table on the Config sheet in Excel. Run `python -m budget_planner rebuild` + pytest.

## "Replace example team with real data"
Edit the TEAM ROSTER section on the Config sheet (Person + Max FTE blue cells). Run `python -m budget_planner rebuild` + pytest.

## "Add a new work package"
Add a row to `tbl_WorkPackages` on the Config sheet in Excel. One row per WP — populate id/name/FTE/rate plus Start/End/Type (cols H/I/J, blue inputs; Type from `{flat, front-loaded, bell, back-loaded}`). Run `python -m budget_planner rebuild` + pytest.

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

## "Change the uniform vacation month"
Edit the `Vacation Month` blue input on Config — three-letter abbreviation (`"Jan"`…`"Dec"`). Run `python -m budget_planner rebuild`. The Pool formula on Timelog Plan PER-PERSON SUMMARY re-derives and the grid re-locks the new month's columns. Default is `"Jul"`.

## "Add a vacation exception for a specific person"
**Not supported in the current model.** Vacation is uniform across the team via the `vacation_month` scalar. Per-(person, year) overrides are on the roadmap as Option A (see `docs/design-invariants.md` § Timelog Plan — auditor-side invariants) — a future Config `Vacation Schedule` table keyed by person × year will layer on additively without moving Timelog Plan cells.
