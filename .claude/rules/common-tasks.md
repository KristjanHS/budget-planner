---
paths: "**/*.py"
---
# Common Tasks

## "Replace example WPs with real data"
Edit the `tbl_WorkPackages` table on the Config sheet in Excel. Run `python -m budget_planner rebuild` + pytest.

## "Replace example team with real data"
Edit the Team sheet roster (name, max FTE) and WP share seeds (blue cells). Run `python -m budget_planner rebuild` + pytest.

## "Add a new work package"
Add a row to `tbl_WorkPackages` on the Config sheet in Excel. Run `python -m budget_planner rebuild` + pytest.

## "Add a team member"
Add a row to the Team sheet roster, then add share seeds for the new person. Run `python -m budget_planner rebuild` + pytest.

## "Add a new column/metric to Summary"
Modify `budget_planner/summary.py`. Update `budget_planner/meta.py` text. Run + pytest.

## "Adjust team member WP shares"
Edit the share matrix on the Team sheet (blue input cells). Run `python -m budget_planner rebuild` if the structure changed (e.g., new WP columns).

## "Import actuals from timesheet"
Not applicable — the Actuals sheet tracks EUR costs from the accountant, not FTE from timesheets.

## "Add actual costs"
Enter EUR values in the Actuals sheet blue input cells (one row per WP, monthly columns).

## "Compare budget scenarios"
Use `--snapshot` to save current state, make changes, compare. Use `--restore` to revert.
