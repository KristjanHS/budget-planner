---
paths: "**/*.py"
---
# Common Tasks

## "Replace example WPs with real data"
Edit `WORK_PACKAGES` in `budget_planner/config.py`. Run `python -m budget_planner` + pytest.

## "Replace example team with real data"
Edit `TEAM` and `EXTERNAL_COMMITMENTS` in `budget_planner/config.py`. Run `python -m budget_planner` + pytest.

## "Add a new column/metric to Summary"
Modify `budget_planner/summary.py`. Update `budget_planner/meta.py` text. Run + pytest.

## "Adjust team member WP shares"
Edit in the Team sheet directly (blue input cells), or update `TEAM` in `config.py` and regenerate.

## "Import actuals from timesheet"
Not applicable — the Actuals sheet tracks EUR costs from the accountant, not FTE from timesheets.

## "Add actual costs"
Enter EUR values in the Actuals sheet blue input cells (one row per WP, monthly columns).

## "Compare budget scenarios"
Use `--snapshot` to save current state, make changes, compare. Use `--restore` to revert.
