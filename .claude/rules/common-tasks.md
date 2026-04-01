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
Not yet implemented. Would be a separate script that reads CSV/XLSX and uses `generate_workbook(--update)`.

## "Compare budget scenarios"
Use `--snapshot` to save current state, make changes, compare. Use `--restore` to revert.
