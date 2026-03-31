---
paths: "**/*.py"
---
# Dev Workflow

## Build + Test
After code changes: `pytest tests/ -v`
Recalc tests need LibreOffice — auto-skip if unavailable.
Run pytest once per step. Never re-run to "confirm".

## Update mode
When testing workbook updates, always inject realistic actuals
into the test fixture — don't just test fresh generation.

## Linting
Run pyright before committing. Fix all errors.
Subagents create dead code — verify after each task.

## Code Review
Suggest review before committing.
Move completed plans to docs/plans/old_already_implemented/.

## Invariants
- Excel formulas only — never hardcode calculated values
- Font color: Blue=input, Black=formula, Green=cross-sheet
- _META sheet must reflect current workbook structure
- Block builders must accept preserved_data for round-trip safety
