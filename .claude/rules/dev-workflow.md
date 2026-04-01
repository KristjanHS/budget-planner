---
paths: "**/*.py"
---
# Dev Workflow

## Build + Test
After code changes: `pytest tests/ -v`
LibreOffice is a required dependency — install it, don't skip recalc tests.
Use uv for venv and dependency management.
Run pytest once per step. Never re-run to "confirm".
Always verify via pytest, not ad-hoc `python3 -c` scripts.

## Update mode
When testing workbook updates, always inject realistic actuals
into the test fixture — don't just test fresh generation.

## Linting
Run pyright before committing. Fix all errors.
Subagents create dead code — verify after each task.

## Code Review
Suggest review before committing.
Move completed plans to docs/plans/old_already_implemented/.

## Architecture
Don't propose subpackages or abstractions for codebases under ~2K LOC.
Before splitting a module, verify each resulting file has >50 lines of logic.
Prefer flat packages over nested. Don't build generic iteration for N=2.

## Invariants
- Excel formulas only — never hardcode calculated values
- Font color: Blue=input, Black=formula, Green=cross-sheet
- _META sheet must reflect current workbook structure
- Block builders must accept preserved_data for round-trip safety
