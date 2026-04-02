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

## Linting
Run `pyright` (full project, not just changed files) before committing. Fix all errors.
Subagents create dead code — verify after each task.

## Code Review
Suggest review before committing.
Before executing a plan, move completed/superseded plans to docs/plans/old_already_implemented/.

## Worktrees
Before starting work in a worktree, verify it's up to date with main: `git log --oneline -1 HEAD` vs `git log --oneline -1 main`. Merge if behind.

## Architecture
Don't propose subpackages or abstractions for codebases under ~2K LOC.
Before splitting a module, verify each resulting file has >50 lines of logic.
Prefer flat packages over nested. Don't build generic iteration for N=2.

## openpyxl Patterns
`wb.active` can return None — always `assert ws is not None` before use.
Later conditional formatting rules override earlier ones — add lower-priority fills first (e.g., amber before red).

## Invariants
- Excel formulas only — never hardcode calculated values
- Font color: Blue=input, Black=formula, Green=cross-sheet
- _META sheet must reflect current workbook structure
- When updating meta.py, verify text against the actual builder (e.g., count summary columns in summary.py, not the plan)
- Generation is config-only — no round-trip. Do not add preserved_data parameters to builders.
