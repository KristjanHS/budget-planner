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
When a plan claims a module is "unchanged", grep its imports to verify no coupling to changed modules.

## Worktrees
Before starting work in a worktree, verify it's up to date with main: `git log --oneline -1 HEAD` vs `git log --oneline -1 main`. Merge if behind.

## Architecture
Don't propose subpackages or abstractions for codebases under ~2K LOC.
Before splitting a module, verify each resulting file has >50 lines of logic.
Prefer flat packages over nested. Don't build generic iteration for N=2.

## openpyxl Patterns
`wb.active` can return None — always `assert ws is not None` before use.
Later conditional formatting rules override earlier ones — add lower-priority fills first (e.g., amber before red).
When a design specifies Excel formulas, verify assumptions against a generated workbook (e.g. `python3 -c` probes) before committing the design. Check: cell value types (text vs number), whether referenced cells exist, whether values are static or cumulative.
Chart series titles require SeriesLabel: `series.tx = SeriesLabel(v="Name")`, not `series.title = "Name"`.
When CF formulas reference external data varying per row (e.g., different WP end months), apply per-row rules — a range-wide rule with absolute refs uses the same value for all rows.
SUMPRODUCT with mixed text/number arrays returns #VALUE!. Guard with `IF(ISNUMBER(range),range,0)` when cells may contain text like "—".
When overwriting a styled cell (e.g., replacing BLUE_INPUT seed with BLACK_FORMULA), explicitly set `cell.protection` on the second pass — openpyxl retains the previous protection state.
CF test assertions: match rules by fill color + formula pattern, not positional index — adding new CF rules shifts all indices.

## Testing
Tests must not hardcode config.py seed values (WP names, person names, FTE amounts, start/end months).
Derive expected values from config imports or from the generated workbook. Tests must pass after user changes data to real values in Excel and runs rebuild.

## Invariants
- Excel formulas only — never hardcode calculated values
- Font color: Blue=input, Black=formula, Green=cross-sheet
- _META sheet must reflect current workbook structure
- When updating meta.py, verify text against the actual builder (e.g., count summary columns in summary.py, not the plan)
- On rebuild, all config data comes from the existing workbook (Config sheet + Capacity config block), never from config.py. Builders receive config as function parameters.
