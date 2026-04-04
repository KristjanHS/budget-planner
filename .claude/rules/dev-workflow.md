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
When inserting/removing a Config column, grep ALL formula references (`Config!{letter}`) across every builder AND test file. Column shifts cascade silently — a test asserting the old column letter will mask the bug.

## Worktrees
Before starting work in a worktree:
1. `git fetch origin main`
2. Check divergence: `git log --oneline HEAD..origin/main | wc -l`
3. If >0 commits behind, `git merge origin/main` before any edits
4. If >5 commits behind, grep main for the planned change before implementing — another instance may have already done it

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
SUMPRODUCT with mixed text/number arrays returns #VALUE!. `(range)*ISNUMBER(range)` also fails because `text*0` = #VALUE! in LibreOffice. Use individual `IF(ISNUMBER(cell),cell*other,0)` terms instead.
openpyxl retains previous protection state when overwriting a cell — if you must restyle a cell, explicitly set `cell.protection` again.
CF test assertions: match rules by fill color + formula pattern, not positional index — adding new CF rules shifts all indices.
When scanning rows for data, check both None and empty string — openpyxl preserves `""` from `style_cell(ws, r, c, "")` as a non-None value.
LibreOffice `store()` strips chart series data (val/tx/cat XML) from openpyxl files. `recalc.py` preserves chart XML across LO recalc — don't bypass this.
For charts, construct Series manually with `NumDataSource`/`NumRef`/`AxDataSource` instead of `chart.add_data()` — produces explicit OOXML elements.
When debugging chart issues, inspect the zip XML directly (`zipfile` + `ElementTree`) — openpyxl's Python API can show data that wasn't serialized.

## Testing
Tests must not hardcode config.py seed values (WP names, person names, FTE amounts, start/end months).
Derive expected values from config imports or from the generated workbook. Tests must pass after user changes data to real values in Excel and runs rebuild.

## Invariants
- Excel formulas only — never hardcode calculated values
- Font color: Blue=input, Black=formula, Green=cross-sheet
- _META sheet must reflect current workbook structure
- When updating meta.py, verify text against the actual builder (e.g., count summary columns in summary.py, not the plan)
- On rebuild, all config data comes from the existing workbook (Config sheet + Capacity config block), never from config.py. Builders receive config as function parameters.
