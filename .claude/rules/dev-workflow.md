---
paths: "**/*.py"
---
# Dev Workflow

## Build + Test
LibreOffice is a required dependency — install it, don't skip recalc tests. See `docs/cloud-env-setup.md` for cloud/container environments.
When generating output workbooks, check what already exists — only generate missing or explicitly requested files. Don't overwrite existing workbooks unnecessarily.
For Excel-specific bugs (implicit intersection, formula eval): generate file → user verifies in Excel → then run pytest. openpyxl/LibreOffice mask Excel-specific issues.
When tests return None from `data_only=True`: verify recalc actually ran (`<v>` elements in XML), don't assume openpyxl/LO interop failure. No `<v>` = no calc engine ran.

## Linting
Run `pyright` (full project, not just changed files) before committing. Fix all errors.
Subagents create dead code — verify after each task.

## Code Review
Before executing a plan, move completed/superseded plans to docs/plans/old_already_implemented/.
When a plan claims a module is "unchanged", grep its imports to verify no coupling to changed modules.
When inserting/removing a Config column, grep ALL formula references (`Config!{letter}`) across every builder AND test file. Column shifts cascade silently — a test asserting the old column letter will mask the bug.
When inserting/removing a Config param row, grep ALL test files for hardcoded row numbers (e.g. `cell(4, 2)`, `wp_first_row = 8`, merge span tuples). Row shifts cascade into tests the plan doesn't list.
When changing error message format in reader/builders, grep test files for the old message pattern (`pytest.raises(match=` strings). Semantic changes in validation produce silent test failures.
When changing a Config cell from input (blue) to formula (black), grep all consumers of the corresponding reader field (e.g. `params["rate"]`) — validators, builders, demo seeder, and tests may assume a numeric type.

## Planning & Implementation (/impag)
Write plans to `docs/plans/` as dated .md files immediately — never keep plans only in conversation context (context gets compressed).
Run brainstorm → write-plan → worktree isolation for each P-item. Don't skip planning even if the spec looks detailed.
Run `superpowers:code-reviewer` after every /impag stage — not just at the end. Each stage boundary is a review checkpoint.
Fix pre-existing test failures when encountered during implementation — don't defer them.
Commit or resolve dirty files at session start before executing any tasks.

## Worktrees
Before starting work in a worktree:
1. `git fetch origin main`
2. Check divergence: `git log --oneline HEAD..origin/main | wc -l`
3. If >0 commits behind, `git merge origin/main` before any edits
4. If >5 commits behind, grep main for the planned change before implementing — another instance may have already done it

## Formatting
Preserve all existing cell formatting (fonts, fills, table styles, column widths) unless the user explicitly asks for a change.

## openpyxl Patterns
`wb.active` can return None — always `assert ws is not None` before use.
Excel CF priority: first rule added gets lowest priority number = highest priority = wins conflicts. Add the most critical fill first (e.g., red before amber).
When a design specifies Excel formulas, verify assumptions against a generated workbook (e.g. `python3 -c` probes) before committing the design. Check: cell value types (text vs number), whether referenced cells exist, whether values are static or cumulative.
Chart series titles require SeriesLabel: `series.tx = SeriesLabel(v="Name")`, not `series.title = "Name"`.
When CF formulas reference external data varying per row (e.g., different WP end months), apply per-row rules — a range-wide rule with absolute refs uses the same value for all rows.
SUMPRODUCT with mixed text/number arrays returns #VALUE!. `(range)*ISNUMBER(range)` also fails because `text*0` = #VALUE! in LibreOffice. Use individual `IF(ISNUMBER(cell),cell*other,0)` terms instead.
openpyxl retains previous protection state when overwriting a cell — if you must restyle a cell, explicitly set `cell.protection` again.
CF fills must use `PatternFill(bgColor="FFFFC7CE")` (8-char AARRGGBB hex, not 6-char — 6-char gets `00` alpha = transparent). Use `CF_*_FILL` constants from `styles.py` in all `FormulaRule` calls; regular `*_FILL` constants are for cell formatting only.
CF test assertions: match rules by fill color + formula pattern, not positional index — adding new CF rules shifts all indices.
When scanning rows for data, check both None and empty string — openpyxl preserves `""` from `style_cell(ws, r, c, "")` as a non-None value.
LibreOffice `store()` strips chart series data (val/tx/cat XML) from openpyxl files. `recalc.py` preserves chart XML across LO recalc — don't bypass this.
For charts, construct Series manually with `NumDataSource`/`NumRef`/`AxDataSource` instead of `chart.add_data()` — produces explicit OOXML elements.
Charts referencing hidden rows must set `chart.visible_cells_only = False` — openpyxl defaults to `True` (`<c:plotVisOnly val="1"/>`), which makes Excel skip hidden-row data → blank chart.
When debugging chart issues, inspect the zip XML directly (`zipfile` + `ElementTree`) — openpyxl's Python API can show data that wasn't serialized. Also create minimal isolated test files varying one factor at a time (bisect method) to pinpoint the breaking element.
SUMPRODUCT + range lookups (VLOOKUP, INDEX/MATCH, CHOOSE/MATCH with range args) fail in Excel from openpyxl files — CSE `t="array"` flag is recognized but doesn't force array evaluation on file open. Working patterns: SUMPRODUCT + equality tests per category `(range=$A$1)*$B$1`, per-cell lookups (single-cell VLOOKUP/INDEX), SUMIFS. LibreOffice handles all approaches correctly, so always verify in Excel.
When adding/renaming Config param labels, check `_find_header_row` substring collisions — it uses case-insensitive `in` matching, so "PM/person/year" matches a "Person" search.

## Testing
Tests must not hardcode config.py seed values (WP names, person names, FTE amounts, start/end months).
Derive expected values from config imports or from the generated workbook. Tests must pass after user changes data to real values in Excel and runs rebuild.

## Invariants (extends CLAUDE.md Critical Rules)
- When updating meta.py, verify text against the actual builder (e.g., count summary columns in summary.py, not the plan)
- On rebuild, all config data comes from the existing workbook (Config sheet + Capacity config block), never from config.py. Builders receive config as function parameters.
