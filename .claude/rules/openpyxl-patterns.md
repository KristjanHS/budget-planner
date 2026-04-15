---
# Add new sheet builders here when created
paths:
  - "budget_planner/plan.py"
  - "budget_planner/summary.py"
  - "budget_planner/capacity.py"
  - "budget_planner/actuals.py"
  - "budget_planner/actuals_hours.py"
  - "budget_planner/plan_forecast.py"
  - "budget_planner/trends.py"
  - "budget_planner/config_sheet.py"
  - "budget_planner/styles.py"
  - "budget_planner/workbook.py"
---
# openpyxl Patterns (core)

Chart-specific tips → `openpyxl-charts.md` (trends.py only). Reader/header-row tips → `openpyxl-readers.md` (workbook.py, reader.py, config_sheet.py).

`wb.active` can return None — always `assert ws is not None` before use.
Excel CF priority: first rule added gets lowest priority number = highest priority = wins conflicts. Add the most critical fill first (e.g., red before amber).
When CF formulas reference external data varying per row (e.g., different WP end months), apply per-row rules — a range-wide rule with absolute refs uses the same value for all rows.
SUMPRODUCT with mixed text/number arrays returns #VALUE!. `(range)*ISNUMBER(range)` also fails because `text*0` = #VALUE! in LibreOffice. Use individual `IF(ISNUMBER(cell),cell*other,0)` terms instead.
openpyxl retains previous protection state when overwriting a cell — if you must restyle a cell, explicitly set `cell.protection` again.
CF fills must use `PatternFill(bgColor="FFFFC7CE")` (8-char AARRGGBB hex, not 6-char — 6-char gets `00` alpha = transparent). Use `CF_*_FILL` constants from `styles.py` in all `FormulaRule` calls; regular `*_FILL` constants are for cell formatting only.
CF test assertions: match rules by fill color + formula pattern, not positional index — adding new CF rules shifts all indices.
CF rule-count churn is a design signal. When a single task's implementation bumps a CF-count assertion more than once, pause rather than re-fix the count — usually the task is duplicating a signal already rendered elsewhere on the sheet.
Prefer outline groups (`column_dimensions[L].outline_level` / `row_dimensions[r].outline_level`) over `hidden=True` for derived cells — keeps formula provenance inspectable.
To inspect CF rules from a test, use public iteration `for cf in ws.conditional_formatting: ws.conditional_formatting[cf]` (yields rule list). `ws.conditional_formatting._cf_rules` works at runtime but pyright flags it.
SUMPRODUCT + range lookups (VLOOKUP, INDEX/MATCH, CHOOSE/MATCH with range args) fail in Excel from openpyxl files — CSE `t="array"` flag is recognized but doesn't force array evaluation on file open. Working patterns: SUMPRODUCT + equality tests per category `(range=$A$1)*$B$1`, per-cell lookups (single-cell VLOOKUP/INDEX), SUMIFS. LibreOffice handles all approaches correctly, so always verify in Excel.
LibreOffice 7.3 requires `_xlfn.` prefix for newer Excel functions (MINIFS, MAXIFS, IFS, SWITCH, TEXTJOIN, CONCAT). Without it: `#NAME?` in LO recalc. openpyxl doesn't add this automatically.
`ws.insert_rows()` shifts cells but does NOT update `DefinedName` targets — Excel auto-extends named ranges when inserting strictly inside `[row1, last]`. Tests for shape-drift scenarios must emulate this by rewriting `dn.attr_text` explicitly. Use the helper pattern in `tests/test_rebuild_shape_drift.py::_simulate_excel_insert_in_named_range`. Boundary is `[row1, last]` — inserts at `last + 1` are appends and leave the name untouched.
When a cell holding a cross-sheet formula (e.g. `=Config!A5`) is used as the lookup value of VLOOKUP/MATCH in another formula, the lookup fails silently under `data_only=True` — the formula cell returns None. Write a literal value if the cell is consumed by lookups; reserve `=Config!...` formula refs for display-only cells.
EUR/money leaf formulas must wrap the math in `ROUND(..., 0)` when the output format is whole units (e.g. `€#,##0`). Fractional leaves compound across SUMPRODUCT/SUM aggregation into multi-unit mismatches against auditor totals that round at each step.
Cross-sheet row-map plumbing between builders (passing `wp_rows`/`roster_rows` into a consumer so it can build `=OtherSheet!<col><row>` formulas) is a smell. Prefer: add a helper column on the source sheet that pre-computes the aggregate, and let the consumer do a single-cell lookup. Collapses SUM+INDEX cross-sheet formulas to a direct ref and eliminates layout-map parameters on the two-pass build.
