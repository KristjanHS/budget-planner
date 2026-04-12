---
# Add new sheet builders here when created
paths:
  - "budget_planner/plan.py"
  - "budget_planner/summary.py"
  - "budget_planner/capacity.py"
  - "budget_planner/actuals.py"
  - "budget_planner/trends.py"
  - "budget_planner/config_sheet.py"
  - "budget_planner/styles.py"
  - "budget_planner/workbook.py"
---
# openpyxl Patterns

`wb.active` can return None — always `assert ws is not None` before use.
Excel CF priority: first rule added gets lowest priority number = highest priority = wins conflicts. Add the most critical fill first (e.g., red before amber).
Chart series titles require SeriesLabel: `series.tx = SeriesLabel(v="Name")`, not `series.title = "Name"`.
When CF formulas reference external data varying per row (e.g., different WP end months), apply per-row rules — a range-wide rule with absolute refs uses the same value for all rows.
SUMPRODUCT with mixed text/number arrays returns #VALUE!. `(range)*ISNUMBER(range)` also fails because `text*0` = #VALUE! in LibreOffice. Use individual `IF(ISNUMBER(cell),cell*other,0)` terms instead.
openpyxl retains previous protection state when overwriting a cell — if you must restyle a cell, explicitly set `cell.protection` again.
CF fills must use `PatternFill(bgColor="FFFFC7CE")` (8-char AARRGGBB hex, not 6-char — 6-char gets `00` alpha = transparent). Use `CF_*_FILL` constants from `styles.py` in all `FormulaRule` calls; regular `*_FILL` constants are for cell formatting only.
CF test assertions: match rules by fill color + formula pattern, not positional index — adding new CF rules shifts all indices.
When scanning rows for data, check both None and empty string — openpyxl preserves `""` from `style_cell(ws, r, c, "")` as a non-None value.
Section readers relying on blank col A to stop should also type-guard a known-numeric column: `if val is not None and not isinstance(val, (int, float)): break`. The blank separator row can be filled by user edits, test `_set` calls, or inserted rows — without the type guard the reader scans into the next section header. Mirrors `read_team_from_config`.
LibreOffice `store()` strips chart series data (val/tx/cat XML) from openpyxl files. `recalc.py` preserves chart XML across LO recalc — don't bypass this.
For charts, construct Series manually with `NumDataSource`/`NumRef`/`AxDataSource` instead of `chart.add_data()` — produces explicit OOXML elements.
Charts referencing hidden rows must set `chart.visible_cells_only = False` — openpyxl defaults to `True` (`<c:plotVisOnly val="1"/>`), which makes Excel skip hidden-row data → blank chart.
When debugging chart issues, inspect the zip XML directly (`zipfile` + `ElementTree`) — openpyxl's Python API can show data that wasn't serialized.
SUMPRODUCT + range lookups (VLOOKUP, INDEX/MATCH, CHOOSE/MATCH with range args) fail in Excel from openpyxl files — CSE `t="array"` flag is recognized but doesn't force array evaluation on file open. Working patterns: SUMPRODUCT + equality tests per category `(range=$A$1)*$B$1`, per-cell lookups (single-cell VLOOKUP/INDEX), SUMIFS. LibreOffice handles all approaches correctly, so always verify in Excel.
When adding/renaming Config param labels, check `_find_header_row` substring collisions — it uses case-insensitive `in` matching, so "PM/person/year" matches a "Person" search.
LibreOffice 7.3 requires `_xlfn.` prefix for newer Excel functions (MINIFS, MAXIFS, IFS, SWITCH, TEXTJOIN, CONCAT). Without it: `#NAME?` in LO recalc. openpyxl doesn't add this automatically.
