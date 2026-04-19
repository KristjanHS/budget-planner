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
Writing to a cell via `ws.cell(r, c).value = x` trips pyright on the `Cell | MergedCell` union. Use the keyword form `ws.cell(row=r, column=c, value=x)` instead — its return type excludes MergedCell.
Excel CF priority: first rule added gets lowest priority number = highest priority = wins conflicts. Add the most critical fill first (e.g., red before amber).
When CF formulas reference external data varying per row (e.g., different WP end months), apply per-row rules — a range-wide rule with absolute refs uses the same value for all rows.
SUMPRODUCT with mixed text/number arrays returns #VALUE!. `(range)*ISNUMBER(range)` also fails because `text*0` = #VALUE! in LibreOffice. Use individual `IF(ISNUMBER(cell),cell*other,0)` terms instead.
openpyxl retains previous protection state when overwriting a cell — if you must restyle a cell, explicitly set `cell.protection` again.
CF fills must use `PatternFill(bgColor="FFFFC7CE")` (8-char AARRGGBB hex, not 6-char — 6-char gets `00` alpha = transparent). Use `CF_*_FILL` constants from `styles.py` in all `FormulaRule` calls; regular `*_FILL` constants are for cell formatting only.
CF test assertions: match rules by fill color + formula pattern, not positional index — adding new CF rules shifts all indices.
CF rule-count churn is a design signal. When a single task's implementation bumps a CF-count assertion more than once, pause rather than re-fix the count — usually the task is duplicating a signal already rendered elsewhere on the sheet.
Prefer outline groups (`column_dimensions[L].outline_level` / `row_dimensions[r].outline_level`) over `hidden=True` for derived cells — keeps formula provenance inspectable.
To inspect CF rules from a test, use public iteration `for cf in ws.conditional_formatting: ws.conditional_formatting[cf]` (yields rule list). `ws.conditional_formatting._cf_rules` works at runtime but pyright flags it.
When filtering CF rules by cell-address prefix in a test, use `str(cf_range.sqref)` — bare `str(cf_range)` returns the `"<ConditionalFormatting A1:B3>"` wrapper form, not the sqref, and prefix matches silently return zero.
**`MATCH(1, <arr>, 0)` and `LOOKUP(2, 1/<arr>, <hdr>)` are silently broken under openpyxl+LO recalc.** Both need CSE array-evaluation; openpyxl writes regular-entry formulas (no CSE marker), LibreOffice returns `#N/A` (MATCH) or drifts to a false neighbour (LOOKUP), and any `IFERROR` wrapper swallows the signal (cells look "always empty" or "always last header"). Use SUMPRODUCT-forcing-array instead: first non-empty → `=IFERROR(INDEX(hdr, SUMPRODUCT(MIN((row<>"")*midx + (row="")*99999))), "")`; last non-empty → `=IFERROR(IF(SUMPRODUCT(MAX((row<>"")*midx))=0,"", INDEX(hdr, SUMPRODUCT(MAX((row<>"")*midx)))), "")`. SUMPRODUCT wrapping MIN/MAX is the documented-stable array-coercion pattern for this project. **The `IF(MAX=0,"",…)` guard on Last log is required** — LO's `INDEX(range, 0)` does NOT raise; it returns the full range and implicit-intersects at the formula cell's column, leaking a neighbour's header into the cell (IFERROR alone is insufficient). First log's `99999` sentinel dodges this by forcing INDEX out of range on all-empty rows. *2026-04-19: Timelog Plan v1 shipped with both broken cells; v2 Phase 1 (commit 7d8c399) rewrote them using the SUMPRODUCT(MIN/MAX(...)) pattern — grep `timelog_plan.py` for the reference form, not the v1 MATCH/LOOKUP variant.* Cross-ref: banned-pattern invariant in `docs/design-invariants.md`.
CF count assertions that target a *band* range on a column also shared by dashboard cells must filter out single-cell CF rules (e.g. require `":" in rng`). A `startswith("<col>")` filter alone over-counts when the sheet has a dashboard row on the same column.
SUMPRODUCT + range lookups (VLOOKUP, INDEX/MATCH, CHOOSE/MATCH with range args) fail in Excel from openpyxl files — CSE `t="array"` flag is recognized but doesn't force array evaluation on file open. Working patterns: SUMPRODUCT + equality tests per category `(range=$A$1)*$B$1`, per-cell lookups (single-cell VLOOKUP/INDEX), SUMIFS. LibreOffice handles all approaches correctly, so always verify in Excel.
`MATCH(TRUE, range<>"", 0)` and similar boolean-array MATCH formulas silently require CSE array entry in pre-365 Excel — they work in LO and Excel-365 dynamic arrays but return `#VALUE!` in older Excel as a regular formula. Use `MATCH(1, (range<>"")*1, 0)` (or `--(range<>"")` coercion) — works as a plain (non-CSE) entry everywhere.
LibreOffice 7.3 requires `_xlfn.` prefix for newer Excel functions (MINIFS, MAXIFS, IFS, SWITCH, TEXTJOIN, CONCAT). Without it: `#NAME?` in LO recalc. openpyxl doesn't add this automatically.
`ws.insert_rows()` shifts cells but does NOT update `DefinedName` targets — Excel auto-extends named ranges when inserting strictly inside `[row1, last]`. Tests for shape-drift scenarios must emulate this by rewriting `dn.attr_text` explicitly. Use the helper pattern in `tests/test_rebuild_shape_drift.py::_simulate_excel_insert_in_named_range`. Boundary is `[row1, last]` — inserts at `last + 1` are appends and leave the name untouched.
`DefinedName.attr_text` targeting a sheet whose name contains a hyphen, space, or other non-alphanumeric MUST single-quote the sheet name — e.g. `f"'{SHEET_ACTUALS_HOURS}'!$D$1"`, not `f"{SHEET_ACTUALS_HOURS}!$D$1"`. Unquoted, openpyxl lowercases the ref on save (`actuals-hours!$d$1`); Excel/LO then resolve that as `#NAME?` at recalc, cascading through every downstream consumer of the named range. Unit-level formula tests see correct text and pass — only recalc-level e2e tests (`test_*_recalc_zero_errors`, `test_hours_propagate_through_rebuild_and_recalc`) catch it. Matching test assertions store the quoted form verbatim.
When a cell holding a cross-sheet formula (e.g. `=Config!A5`) is used as the lookup value of VLOOKUP/MATCH in another formula, the lookup fails silently under `data_only=True` — the formula cell returns None. Write a literal value if the cell is consumed by lookups; reserve `=Config!...` formula refs for display-only cells.
EUR/money leaf formulas must wrap the math in `ROUND(..., 0)` when the output format is whole units (e.g. `€#,##0`). Fractional leaves compound across SUMPRODUCT/SUM aggregation into multi-unit mismatches against auditor totals that round at each step.
Cross-sheet row-map plumbing between builders (passing `wp_rows`/`roster_rows` into a consumer so it can build `=OtherSheet!<col><row>` formulas) is a smell. Prefer: add a helper column on the source sheet that pre-computes the aggregate, and let the consumer do a single-cell lookup. Collapses SUM+INDEX cross-sheet formulas to a direct ref and eliminates layout-map parameters on the two-pass build.
