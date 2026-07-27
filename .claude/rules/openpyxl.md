---
# Add new sheet builders here when created
paths:
  - "budget_planner/plan.py"
  - "budget_planner/summary.py"
  - "budget_planner/capacity.py"
  - "budget_planner/actuals.py"
  - "budget_planner/actuals_hours.py"
  - "budget_planner/trends.py"
  - "budget_planner/config_sheet.py"
  - "budget_planner/workbook.py"
---
# openpyxl Patterns

Infrastructure (workbook, reader, named ranges, section readers) → `openpyxl-infra.md`.

## Core patterns (all builders)
Excel CF priority: first rule added = highest priority = wins conflicts. Add the most critical fill first (red before amber). For CF formulas referencing per-row external data, apply per-row rules — range-wide rules with absolute refs use the same value for every row.
CF fills must use `PatternFill(bgColor="FFFFC7CE")` (8-char AARRGGBB — 6-char gets `00` alpha = transparent). Use `CF_*_FILL` constants from `styles.py` in all `FormulaRule` calls; regular `*_FILL` constants are cell-formatting only.
CF test assertions: match rules by fill color + formula pattern, not positional index — adding new rules shifts all indices. CF count assertions targeting a *band* range on a column shared by dashboard cells must filter out single-cell rules (e.g. require `":" in rng`); `startswith("<col>")` alone over-counts. When a task's implementation bumps a CF-count assertion more than once, pause — usually the task is duplicating a signal already rendered elsewhere.
To inspect CF rules from a test, use `for cf in ws.conditional_formatting: ws.conditional_formatting[cf]` (public iteration — `_cf_rules` works at runtime but pyright flags it). When filtering by cell-address prefix, use `str(cf_range.sqref)` — bare `str(cf_range)` returns the `<ConditionalFormatting A1:B3>` wrapper form.
**`MATCH(1, <arr>, 0)` and `LOOKUP(2, 1/<arr>, <hdr>)` are silently broken under openpyxl+LO recalc** — both need CSE array-evaluation. Use SUMPRODUCT-forcing-array form (reference: `timelog_plan.py`); banned-pattern invariant + full first/last formulas + `IF(MAX=0,"",…)` guard rationale in `docs/spec.md` § Banned formula patterns.
SUMPRODUCT + mixed text/number arrays return `#VALUE!`; `(range)*ISNUMBER(range)` also fails (`text*0` = `#VALUE!`). Use individual `IF(ISNUMBER(cell),cell*other,0)` terms. SUMPRODUCT + range lookups (VLOOKUP, INDEX/MATCH, CHOOSE/MATCH with range args) fail in Excel from openpyxl files (the CSE `t="array"` flag doesn't force array evaluation on file open); use SUMPRODUCT + equality per category `(range=$A$1)*$B$1`, per-cell lookups, or SUMIFS. LibreOffice handles all approaches; always verify in Excel.
**Never pass a multi-cell range to `ISNUMBER` in an array scan** — Excel gives `ISNUMBER(range)` no array context outside CSE (the `SUMPRODUCT` wrapper does not rescue it, unlike `COLUMN(range)`, which is fine), so the scan returns 0; LO array-evaluates it correctly, so recalc tests stay green on a broken workbook. Use `(range<>"")`; scalar `ISNUMBER(cell)` is unaffected. Static guard: `tests/test_formula_array_context_guard.py`; rationale in `docs/spec.md` § Banned: `ISNUMBER(range)`.
`MATCH(TRUE, range<>"", 0)` and boolean-array MATCH silently require CSE array entry in pre-365 Excel. Use `MATCH(1, (range<>"")*1, 0)` or `--(range<>"")` — works as a plain entry everywhere.
LibreOffice 7.3 requires `_xlfn.` prefix for newer Excel functions (MINIFS, MAXIFS, IFS, SWITCH, TEXTJOIN, CONCAT). Without it: `#NAME?` in LO recalc. openpyxl doesn't add this automatically.
EUR/money leaf formulas must wrap math in `ROUND(..., 0)` when the output format is whole units (e.g. `€#,##0`). Fractional leaves compound across SUMPRODUCT/SUM into multi-unit mismatches against auditor totals that round at each step.
When a cell holding a cross-sheet formula (`=Config!A5`) is used as the lookup value of VLOOKUP/MATCH in another formula, the lookup fails silently under `data_only=True` — the formula cell returns `None`. Write a literal value if the cell is consumed by lookups; reserve `=Config!...` refs for display-only cells.
`layout.cell_ref(sheet, row, col)` defaults to **absolute** refs (`Sheet!$B$3`). Tests asserting verbatim formula text emitted through `cell_ref(...)` must include the `$` markers.
**Rule 3 (font color, blue/black/green)**: when a doc row claims a specific font color or formula-vs-input style for a named cell, grep the builder's actual `style_cell(...)` / `font=...` call to verify. Rule 3 is load-bearing; mislabelling is worse than silent. Don't generalize from sibling cells.
Rule 3 carve-out: builder-written audit metadata (e.g. `version_stamp` on Config) uses BLACK_FORMULA even though the cell value is a literal string — signals "do-not-edit by user." Required mitigation when deviating: builder-code comment AND in-cell comment explicitly naming the deviation (see `config_sheet.py` Version Stamp rows).
Prefer outline groups (`column_dimensions[L].outline_level` / `row_dimensions[r].outline_level`) over `hidden=True` for derived cells — keeps formula provenance inspectable.

## Charts (trends.py only)
Chart series titles require `SeriesLabel`: `series.tx = SeriesLabel(v="Name")`, not `series.title = "Name"`.
LibreOffice `store()` strips chart series data (val/tx/cat XML) from openpyxl files. `recalc.py` preserves chart XML across LO recalc — don't bypass this.
Construct Series manually with `NumDataSource`/`NumRef`/`AxDataSource` instead of `chart.add_data()` — produces explicit OOXML elements.
Charts referencing hidden rows must set `chart.visible_cells_only = False` — openpyxl defaults to `True`, making Excel skip hidden-row data → blank chart.
When debugging chart issues, inspect the zip XML directly (`zipfile` + `ElementTree`) — openpyxl's Python API can show data that wasn't serialized.
When adding a chart to an existing sheet, grep tests for `len(ws._charts)` count assertions and `ws._charts[0]` index hardcoding — both snap-fail silently when a second chart appears.
