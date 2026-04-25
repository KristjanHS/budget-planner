---
paths:
  - "budget_planner/workbook.py"
  - "budget_planner/reader.py"
  - "budget_planner/config_sheet.py"
  - "scripts/**/*.py"
---
# openpyxl Workbook / Reader / Named-range Infrastructure

`wb.active` can return None — `assert ws is not None` before use. Writing via `ws.cell(r, c).value = x` trips pyright on the `Cell | MergedCell` union; use `ws.cell(row=r, column=c, value=x)` (keyword form's return type excludes MergedCell). openpyxl retains previous protection state when overwriting — if you restyle, explicitly set `cell.protection` again.
`ws.insert_rows()` shifts cells but does NOT update `DefinedName` targets — Excel auto-extends named ranges when inserting strictly inside `[row1, last]`. Shape-drift tests must rewrite `dn.attr_text` explicitly; see `tests/test_rebuild_shape_drift.py::_simulate_excel_insert_in_named_range`. Inserts at `last + 1` are appends and leave the name untouched.
`DefinedName.attr_text` targeting a sheet whose name contains a hyphen, space, or other non-alphanumeric MUST single-quote the sheet name — `f"'{SHEET_ACTUALS_HOURS}'!$D$1"`, not `f"{SHEET_ACTUALS_HOURS}!$D$1"`. Unquoted, openpyxl lowercases the ref on save; Excel/LO resolve as `#NAME?` at recalc, cascading through every downstream consumer. Only recalc-level e2e tests catch it; matching test assertions store the quoted form verbatim.
Section readers relying on blank col A to stop should also type-guard a known-numeric column (`if val is not None and not isinstance(val, (int, float)): break`) — the blank separator row can be filled by user edits or inserted rows. When consolidating guards into a shared parser, preserve break-on-invalid: bool → None (int subclass trap), empty/whitespace string → None, formula string `=...` → None; only `None` on a caller-validated row maps to 0.0.
When adding/renaming Config param labels, check `_find_header_row` substring collisions — case-insensitive `in` matching means `PM/person/year` matches a `Person` search. When scanning rows for data, check both None and empty string — `style_cell(ws, r, c, "")` preserves `""` as non-None.
When inserting a new column into a structured section (WP BUDGET, Team Roster, etc.), every position-based reader must branch on layout detection AND shift `start_col/end_col/type_col`-style indices, not just the renamed cell — readers index by numeric column, not header. Detect legacy vs post-shift via header-string at the inserted position (NOT `version_stamp` — that's a git-describe literal). Full rationale + fixture pattern: `docs/spec.md` § Reader column-shift invariant.
