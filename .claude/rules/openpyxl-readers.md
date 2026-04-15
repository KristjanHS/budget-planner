---
paths:
  - "budget_planner/workbook.py"
  - "budget_planner/reader.py"
  - "budget_planner/config_sheet.py"
---
# openpyxl Reader Patterns

Section readers relying on blank col A to stop should also type-guard a known-numeric column: `if val is not None and not isinstance(val, (int, float)): break`. The blank separator row can be filled by user edits, test `_set` calls, or inserted rows — without the type guard the reader scans into the next section header. Mirrors `read_team_from_config`. When consolidating those guards into a shared parser (e.g. `_parse_config_float`), preserve break-on-invalid: bool → None (int subclass trap), empty/whitespace string → None (not 0.0), formula string `=...` → None. Only `None` (truly empty cell on a caller-validated row) should map to 0.0. Otherwise a named row with a blank numeric cell silently becomes a zero-capacity ghost record.
When adding/renaming Config param labels, check `_find_header_row` substring collisions — it uses case-insensitive `in` matching, so "PM/person/year" matches a "Person" search.
When scanning rows for data, check both None and empty string — openpyxl preserves `""` from `style_cell(ws, r, c, "")` as a non-None value.
