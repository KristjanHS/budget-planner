---
paths:
  - "budget_planner/trends.py"
---
# openpyxl Chart Patterns

Chart series titles require SeriesLabel: `series.tx = SeriesLabel(v="Name")`, not `series.title = "Name"`.
LibreOffice `store()` strips chart series data (val/tx/cat XML) from openpyxl files. `recalc.py` preserves chart XML across LO recalc — don't bypass this.
For charts, construct Series manually with `NumDataSource`/`NumRef`/`AxDataSource` instead of `chart.add_data()` — produces explicit OOXML elements.
Charts referencing hidden rows must set `chart.visible_cells_only = False` — openpyxl defaults to `True` (`<c:plotVisOnly val="1"/>`), which makes Excel skip hidden-row data → blank chart.
When debugging chart issues, inspect the zip XML directly (`zipfile` + `ElementTree`) — openpyxl's Python API can show data that wasn't serialized.
When adding a chart to an existing sheet, grep tests for `len(ws._charts)` count assertions and `ws._charts[0]` index hardcoding — both snap-fail silently when a second chart appears.
