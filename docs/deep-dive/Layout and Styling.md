# Deep Dive: Layout and Styling

## Overview

`layout.py` and `styles.py` are the shared infrastructure that every builder imports. They enforce consistency across all 7 sheets and centralize constants that would otherwise be scattered and duplicated.

## layout.py — Column Structure

### Month Column Positioning

All sheets with monthly data use the same column structure:

```
Col A (1) = WP ID / Person name
Col B (2) = WP Name / label
Col C (3) = Month 1        <-- MONTH_COL_START = 3
Col D (4) = Month 2
...
Col N     = Month N         <-- month_col_end(N)
Col N+1   = Sum             <-- val_columns(N)["sum"]
Col N+2   = Budget          <-- val_columns(N)["budget"]
Col N+3   = Variance        <-- val_columns(N)["diff"]
Col N+4   = Cum. EUR        <-- val_columns(N)["cum"]      (Actuals only)
Col N+5   = % Ceiling       <-- val_columns(N)["pct_ceil"] (Actuals only)
Col N+6   = Breach In       <-- val_columns(N)["breach"]   (Actuals only)
```

Validation columns are defined as offsets from the last month column, so changing project duration automatically repositions them.

### Calendar Year Grouping

`_calendar_year_groups()` yields `(start_month_index, span)` tuples for each calendar year. For a project starting Nov 2026 with 14 months:

| Year | Months | Indices | Span |
|------|--------|---------|------|
| 2026 | M1-M2  | 0-1     | 2    |
| 2027 | M3-M14 | 2-13    | 12   |

### build_year_month_headers()

A shared function that writes the two-row year/month header used by Plan, Capacity, and Actuals:

```
Row N:   | [blank] | [blank] | 2026 (merged) | 2027 (merged) | ... | Sum | Budget | Var |
Row N+1: | WP ID   | WP Name | M1 Nov 26     | M2 Dec 26     | ... |     |        |     |
```

### Alert Thresholds

Centralized constants used by conditional formatting rules:

| Constant | Value | Used By |
|----------|-------|---------|
| `TREND_OK_BAND` | 0.05 | Trends burn delta |
| `TREND_OVER_THRESHOLD` | 0.15 | Trends burn delta |
| `TREND_SLOW_THRESHOLD` | -0.10 | Trends burn delta |
| `CAPACITY_AMBER_LOAD` | 0.80 | Capacity + Trends person load |
| `ACTUALS_AMBER_PCT_CEIL` | 0.85 | Actuals % ceiling |
| `ACTUALS_AMBER_BREACH_MONTHS` | 12 | Actuals breach column |
| `ACTUALS_RED_BREACH_MONTHS` | 6 | Actuals breach column |

## styles.py — Visual Constants

### Font Color Convention

```python
BLUE_INPUT    = Font(color="0000FF")  # User edits these cells
BLACK_FORMULA = Font(color="000000")  # Formula cells (locked)
GREEN_XREF    = Font(color="008000")  # Cross-sheet reference formulas
```

### Fill Colors

```python
GREY_FILL        = "D9D9D9"  # Inactive months, protected areas
LIGHT_BLUE_FILL  = "D6E4F0"  # Header rows
LIGHT_GREEN_FILL = "E2EFDA"  # Total rows
SECTION_FILL     = "2F5496"  # Section headers (dark blue, white text)
YELLOW_FILL      = "FFFF00"  # Attention needed
AMBER_FILL       = "FFD966"  # Approaching limit
RED_FILL         = "FFC7CE"  # Error / overrun
```

### style_cell() Helper

The workhorse function used everywhere:

```python
def style_cell(ws, row, col, value=None, font=None, fill=None,
               alignment=None, fmt=None, border=True, unlocked=False):
```

**Auto-unlock feature:** When font color is blue (0000FF), the cell is automatically set to `Protection(locked=False)`. This means when sheet protection is re-enabled, all blue input cells are editable without any extra code.

### Column Grouping

`apply_year_groups()` creates collapsible column groups by calendar year. The group containing `current_month` is expanded; all others are collapsed. This keeps the visible area focused on the current reporting period.
