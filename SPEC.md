# Budget Planner — Project Specification

## Purpose

Excel workbook for tracking planned and actual FTE allocation per work package (WP) per month across a 5+ year EU-funded (programme) project. Derives euro amounts from FTE using a configurable rate. Designed for a single organisation's share of a multi-partner project.

## Business Context

- **User**: Manages an R&D team's FTE across EU project work packages
- **Funder constraint**: Total FTE per WP must stay within budget. Auditors flag time logged to a WP outside its active period. Annual planned-vs-actual is informational, not strictly enforced.
- **Primary unit**: FTE (person-months). Euros are derived internally (FTE × monthly rate), not reported to funder.
- **Single org**: Only the organisation's share is tracked. No multi-partner view needed.

## Architecture

### Sheets

| Sheet | Purpose | Tab Color |
|-------|---------|-----------|
| `_META` | AI-agent-readable documentation of workbook structure, conventions, modification instructions | Red |
| `Plan` | Operational sheet: config + planned FTE + actual FTE | Blue |
| `Summary` | Annual aggregation with budget health indicators | Green |

### Plan Sheet — Block Layout

```
Row 1-2:   GLOBAL PARAMETERS (project title, start date, FTE rate)
Row 4-N:   CONFIGURATION TABLE (WP definitions)
Row N+3:   PLANNED FTE block (monthly grid, validation)
Row M+2:   ACTUAL FTE block (monthly grid, remaining budget validation)
```

### Configuration Table (per WP)

| Field | Type | Description |
|-------|------|-------------|
| WP ID | input | e.g. "WP1" |
| WP Name | input | e.g. "Project Management" |
| Total FTE (PM) | input | Total person-months budget for this WP |
| Start Month | input | First active project month (1-based) |
| End Month | input | Last active project month (1-based) |
| Active Months | formula | End - Start + 1 |
| Budget (€) | formula | Total FTE × rate |

### Monthly FTE Blocks (Planned and Actual)

- Rows = WPs (linked to config via formula)
- Columns = project months M1..M60 (with calendar date labels derived from project start date)
- Columns grouped visually by year with merged year headers
- Inactive months (outside WP's start–end range) are greyed out with empty values
- Planned block: pre-filled with even distribution (Total FTE ÷ Active Months), user overrides freely
- Actual block: starts at 0, user enters logged time

### Validation Columns (rightmost in each block)

| Column | Planned Block | Actual Block |
|--------|--------------|--------------|
| Sum | SUM of monthly values | SUM of monthly values |
| Budget FTE | Cross-ref to config Total FTE | Cross-ref to config Total FTE |
| Variance/Remaining | Budget - Sum (yellow if ≠ 0) | Budget - Sum (red if < 0) |

### Summary Sheet

- Rows = WPs
- Column groups per year (5 cols each): Plan FTE, Act FTE, Var FTE, Plan €, Act €
- Project Total group (7 cols): Plan FTE, Act FTE, Var FTE, Plan €, Act €, Remaining €, % Used
- Totals row at bottom
- Red fill on negative Remaining €

## Conventions

### Font Colors (financial model standard)

- **Blue (0,0,255)**: Hardcoded inputs — user edits these
- **Black (0,0,0)**: Formulas — never overwrite
- **Green (0,128,0)**: Cross-sheet reference formulas

### Fill Colors

- **Grey (D9D9D9)**: Inactive / protected area
- **Yellow (FFFF00)**: Attention needed (planned variance ≠ 0)
- **Red (FFC7CE)**: Error / budget overrun
- **Light blue (D6E4F0)**: Headers
- **Light green (E2EFDA)**: Totals rows

### Number Formats

- FTE values: `0.00`
- EUR values: `€#,##0`
- Percentages: `0.0%`
- Years: text (not formatted as number)

### Named Ranges

- `project_start` → `Plan!$B$2`
- `fte_rate` → `Plan!$E$2`

## Known Limitations / Future Work

- **Rounding**: Even distribution produces tiny variance (~0.001 FTE) when Total FTE doesn't divide evenly by Active Months. User can manually adjust one cell to zero it out. A future version could auto-correct the last month.
- **No per-person breakdown**: Currently tracks WP-level FTE only. Could add a per-team-member layer.
- **No cost categories**: Only personnel costs (FTE-based). Travel, equipment, subcontracting are not modeled.
- **No reporting period alignment**: Months are calendar months from project start. If reporting periods don't align with calendar years, Summary grouping would need adjustment.
- **No data validation dropdowns**: WP start/end months are free-entry. Could add data validation to constrain range.
- **No chart/dashboard**: Could add a visual dashboard sheet with burn-down charts per WP.
- **Protection**: Cells are not locked/protected. Could add sheet protection with unlocked input cells only.
- **Actuals import**: Currently manual entry. Could add a macro or script to import from timesheet exports.

## Dependencies

- Python 3.10+
- `openpyxl` >= 3.1.0
- `python-dateutil`
- LibreOffice (for formula recalculation via `scripts/recalc.py`)
