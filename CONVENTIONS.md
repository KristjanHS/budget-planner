# CONVENTIONS.md — Spreadsheet Formatting and Structural Conventions

## Cell Type Identification

Every cell in the workbook falls into exactly one category, identifiable by font color:

| Font Color | RGB | Meaning | Editable by User? |
|-----------|-----|---------|-------------------|
| Blue | (0,0,255) | Hardcoded input value | YES |
| Black | (0,0,0) | Formula (calculated) | NO — overwriting destroys logic |
| Green | (0,128,0) | Cross-sheet reference formula | NO — overwriting destroys logic |

**Rule for AI agents**: When modifying cells, always check the font color first. Never overwrite black or green cells with values. Only blue cells are safe to change.

## Fill Color Semantics

| Fill Color | Hex | Meaning |
|-----------|-----|---------|
| Grey | D9D9D9 | Inactive / do not enter data here |
| Yellow | FFFF00 | Attention needed (e.g., planned variance ≠ 0) |
| Red (light) | FFC7CE | Error or budget overrun |
| Light blue | D6E4F0 | Header row/cell |
| Light green | E2EFDA | Totals row |
| Amber | FFF2CC | Warning (e.g., >80% utilization) |
| Dark blue | 2F5496 | Section header (white text) |

## Number Format Standards

| Data Type | Format String | Example |
|----------|--------------|---------|
| FTE (person-months) | `0.00` | 1.50 |
| Currency (EUR) | `€#,##0` | €12,000 |
| Percentage | `0.0%` | 45.2% |
| Month number | `0` | 18 |
| Date | `MMM YYYY` | Jul 2026 |

## Structural Patterns

### Config-to-Grid Linkage

Every WP row in the Planned/Actual blocks references its config row via formula:
- Column A (WP ID): `=A{config_row}` (green font)
- Column B (WP Name): `=B{config_row}` (green font)

This means renaming a WP in config automatically propagates everywhere.

### Validation Pattern

Every monthly grid block has three validation columns at the right edge:

| Column | Formula | Purpose |
|--------|---------|---------|
| Sum | `=SUM(C{row}:BJ{row})` | Total of all monthly values for this WP |
| Budget FTE | `=C{config_row}` | Cross-ref to config (green font) |
| Variance/Remaining | `=Budget - Sum` | Zero = OK for planned; negative = overrun for actuals |

### Year Grouping

Month columns are grouped into years via merged header cells. Each year spans up to 12 columns. The calendar dates are derived from the `project_start` named range (Config!$B$2).

### Column Grouping

Month columns on Plan, Capacity, and Actuals sheets are grouped by year at outline level 1. On `rebuild`, the group containing the current month is expanded and all others are collapsed. On `init`, Year 1 is expanded and the rest are collapsed. This reduces horizontal scrolling across 60+ month columns.

### Cell Protection

All sheets are protected with password "grant". By default, all cells are locked. Only blue input cells are unlocked via `Protection(locked=False)`. This prevents accidental formula overwrites while allowing PMs to edit input values freely. The password is documented in the _META sheet.

## Cross-Sheet Reference Map

| Target Cell | References | Via |
|-------------|-----------|-----|
| **Config → Plan** | | |
| Named range `project_start` | Config!$B$2 | Named range definition |
| Named range `fte_rate` | Config!$B$4 | Named range definition |
| **Plan → Config** | | |
| Plan global cells (project start, rate) | Config sheet params | Formula refs to Config (green formula) |
| **Plan → Capacity** | | |
| Plan Planned FTE (monthly) | Capacity per-person WP rows for that month | `=SUM(Capacity!{cells})` (green formula) |
| **Capacity → Plan** | | |
| Capacity WP ID/Name | Plan config rows | `=Plan!A{row}` / `=Plan!B{row}` (green formula) |
| Capacity WP Budget | Team share × Plan config Total FTE | `=Team!{share}*Plan!C{row}` (green formula) |
| **Capacity → Team** | | |
| Capacity person header | Team roster name | `=Team!A{row}` (formula) |
| Capacity REMAINING | Team roster Max FTE | `=Team!B{row}-{TOTAL}` (formula) |
| **Team (internal)** | | |
| Share matrix person name | Team roster person name | `=A{roster_row}` (formula) |
| **Summary → Plan** | | |
| Summary WP ID/Name | Plan config rows | Green formula |
| Year N Plan FTE | Plan Planned block, months for that year | `SUM(Plan!{range})` |
| Year N Act FTE | Plan Actual block, months for that year | `SUM(Plan!{range})` |
| Year N EUR | Corresponding FTE × `Plan!$E$2` | Formula |
| Remaining EUR | `Plan!G{config_row}` - Actual EUR total | Formula |
| % Consumed | Actual FTE total / `Plan!C{config_row}` | Formula |

## Conditional Formatting Rules

### Plan Sheet
- **Planned variance column**: Yellow fill when `ABS(value) > 0.01`
- **Actual remaining column**: Red fill when `value < 0`

### Team Sheet
- **WP share SUM row**: Yellow fill when `ABS(SUM - 1.0) > 0.001`

### Capacity Sheet
- **TOTAL row per person**: Amber fill when `TOTAL > 0.8 × Max FTE`, Red fill when `TOTAL > Max FTE`
- **REMAINING row per person**: Red fill when `REMAINING < 0`

### Config Sheet
- **WP start/end columns**: Yellow fill when WP start month > end month

### Summary Sheet
- **Remaining EUR column**: Red fill when `value < 0`
