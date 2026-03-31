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

### Conditional Formatting Rules

- **Planned variance column**: Yellow fill when `ABS(value) > 0.01`
- **Actual remaining column**: Red fill when `value < 0`
- **Summary remaining EUR column**: Red fill when `value < 0`

### Year Grouping

Month columns are grouped into years via merged header cells. Each year spans up to 12 columns. The calendar dates are derived from the `project_start` named range (Plan!$B$2).

## Cross-Sheet Reference Map

| Summary Cell | References | Via |
|-------------|-----------|-----|
| WP ID/Name | Plan config rows | Green formula |
| Year N Plan FTE | Plan Planned block, months for that year | `SUM(Plan!{range})` |
| Year N Act FTE | Plan Actual block, months for that year | `SUM(Plan!{range})` |
| Year N EUR | Corresponding FTE × `Plan!$E$2` | Formula |
| Remaining EUR | `Plan!G{config_row}` - Actual EUR total | Formula |
| % Consumed | Actual FTE total / `Plan!C{config_row}` | Formula |
