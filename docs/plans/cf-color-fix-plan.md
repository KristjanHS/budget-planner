# Plan: Fix Conditional Formatting Colors

**Status:** Draft
**Date:** 2026-04-05

## Problem

Most CF rules show wrong colors in Excel. The fills render incorrectly despite
commit fa772b1 switching to `PatternFill(bgColor=...)` for DXF context.

## Root Cause

All four `CF_*_FILL` constants in `styles.py` use **6-character hex** codes:

```python
CF_YELLOW_FILL = PatternFill(bgColor="FFFF00")   # openpyxl stores as 00FFFF00
CF_AMBER_FILL  = PatternFill(bgColor="FFD966")   # openpyxl stores as 00FFD966
CF_RED_FILL    = PatternFill(bgColor="FFC7CE")    # openpyxl stores as 00FFC7CE
CF_GREEN_FILL  = PatternFill(bgColor="E2EFDA")    # openpyxl stores as 00E2EFDA
```

openpyxl's `Color` class prepends `00` to 6-char hex, making the alpha channel
**00 (transparent)**. The generated DXF XML reads:

```xml
<patternFill><bgColor rgb="00FFFF00"/></patternFill>
```

Excel's ARGB format is `AARRGGBB`. Alpha `00` = fully transparent.
The correct alpha is `FF` (opaque).

## Evidence

Generated `xl/styles.xml` from demo workbook confirms all four DXF entries
have `00` alpha prefix. Tested that 8-char hex `"FFFFFF00"` produces
`<bgColor rgb="FFFFFF00"/>` (correct).

## Formula & Priority Audit

All CF formulas and priority ordering reviewed — **no logic issues found**:

| Sheet | Rules | Colors | Logic |
|-------|-------|--------|-------|
| Config | 1 per WP | Yellow | Start > End validation OK |
| Plan | 1 range | Red | Variance < 0 OK |
| Actuals | 5 range | Red, Amber | Variance + %Ceiling + Breach OK |
| Capacity | 1/WP + 4/person | Red, Amber | Per-row refs correct |
| Summary | per-year + 3 | Red, Amber | Variance + Remaining OK |
| Trends | 10 + 3/WP + 4/person | Green, Amber, Red | All thresholds correct |

Priority order (amber first, red last) is correct everywhere.
Per-row rules used correctly where each row needs different Config refs.

## Fix Plan

### Step 1: Fix CF fill constants in styles.py (the only change)

```python
# CF (conditional formatting) fills — DXF context requires bgColor, not fgColor
# Use 8-char AARRGGBB hex; 6-char hex gets 00 alpha (transparent)
CF_YELLOW_FILL = PatternFill(bgColor="FFFFFF00")
CF_AMBER_FILL  = PatternFill(bgColor="FFFFD966")
CF_RED_FILL    = PatternFill(bgColor="FFFFC7CE")
CF_GREEN_FILL  = PatternFill(bgColor="FFE2EFDA")
```

No other files need changes — all builders import these constants.

### Step 2: Update tests

`tests/test_conditional_formatting.py` extracts fill colors for assertions.
The helper function needs to expect `FF`-prefixed colors:

- `"00FFFF00"` -> `"FFFFFF00"` (yellow)
- `"00FFD966"` -> `"FFFFD966"` (amber)
- `"00FFC7CE"` -> `"FFFFC7CE"` (red)
- `"00E2EFDA"` -> `"FFE2EFDA"` (green)

### Step 3: Update dev-workflow rule

Update the CF fill pattern in `.claude/rules/dev-workflow.md` to document
the 8-char hex requirement:

```
CF fills must use `PatternFill(bgColor="FFFFC7CE")` (8-char AARRGGBB hex,
not 6-char) — 6-char hex gets 00 alpha (transparent).
```

### Step 4: Verify

1. `pytest tests/ -q`
2. Generate demo workbook + recalc
3. Inspect DXF XML confirms `FF` alpha prefix
4. Visual check in Excel/LibreOffice
