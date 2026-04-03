# Fix Plan Seed-Then-Overwrite Pattern (C1)

**Date:** 2026-04-03
**Status:** Complete
**Origin:** Code review finding C1 from improvements2 implementation

## Problem

`build_plan_sheet()` writes planned FTE month cells in two passes:

1. **Pass 1** (plan.py lines 187–195): Writes `even_fte` numeric values with `BLUE_INPUT` font for active months, empty + grey for inactive months.
2. **Pass 2** (`_apply_planned_formulas`, plan.py lines 58–73): Overwrites every cell with `=SUM(Capacity!...)` formulas using `BLACK_FORMULA` font, and explicitly re-locks the cell.

This violates the font-color invariant ("Blue = input, Black = formula") during construction. The final workbook is correct, but:
- `build_plan_sheet()` is public and called independently in test fixtures without the follow-up patch — those fixtures see blue-colored numeric cells instead of formulas.
- If `_build_all_sheets` fails partway (after Plan, before Capacity), the partial workbook has blue-colored seed values that look like user input but are actually scaffolding.
- The comment on line 67 ("Re-lock: `_build_month_block` wrote these with BLUE_INPUT first") documents this as intentional but it's a code smell.
- The seed values themselves are structurally redundant — `_apply_planned_formulas` unconditionally overwrites every cell.

## Design

Remove the seed value write-pass. Initialize Plan month cells with the correct font/protection from the start.

### Changes to `build_plan_sheet()` (plan.py)

Replace the month cell loop (lines 187–195):

```python
# BEFORE:
for m in range(project_months):
    c = month_col_start + m
    month_num = m + 1
    if wp["start"] <= month_num <= wp["end"]:
        style_cell(ws, row, c, even_fte, font=BLUE_INPUT, fmt="0.00")
    else:
        style_cell(ws, row, c, "", font=BLUE_INPUT, fill=GREY_FILL, fmt="0.00")

# AFTER:
for m in range(project_months):
    c = month_col_start + m
    month_num = m + 1
    if wp["start"] <= month_num <= wp["end"]:
        cell = style_cell(ws, row, c, font=BLACK_FORMULA, fmt="0.00")
        cell.protection = Protection(locked=True)
    else:
        cell = style_cell(ws, row, c, "", font=BLACK_FORMULA,
                          fill=GREY_FILL, fmt="0.00")
        cell.protection = Protection(locked=True)
```

Active cells are written with `BLACK_FORMULA` font, no value, and locked protection from the start. `_apply_planned_formulas` then fills in the `=SUM(Capacity!...)` formula values.

### Changes to `_apply_planned_formulas()` (plan.py)

Simplify the overwrite pass — remove the "Re-lock" comment and the explicit `cell.protection` lines since cells are already correctly styled:

```python
# BEFORE:
cell = style_cell(ws, row, c, font=BLACK_FORMULA, fmt="0.00")
cell.value = f"=SUM({refs})"
cell.protection = Protection(locked=True)  # Re-lock

# AFTER:
ws.cell(row=row, column=c).value = f"=SUM({refs})"
```

Only the value changes — font, format, and protection are already correct from `build_plan_sheet`.

### Test impact

Test fixtures that call `build_plan_sheet()` without `_apply_planned_formulas()` (e.g. `test_plan_v2.py`) will now see empty BLACK_FORMULA cells instead of blue seed values. Update assertions accordingly:
- Planned FTE cells should have `BLACK_FORMULA` font (not `BLUE_INPUT`)
- Active cells should be empty/None until Capacity formulas are applied
- Font-color tests should pass without the two-pass workaround

## Dependency Graph

```
Step 1: Modify build_plan_sheet month cell loop     (plan.py)
Step 2: Simplify _apply_planned_formulas             (plan.py, same file)
Step 3: Update test assertions                       (tests/test_plan_v2.py)
        All sequential — same file
```

## Files

- `budget_planner/plan.py` — both functions
- `tests/test_plan_v2.py` — assertion updates
