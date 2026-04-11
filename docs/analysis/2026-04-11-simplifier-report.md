# Python Simplifier Analysis — Plan Target Files

**Date**: 2026-04-11
**Scope**: Files modified by `docs/plans/2026-04-11-assignments-to-config.md`
**Tool**: python-simplifier (analyze_all + individual analyzers)

## Summary

| File | Lines | HIGH | MED | LOW | Worst function |
|------|------:|-----:|----:|----:|----------------|
| capacity.py | 543 | 4 | 1 | 3 | `build_capacity_sheet` CC=50, CogC=76, 356 lines, 11 params |
| reader.py | 681 | 5 | 2 | 6 | `read_capacity_sheet` CC=40, CogC=73, 131 lines, nesting=6 |
| config_sheet.py | 384 | 2 | 1 | 1 | `build_config_sheet` CC=20, 283 lines, 8 params |
| workbook.py | 517 | 4 | 2 | 1 | `_validate_config` CC=21, CogC=32; `_build_all_sheets` 11 params |
| meta.py | 583 | 1 | 0 | 0 | `build_meta_sheet` 583 lines (data-driven, low risk) |

Package-wide: 322 issues (25 HIGH, 28 MED, 269 LOW). 6 duplicate patterns, 4 unused params, 10 unpythonic patterns.

## Plan Impact Assessment

The plan adds **two new sections** to `build_config_sheet` (assignments + external commitments), **validation formula columns**, and **named ranges** — then removes config blocks from `build_capacity_sheet`. Without refactoring:

| Risk | Description |
|------|-------------|
| `build_config_sheet` grows to ~500 lines | Already 283 lines. Adding ~100 lines of section writing + formulas + named ranges pushes it past any maintainability threshold. |
| `build_config_sheet` params grow to 10+ | Already at 8. Plan adds `assignments` and `external_commitments`. |
| `read_config_sheet` adds complexity to CC=26 | Plan adds section detection + parsing for two new tables. |
| `read_capacity_sheet` backward compat in CC=40 | Plan adds fallback path — more branches in the most complex function. |
| Magic column numbers multiply | Plan adds columns F-H with formulas. Each is a hardcoded integer. |

## Refactoring Recommendations (prioritized)

### R1 — Extract section builders from `build_config_sheet` (HIGH)

**Why**: At 283 lines, adding two sections makes this unmaintainable. Each section follows the same pattern: write section header, write column headers, write data rows, write formulas/totals.

**How**: Extract into focused functions:
- `_write_budget_section(ws, row, work_packages, ...) -> next_row`
- `_write_schedule_section(ws, row, wp_schedule, ...) -> next_row`
- `_write_roster_section(ws, row, team, ...) -> next_row`
- `_write_assignments_section(ws, row, assignments, ...) -> next_row` (new)
- `_write_external_section(ws, row, external, ...) -> next_row` (new)
- `_write_reporting_section(ws, row, reporting_periods, ...) -> next_row`
- `_write_rate_by_grade_section(ws, row, grades, ...) -> next_row`

The top-level function becomes a coordinator that calls sections in order and collects layout keys.

**When**: During the plan — natural point since you're adding sections anyway.

### R2 — Extract section builders from `build_capacity_sheet` (HIGH)

**Why**: CC=50, CogC=76, 356 lines. Even after the plan removes config blocks (~40 lines), it remains ~310 lines with CC~45.

**How**: Split into:
- `_write_capacity_config_blocks(ws, row, ...) -> next_row` (removed by plan)
- `_write_person_grid(ws, row, member, ...) -> next_row`
- `_write_capacity_totals(ws, row, ...) -> next_row`
- `_write_capacity_summary_row(ws, row, ...) -> next_row`

**When**: During the plan — you're already removing code from this function.

### R3 — Split `read_capacity_sheet` (HIGH)

**Why**: CC=40, CogC=73, nesting=6. Plan adds a backward-compat branch, making it worse.

**How**: Extract per-section readers:
- `_read_assignments_from_capacity(ws, ...) -> list[dict]`
- `_read_externals_from_capacity(ws, ...) -> list[dict]`
- `_read_person_grids(ws, ...) -> dict`

The top-level function detects section headers and delegates.

**When**: Before or during the plan — the backward compat path needs clean separation.

### R4 — Parameter bundling (MEDIUM)

**Why**: `build_capacity_sheet` has 11 params, `_build_all_sheets` has 11, `build_config_sheet` has 8 (growing to 10+). These signatures are error-prone — easy to swap positional args.

**How**: Create a `BuildContext` dataclass:
```python
@dataclass
class BuildContext:
    project_params: dict
    work_packages: list
    wp_schedule: list
    team: list
    assignments: list
    external_commitments: list
    reporting_periods: list
    grade_rates: dict
```

Builders receive `(ws, wb, ctx)` instead of 8-11 individual params.

**When**: After the plan — this is a cross-cutting refactor best done once the new structure stabilizes.

### R5 — Column constants per section (MEDIUM)

**Why**: 44 magic numbers across config_sheet.py and reader.py. Column indices (3, 4, 5, 6) are repeated without names. The plan adds columns F-H, compounding this.

**How**: Per section, define column constants:
```python
# WP Budget columns
COL_WP_ID, COL_WP_NAME, COL_WP_FTE, COL_WP_RATE, COL_WP_CEILING = range(1, 6)
# (extend for new Assigned PM, Variance columns)
```

Both builder and reader reference the same constants — a column shift becomes a one-line change.

**When**: During the plan — particularly for the new sections. Retrofitting old sections can follow.

### R6 — Clean up unused params in `build_capacity_sheet` (LOW)

**Why**: `preserved` and `wb` params are never used. Dead parameters hide the actual interface.

**How**: Remove `preserved` and `wb` from the signature. Update all callers.

**When**: During the plan — you're already modifying the function signature.

### R7 — Replace manual index tracking with `enumerate()` (LOW)

**Why**: 7 instances of manual `i += 1` patterns in capacity.py and config_sheet.py.

**When**: Opportunistically during the plan when touching those loops.

## Duplicates Worth Addressing

| Pattern | Locations | Lines saved |
|---------|-----------|-------------|
| `if project_params is None or work_packages is None or wp_schedule is None` | actuals.py:34, plan.py:36, plan.py:86 | ~16 |
| Validation loops in workbook.py | lines 86, 109 | ~5 |
| Summary year/RP column formatting | summary.py: 173/383, 174/384, 180/393, 187/404 | ~20 |

The summary.py duplicates are outside plan scope. The validation pattern in workbook.py is worth extracting if the plan adds more validation.

## Execution Strategy

**During the plan (R1, R2, R3, R5 partial, R6, R7)**:
You're already modifying config_sheet.py, capacity.py, and reader.py. Extract sections as you add/remove code — it costs little extra and prevents the functions from growing worse.

**After the plan (R4, R5 complete)**:
Parameter bundling and full column-constant coverage are best done once the new structure is stable and tests pass.
