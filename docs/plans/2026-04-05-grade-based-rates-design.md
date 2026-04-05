# Design: Grade-Based Personnel Rates

**Date**: 2026-04-05
**Status**: Design complete, not yet implemented
**Depends on**: None (can be implemented independently)

## Problem

The tool uses a single global `fte_rate` (EUR/person-month) for all personnel. EU auditors expect costs based on actual per-person rates derived from annual employment cost / 215 day-equivalents. The PM won't receive individual salary data, but bookkeeping can provide rates by job grade (Senior / Mid / Junior).

## Decision Log

| Choice | Options considered | Selected | Rationale |
|--------|-------------------|----------|-----------|
| Rate granularity | Single rate / Per-grade / Per-person / Per-person-per-year | Per-grade (3 bands) | PM won't get individual salary data; bookkeeping can provide grade bands |
| Grade table location | Between params and WP table / Hidden reference | Visible table in Config | Auditable formula chain; PM sees rates at a glance |
| Ceiling calculation | Weighted average per WP / Blended team average | Blended team average | Preserves existing sheet structure; avoids complex per-WP formulas |
| Existing fte_rate | Remove / Keep as formula | Keep as formula | All existing formulas referencing fte_rate continue to work unchanged |

## Design

### 1. Grade Rate Table (Config sheet)

New section between project params (row 5) and WP table, labeled "RATE BY GRADE":

- 3 rows: Senior, Mid, Junior
- Column A: grade name, Column B: EUR/person-month (blue input)
- Named range `tbl_GradeRates` over the 2-column table (for VLOOKUP)
- Default values: Senior €10,000 / Mid €7,500 / Junior €5,500 (illustrative)

### 2. Roster Grade Column

New column in the team roster table: "Grade"

- Blue input cells with data validation dropdown: Senior, Mid, Junior
- Default: "Mid" for all existing members
- Used by the fte_rate formula to compute weighted average

### 3. fte_rate Becomes a Formula

Current: `fte_rate` = Config!B4, blue input, user-entered constant.

New: `fte_rate` = Config!B4, black formula, auto-calculated:

```
=SUMPRODUCT(VLOOKUP(roster_grades, tbl_GradeRates, 2, FALSE) * roster_ftes) / SUM(roster_ftes)
```

Weighted average: each person's grade rate weighted by their total FTE commitment.

### 4. Downstream Impact

**No changes required** to these sheets — they all reference `fte_rate` which still exists:
- Config: WP Personnel Ceiling = `FTE_Budget * fte_rate` (unchanged)
- Plan: EUR display uses `fte_rate` (unchanged)
- Summary: Plan EUR, Variance EUR use `fte_rate` (unchanged)
- Actuals: Ceiling references Config column H (unchanged)

### 5. Reader/Rebuild

The reader must:
- Read grade table (grade names + rates) from Config
- Read each person's grade from the roster
- Preserve both through rebuild

config.py additions:
- `GRADE_RATES` dict: `{"Senior": 12500, "Mid": 10000, "Junior": 7500}`
- `DEFAULT_GRADE`: `"Mid"`
- Grade field added to team member data structure

### 6. Out of Scope

- Per-person EUR tracking in Actuals (stays WP-level)
- Mid-year rate changes (use annual average per grade)
- Per-WP weighted rates (blended team average only)
- 20% budget transfer threshold monitoring (handled by overall PM, not this tool)
- Additional grade bands (start with 3; can add more later)
