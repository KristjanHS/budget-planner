# Trends & Burn Rate Dashboard — Design Document

**Date:** 2026-04-02
**Status:** Approved
**Motivation:** Web research identified the top budgeting problems in programme projects: budget-to-workload misalignment and late variance detection. The current tool gives point-in-time snapshots but no temporal intelligence — users discover WP-level burn rate problems only at quarterly reporting, and even then the trend isn't visible.

## Goals

1. Surface WP-level burn rate trends so PMs see whether spending is ahead/behind schedule
2. Project budget exhaustion dates per WP based on current spending rate
3. Visualize cumulative planned-vs-actual EUR over the project lifetime
4. Flag person-level overload before it becomes a crisis
5. Give the accountant inline ceiling compliance signals on the Actuals sheet

## Chosen Approach: Trends Sheet + Actuals Cumulative View

A new `Trends` sheet provides the PM dashboard. The Actuals sheet gets 3 inline columns so the accountant sees cumulative spend and breach risk without switching sheets. Both personas are served where they already work.

### Why not extend Summary?

Summary is already dense with annual aggregation. Adding temporal columns would clutter it. A dedicated sheet follows the existing pattern (each sheet has one clear purpose) and is more extensible.

---

## Design Details

### 1. Actuals Sheet Additions

Three new columns added to the right of the existing Variance column:

| Column | Header | Formula Pattern | Conditional Formatting |
|--------|--------|----------------|----------------------|
| VAL_COL_DIFF + 1 | Cum. EUR | `=SUMPRODUCT((COLUMN(C{r}:BJ{r})-COLUMN($C$1)+1<=current_month)*(C{r}:BJ{r}))` | None |
| VAL_COL_DIFF + 2 | % Ceiling | `=CumEUR / Ceiling` | Amber >85%, Red >100% |
| VAL_COL_DIFF + 3 | Breach In | `=IF(rate=0,"—",(Ceiling-CumEUR)/rate)` | Red <6 months, Amber <12 |

- `CumEUR` uses SUMPRODUCT with **column-position math** (not month header comparison — headers are text like `"M1\nJul 26"`, not numbers). The pattern `COLUMN(range)-COLUMN($C$1)+1` yields month numbers 1..60 which can be compared to `current_month`.
- `% Ceiling` divides cumulative EUR by the existing Ceiling column value.
- `Breach In` calculates months remaining at current monthly burn rate. Monthly rate = `CumEUR / current_month` (simplified since actuals accumulate from M1). Falls back to "—" if no actuals entered.
- TOTAL row gets the same 3 columns aggregated against `partner_ceiling`.

### 2. Trends Sheet — Layout

```
Row 1:    Current Month: [blue input cell B1, named range current_month, default=1]
Row 2:    (blank)
Row 3:    Section header: "WP BURN RATE & EXHAUSTION"
Row 4:    Column headers (7 cols: WP ID, WP Name, % Time, % Budget EUR, Burn Delta, Status, Exhaustion Mo.)
Row 5-12: WP data rows (8 WPs)
Row 13:   TOTAL row (weighted average)
Row 14:   (blank)
Row 15:   LineChart (project-level cumulative planned vs actual vs ceiling)
Row ~31:  (blank, chart occupies ~15 rows)
Row 32:   Section header: "PERSON LOAD"
Row 33:   Column headers
Row 34-37: Person data rows (4 team members)
Row 39+:  Hidden helper rows for chart cumulative running totals
```

### 3. Trends Sheet — WP Burn Rate & Exhaustion Table (merged)

Single table combining burn rate signals with projected exhaustion date. The previous design had separate Burn Rate and Projected Exhaustion tables — merged to eliminate redundancy (exhaustion month is a reformulation of burn delta).

**Dropped `% FTE Consumed` column:** `Plan!Sum` (col BK) is the TOTAL planned FTE across all 60 months, not cumulative-to-date. It would show a static ratio regardless of `current_month`, which is misleading as a burn indicator. EUR is the primary burn metric; FTE can be derived as `EUR / fte_rate` where needed.

| Col | Header | Formula | Notes |
|-----|--------|---------|-------|
| A | WP ID | `=Plan!A{config_row}` | Green cross-sheet ref |
| B | WP Name | `=Plan!B{config_row}` | Green cross-sheet ref |
| C | % Time Elapsed | `=IF(current_month<wp_start, 0, MIN(1, (current_month-wp_start+1)/(wp_end-wp_start+1)))` | Clamped 0-100% |
| D | % Budget EUR | `=Actuals!{cum_col}{row} / Actuals!{ceiling_col}{row}` | Cross-sheet to new Actuals CumEUR / Ceiling cols |
| E | Burn Delta | `=D{row} - C{row}` | Positive = burning faster than time |
| F | Status | `=IF(ABS(E{row})<=0.05,"OK",IF(E{row}>0.15,"OVER",IF(E{row}>0.05,"WATCH",IF(E{row}<-0.10,"SLOW","OK"))))` | Text + RAG conditional formatting |
| G | Exhaustion Mo. | `=IF(D{row}=0,"—",IF(COUNTIF(actuals_range,">0")=0,"No data", wp_start+ROUND(Actuals!{ceiling_col}{row}/(Actuals!{cum_col}{row}/COUNTIF(actuals_range,">0")),0)-1))` | Month when ceiling hit at current rate |

**Conditional formatting on Burn Delta (col E):**
- Green: ABS(delta) <= 0.05
- Amber (overspend): delta > 0.05 AND delta <= 0.15
- Amber (underspend): delta < -0.10
- Red: delta > 0.15

**Conditional formatting on Exhaustion Month (col G):**
- Red: exhaustion month < wp_end (budget runs out before WP ends)
- Amber: exhaustion month within 3 months of wp_end
- Green: exhaustion month >= wp_end

**Status column (col F):** Uses both a text formula AND conditional formatting. Text ensures readability in black-and-white printing; color gives the at-a-glance signal. CF rules match the text thresholds.

For WPs not yet started (`current_month < wp_start`), all cells show "—" via `IF` guard.

TOTAL row: weighted average burn delta, weighted by each WP's ceiling as share of partner_ceiling.

**Cell comment on col G header (written by generator):** "Projection is linear based on quarterly actuals. Accuracy improves after Q2."

### 4. Trends Sheet — Cumulative LineChart

**Chart: Project-level cumulative planned vs actual EUR**

- Series 1 (blue solid): Cumulative planned EUR — running sum across all WPs of Plan Planned FTE × fte_rate per month. Helper row formula per column: `=SUM(Plan!{col}{first_wp}:{col}{last_wp})*fte_rate + {prev_cumulative}` (sum all WP planned FTE for that month, convert to EUR, add to running total).
- Series 2 (red solid): Cumulative actual EUR — running sum of Actuals TOTAL row per month
- Series 3 (grey dashed): Ceiling — flat horizontal line at `partner_ceiling`
- X-axis: M1–M60 month labels
- Title: "Cumulative EUR: Planned vs Actual"

**Helper rows (hidden, below chart):**
- Row H1: Cumulative planned EUR — per column: `=SUM(Plan!{col}{first_wp}:{col}{last_wp})*fte_rate` (monthly planned EUR for all WPs), then accumulated: `=H1_prev + H1_current`
- Row H2: Cumulative actual EUR — per column: `=Actuals!{col}{total_row}`, then accumulated: `=H2_prev + H2_current`
- Row H3: Ceiling — constant `=partner_ceiling` repeated across all 60 columns
- Columns C through BJ (60 months, matching MONTH_COL_START layout)

Chart generated by openpyxl `LineChart` at workbook creation time. Data ranges are fixed (60 months, always same layout). Chart is recreated on each `python -m budget_planner` run.

### 5. Trends Sheet — Person Load Table

**Note:** Capacity TOTAL row has monthly SUM formula cells (one per month) but NO validation columns (Sum/Budget/Diff are empty). All Person Load formulas must use SUMPRODUCT over the TOTAL row's monthly cell range `C{total_row}:BJ{total_row}`, not a single summary cell.

| Col | Header | Formula | Notes |
|-----|--------|---------|-------|
| A | Person | `=Team!A{roster_row}` | Green cross-sheet ref |
| B | Avg FTE (to date) | `=IF(current_month=0, 0, SUMPRODUCT((COLUMN(Capacity!C{tr}:BJ{tr})-COLUMN(Capacity!$C$1)+1<=current_month)*(Capacity!C{tr}:BJ{tr})) / current_month)` | Average allocation over elapsed months |
| C | Avg FTE (remaining) | `=IF(current_month>=PROJECT_MONTHS, 0, SUMPRODUCT((COLUMN(Capacity!C{tr}:BJ{tr})-COLUMN(Capacity!$C$1)+1>current_month)*(Capacity!C{tr}:BJ{tr})) / (PROJECT_MONTHS-current_month))` | Implied future load |
| D | Peak Month FTE | `=MAX(Capacity!C{tr}:BJ{tr})` | Worst single month |
| E | Status | `=IF(C{row}>Team!B{roster_row},"OVER",IF(C{row}>0.8*Team!B{roster_row},"WATCH","OK"))` | Text + RAG conditional formatting |

`{tr}` = the person's TOTAL row in Capacity. The SUMPRODUCT pattern uses column-position math (same as Actuals CumEUR) to select months before/after `current_month`.

**Conditional formatting on Status (col E) — based on col C value:**
- Red: Avg Remaining > max_fte (from Team!B{roster_row})
- Amber: Avg Remaining > 0.8 * max_fte
- Green: otherwise

**Conditional formatting on Peak Month (col D):**
- Red: Peak > max_fte

### 6. current_month Named Range

- Location: Trends!$B$1
- Type: Blue input cell (manual entry)
- Default value: 1
- Cell comment: "Set to the last month with completed actuals (e.g., 12 for end of Year 1)"
- Referenced by: Actuals cumulative columns, all Trends sheet formulas

### 7. _META Sheet Update

Add documentation for the Trends sheet structure, column definitions, and the current_month workflow.

### 8. Builder Architecture

New file `budget_planner/trends.py` with function `build_trends_sheet(wb, layout, actuals_layout, capacity_layout, team_layout)`.

Called from `workbook.py` after `build_summary_sheet()`, before `build_meta_sheet()`.

The Actuals additions (cumulative columns) are added in `actuals.py` via extending `build_actuals_sheet()` — 3 new columns after VAL_COL_DIFF.

Sheet tab color: Gold/Yellow (`FFD700`) — distinct from existing sheets.

---

## Key Decisions

1. **Manual `current_month` over auto-detect** — explicit is better for a financial tool; avoids volatile `TODAY()` and mid-quarter noise
2. **Project-level chart, not per-WP** — 16 series is unreadable; per-WP detail lives in the tables
3. **No sparklines** — openpyxl can't generate them reliably; full LineChart instead
4. **Linear projection with documented caveat** — simple and transparent; cell comments set expectations
5. **Actuals inline columns** — accountant sees ceiling compliance without switching sheets
6. **Underspend flagged as Amber** — underspend often signals delayed deliverables, not savings

## Out of Scope

- Per-WP individual charts
- Sparklines
- Auto-detecting current_month from TODAY() or actuals data
- 215-day cross-grant FTE ceiling (requires external data sources)
- Multi-partner budget view
- Actuals FTE tracking (only EUR actuals exist; FTE is planned only)

## Dependency Graph (for implementation)

```
Step 1: Actuals sheet additions (3 new columns + conditional formatting)
  Files: budget_planner/actuals.py, budget_planner/config.py (new column constants)

Step 2: Trends builder — current_month + burn rate & exhaustion table (merged)
  Files: budget_planner/trends.py (new), budget_planner/workbook.py
  Depends on: Step 1 (cross-sheet refs to new Actuals columns)

Step 3: Trends builder — cumulative chart + helper rows
  Files: budget_planner/trends.py
  Depends on: Step 2 (needs sheet context)

Step 4: Trends builder — person load table
  Files: budget_planner/trends.py
  Depends on: Step 2 (needs current_month)

Step 5: Wire into workbook orchestrator + _META update
  Files: budget_planner/workbook.py, budget_planner/meta.py
  Depends on: Steps 1-4

Step 6: Tests
  Files: tests/test_trends.py (new), tests/test_actuals.py (extend)
  Depends on: Step 5

Parallel groups:
  Steps 3 and 4 are independent sections of trends.py but write to the
  same file — must be sequential within the builder.
  Step 1 can run before Steps 2-4.
  Step 6 can only run after Step 5.
```
