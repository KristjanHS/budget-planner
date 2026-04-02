# Programme Budget Sheet Flow Design

Design for extending the programme budget planner workbook to support the full programme actual-cost budgeting lifecycle (Phases 2–7, excluding Phase 1: Proposal).

## Goals

- Track personnel FTE planning and EUR actuals within our partner envelope
- Provide ceiling tracking against grant agreement personnel budget
- Auto-calculate 25% indirect flat rate for claimable amount visibility
- Keep actuals separate from planning (actuals come as EUR from accountant, not FTE per person)
- Remove dead weight (Actual FTE block on Plan) to avoid source-of-truth confusion

## Scope Decisions

| Decision | Choice | Rationale |
|---|---|---|
| Multi-partner? | **No** — single partner | Other partners use their own tools |
| Non-personnel cost tracking? | **No** — tracked by overall project manager separately | Not this dept's responsibility |
| Reporting periods? | **No** — annual view sufficient | Monthly granularity already available for ad-hoc slicing |
| Actuals granularity? | **EUR per WP per month** (from accountant) | Not per-person, not FTE-based |
| Ceiling scope? | **Personnel costs only** | 25% indirect is additional budget line, not carved from personnel |

## Sheet Overview

| # | Sheet | Tab Color | Role | Programme Phase |
|---|-------|-----------|------|-----------|
| 1 | _META | Red | AI-agent documentation | — |
| 2 | Plan | Default | Project config + Planned FTE grid + personnel ceilings | 2–3 |
| 3 | Team | Orange | Team roster + WP share seeds | 3 |
| 4 | Capacity | Purple | Per-person monthly FTE scheduling | 3–4 |
| 5 | Actuals | Blue | EUR actuals per WP per month (from accountant) | 4–5 |
| 6 | Summary | Green | Annual FTE/EUR aggregation + budget overview | 4–7 |

## Sheet Details

### Plan (revised)

**Changes from current:** Remove Actual FTE block. Add Personnel Ceiling. Compact global params to columns A–B only (no wide labels in month columns).

**Global params (rows 1–4, columns A–B only):**

| Cell | Content | Type |
|------|---------|------|
| A1: Project Title | B1: [Project Name] | input (blue) |
| A2: Project Start | B2: [2026-01-01] | input (blue) |
| A3: FTE Rate (EUR) | B3: [6500] | input (blue) |
| A4: Personnel Ceiling | B4: [=SUM of WP ceilings] | formula (black) |

Named ranges: `project_start` → B2, `fte_rate` → B3, `partner_ceiling` → B4.

**WP config table (rows 6+):**

Columns A–B: WP ID, WP Name. Dense config columns (C–H): FTE Budget (PM), Start Month, End Month, Active Months (formula), Personnel Ceiling EUR (= FTE Budget × fte_rate, formula, black), plus any existing fields.

**Planned FTE grid (below config):**

Unchanged from current. Header + year/month labels + one row per WP + TOTAL. Monthly cells are SUM formulas referencing Capacity. Validation columns: Sum, Budget, Variance.

### Team (unchanged)

Team roster (name, max FTE) + WP share seed matrix. No changes.

### Capacity (unchanged)

Per-person monthly FTE allocation grids. No changes.

### Actuals (new)

New sheet for accountant-provided EUR data.

```
Row 1:  ACTUAL EUR header (merged)
Row 2:  Year labels      | M1 | M2 | ... | M48 | Sum | Ceiling | Variance
Row 3:  Month labels (actual dates)
Row 4+: One row per WP — blue input cells (EUR values from accountant)
Row N:  TOTAL row (SUM formulas per column)
```

- Month columns aligned with Plan sheet (same column positions C onward)
- Blue input cells for EUR values
- Sum column: row total per WP (formula)
- Ceiling column: formula referencing Plan's per-WP Personnel Ceiling
- Variance column: Ceiling − Sum (formula). Conditional formatting: red if negative
- Inactive months greyed out (same convention as Plan)
- TOTAL row: SUM of all WPs per column. Ceiling cell: =Plan!partner_ceiling. Variance cell: Ceiling − Sum.

### Summary (extended)

**Existing annual table (revised):**

Per-year column groups: Plan FTE, Plan EUR, Actual EUR (new — from Actuals sheet), Variance EUR (new — Plan EUR minus Actual EUR). Variance EUR conditional formatting: red if negative (actuals exceed plan). Project Total columns aggregate all years.

Rows: one per WP + TOTAL.

**New budget overview section (below annual table):**

| Row | Label | Value | Type |
|-----|-------|-------|------|
| N+2 | **BUDGET OVERVIEW** | (header) | — |
| N+4 | Personnel Actual | =Actuals!TOTAL sum | formula (green, cross-sheet) |
| N+5 | Total Claimable (+25%) | =Personnel Actual × 1.25 | formula (black) |
| N+6 | Personnel Ceiling | =Plan!partner_ceiling | formula (green, cross-sheet) |
| N+7 | Remaining | =Ceiling − Personnel Actual | formula (black) |
| N+8 | % Consumed | =Personnel Actual ÷ Ceiling | formula (black), percentage format |

Conditional formatting on Remaining: red if negative, amber if < 10% of ceiling.

## Data Flow

```
Team (roster + WP seeds)
  │
  ▼
Capacity (per-person monthly FTE allocation)
  │ SUM formulas
  ▼
Plan (Planned FTE grid + config + ceilings)
  │                          │
  │ WP structure,            │ WP structure,
  │ Plan EUR                 │ ceilings
  ▼                          ▼
Summary ◄──────────── Actuals (EUR input from accountant)
  (annual aggregation +
   budget overview)
```

**No circular references.** Actuals is a pure input sheet — nothing feeds into it except WP structure from Plan.

**Build order:** Plan → Team → Capacity → patch Plan (Planned FTE SUMs) → Actuals → Summary.

### Flow detail

| Source | Target | Data |
|--------|--------|------|
| Team → Capacity | Person names, max FTE | formula refs |
| Plan → Capacity | WP names | formula refs |
| Capacity → Plan | Monthly FTE per WP (summed across people) | SUM formulas |
| Plan → Actuals | WP names, WP start/end months, month structure, per-WP personnel ceiling | formula refs |
| Plan → Summary | Planned FTE, Plan EUR per WP per year, partner_ceiling | formula refs |
| Actuals → Summary | Actual EUR per WP per year, total actual EUR | formula refs |

## Key Decisions

1. **Actuals in EUR, not FTE** — accountant provides EUR; converting back to FTE adds noise and no value.
2. **Separate Actuals sheet** — Plan stays purely about planning. No source-of-truth confusion.
3. **Personnel ceiling only** — 25% indirect is a separate grant budget line, shown as "Total Claimable" for reporting but ceiling comparison is personnel-to-personnel.
4. **No reporting periods** — monthly granularity in Actuals allows ad-hoc slicing when needed.
5. **Global params in A–B only** — prevents auto-fit from widening month columns.
6. **Per-WP ceiling in config area** — next to Active Months. Keeps related WP properties together; narrow EUR format doesn't inflate month column widths.

## Out of Scope

- Multi-partner consortium view
- Non-personnel cost tracking (travel, equipment, subcontracting, other)
- Reporting period configuration
- Grant amendment tracking
- Per-person actual time logging (handled by timesheet system)
