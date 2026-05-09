# Actuals posted-cumulative semantics — design

**Date:** 2026-05-09
**Status:** Design approved, awaiting implementation plan
**Scope:** Actuals + Actuals-Hours input convention shift; deletes Plan-Forecast tab, Actuals-Hours forecast-EUR band, EU TOTAL row, "Cumulative hours to date" row, "Actual PM to date" row.

**Amendment 2026-05-09 (post brutal-honesty simplification pass):** dropped automatic migration warning + schema stamp (migration is now a manual one-time clear); replaced last-pair slope + manual override with avg-burn slope (§3); collapsed Trends per-month forecast helper grid into 2 anchor cells per WP (§9); collapsed Summary per-year `Forecast EUR` block into single per-WP column (§9); kept named ranges `eur_actuals_through_month` / `hours_actuals_through_month` unchanged (no semantic gain in renaming). Phasing + test inventory split out — to be written when impl starts.

**Amendment 2026-05-09 (vocabulary cleanup, Approach C):** "Posted" lifecycle vocabulary used end-to-end on the per-WP surface — internal helpers `posted_idx` / `posted_value`, column headers `Posted -i` / `Posted +i`, sheet header `Posted through:`. Sheet-wide named ranges keep `..._through_month` because they describe a different semantic axis (sheet-wide reconciliation cutoff = `MIN` across WPs, distinct from per-WP latest posting). Three temporal concepts now have three distinct words: `current_month` (project clock) / `..._through_month` (sheet reconciliation) / `posted_*` (per-WP latest posting).

## Goals

1. Make accountant-driven posted-cumulative EUR snapshots the native input convention. Allow gaps in monthly data (snapshots arrive irregularly).
2. Derive the per-WP forecast automatically from the posted-cumulative slope, removing the need for a separate hours-derived forecast pipeline.
3. Cut the derived/redundant surface area: one source of truth per WP × month for actuals, one slope cell per WP for forecast, one tab fewer.
4. Preserve `±i` indirect-mirror convention and the inactive-month greying rule unchanged.

## Chosen approach: Level C (maximal simplification) + EU TOTAL removal

Rejected alternatives:
- **Level A (semantic shift only)** — keeps Plan-Forecast and the Actuals-Hours forecast-EUR band even though both become redundant. Two surfaces computing the same forecast; not worth it.
- **Level B (delete Plan-Forecast only)** — keeps Actuals-Hours forecast-EUR band for person-level EUR audit. User explicitly said WP-level forecast is enough; the band carries no extra signal.
- **±i mirror simplification** — explicitly out of scope (user choice).

## Design details

### 1. Input convention

Both actuals tabs hold **posted-cumulative values per cell**.

| Tab | Cell convention |
|---|---|
| `Actuals` | `M[wp, m]` = WP total EUR posted through month `m` (`-i` source; `+i` derived) |
| `Actuals-Hours` | `H[person, wp, m]` = person-on-WP total hours posted through month `m` |

Rules:
- Blank cell = "no posting recorded at this month." Not zero.
- A filled cell is the running total posted through that month, regardless of neighbors.
- Monotonic-non-decreasing within a row's active window is expected but **not enforced** (correction postings can go backward).
- The latest filled cell at-or-before `current_month` IS the WP's currently-posted total.
- Inactive-month greying unchanged. Filled cells inside the active window can have gaps.
- ±i mirrors unchanged: `+i = ROUND(-i × 1.25, 0)` per cell.

**Worked example, WP1:**

```
       Jan | Feb | Mar | Apr | May | Jun ... Dec
WP1          200   380         500
              ↑     ↑           ↑
              posted posted    posted (latest)
              Feb    Mar       May = €500
```

Per-month spend (derived, never stored): Mar = 380 − 200 = 180. May = 500 − 380 = 120. Gap-spreading: between two filled points, the slope `(cum_n − cum_prev) / (m_n − m_prev)` spreads the delta evenly across the gap months.

### 2. Aggregation columns (Actuals)

| Column | Formula |
|---|---|
| ~~`Sum -i`~~ | **dropped** — collapses into Posted-at-project-end |
| `Posted -i` (through `current_month`) | latest filled cell at-or-before `current_month` |
| `Budget -i` | `=Config!F<r>` (unchanged) |
| `Variance -i` | `Budget − Posted` |
| `% Budget -i` | `Posted / Budget` |
| `Breach-In -i` | `IF(Posted=0,"—",(Budget−Posted)/(Posted/MAX(1,eur_actuals_through_month)))` |

Banned-pattern-safe formula for "latest posted month index" (per spec §"`MATCH`/`LOOKUP` banned"):

```
posted_idx = SUMPRODUCT(MAX(
  (COLUMN(Jan:M_curr) - COLUMN($Jan$1) + 1)
  * ISNUMBER(Jan:M_curr)
))
posted_value = INDEX(row, posted_idx)   ' wrapped in IFERROR(..., "")
```

### 3. Forecast slope (Actuals)

Per-WP slope cell on Actuals (-i side; +i derives via ROUND(×1.25, 0)).

**Single avg-burn formula:**

```
slope    = IF(posted_idx=0, "", posted_value / posted_idx)
forecast = IF(posted_idx=0, "", posted_value + slope * (project_end - posted_idx))
```

**Why avg-burn, not last-pair:** accountant cumulative-input cadence has 2–3 month lag and quarterly catch-up postings. Last-pair slope amplifies catch-up months into runaway forecasts; avg-burn smooths. Acknowledged trade-off: less responsive to a genuine cadence shift mid-project — but smoothing is the correct behavior for this input pattern. One formula across all filled-point counts (0/1/2+); no `prev_*` helpers, no override cell.

### 4. Tab + feature deletions

**A. Plan-Forecast tab deleted entirely.**
- `budget_planner/plan_forecast.py` (462 lines) removed.
- Workbook stage removed from `workbook.py`.
- Three downstream consumers rewired (all sourced from same-row Actuals posted/slope cells, no cross-sheet helper grids):
  - **Actuals forecast monitoring block** — sources from same-row slope, no cross-sheet ref.
  - **Summary `Forecast EUR`** — single per-WP column (not per-year). Inline formula: `posted_value + slope * (project_end - posted_idx)`, clamped at `MAX(0, ...)`. Cross-sheet refs to two Actuals cells per WP. No helper column on Actuals. (See §9.)
  - **Trends `Cumulative Forecast` chart line** — fed by 2 anchor cells per WP: `(posted_idx, posted_value)` now-anchor and `(project_end, forecast_at_project_end)` future-anchor. Chart series interpolates the linear segment between them. No per-month helper grid. (See §9.)

**B. Actuals-Hours forecast-EUR band deleted.** No replacement (WP-level forecast on Actuals is enough). Side benefit: removes the calendar-year-boundary EU-ceiling-CF mis-fire (spec.md §EUR cadence line 49).

**C. EU TOTAL row + EU per-year ceiling CF deleted.** `include_eu_total` opt-in on `person_section.plan_person_layout` becomes dead code; remove the parameter. **Compliance side effect**: live EU per-year ceiling check goes away. User checks compliance manually (or adds a Summary-level check later — out of scope).

**D. Actuals-Hours derived rows deleted.** "Cumulative hours to date" and "Actual PM to date" rows removed (cells themselves now hold posted-cumulative — derived rows duplicate).

**E. Actuals `Sum -i` and `Sum +i` columns dropped** (already covered §2).

### 5. Preserved invariants (no behavior change)

- **Trust boundaries**: `eur_actuals_through_month` / `hours_actuals_through_month` named ranges unchanged at `Actuals!$D$1` and `'Actuals-Hours'!$D$1`. C1 auto-derivation = "largest month-index where ANY WP has a numeric value" (TOTAL + external rows excluded). Same scan as today, no rename. **Semantic axis**: sheet-wide reconciliation cutoff (max across WPs); distinct from per-WP `posted_idx` (this WP's latest filled cell).
- **Inactive-month rule** unchanged.
- **±i mirror convention** unchanged. New Posted -i is accountant-native source; Posted +i = ROUND(-i × 1.25, 0) per cell.
- **Font color convention (Rule 3)** unchanged: blue = blue input (posted-cumulative cells), black = formula (posted_idx, posted_value, slope, forecast columns), green = cross-sheet ref (Budget pulls from Config).

### 6. Migration (manual, one-time)

User clears all Actuals / Actuals-Hours data cells, re-enters as posted-cumulative totals, runs rebuild. No automated migration warning, no schema version stamp, no reader-side branch. Sole-user workbook; the one-time mental shift is documented in the spec.md prose update + this doc.

The reader's existing "largest 1-based month-index with any numeric value" trust-boundary scan returns the correct posted-through-month under both old and new conventions, so the reader is untouched.

### 7. Renames

| Current | New |
|---|---|
| Column header `Cum -i` / `Cum +i` | `Posted -i` / `Posted +i` |
| Forecast column `Cum forecast at project end -i` | `Forecast -i` |
| Helper cells `latest_idx`, `latest_value` | `posted_idx`, `posted_value` |
| Sheet header `As of:` (Actuals + Actuals-Hours) | `Posted through:` |
| Spec.md prose "cumulative-to-date" / "running cumulative" | "posted-cumulative" / "posted through month" |

Named ranges `eur_actuals_through_month` / `hours_actuals_through_month` kept unchanged — accurate (sheet-wide reconciliation cutoff), stable, no semantic gain in renaming, and a different concept from per-WP `posted_*`.

## Key decisions log

1. **Posted-cumulative semantics on per-WP/per-person cells** (not new dedicated rows). One source of truth, no schema duplication.
2. **Avg-burn slope** for forecast: `posted_value / posted_idx`. Single formula across all filled-point counts. Smooths quarterly catch-up postings instead of amplifying them. No override cell (revisit only if a real anomaly fires).
3. **WP-level forecast only.** Drop person-level EUR forecast surface.
4. **±i mirror complexity preserved** as-is.
5. **EU TOTAL + EU per-year ceiling CF deleted.** User accepts manual compliance check.
6. **Manual one-time migration** — no schema stamp, no reader warning, no version branch.
7. **"Posted" lifecycle vocabulary end-to-end on per-WP surface** — column headers, internal helpers, and sheet header all use `Posted` / `Posted through`. Accountant-native action verb (you "post" entries to the books). Sheet-wide named ranges keep `..._through_month` (different semantic axis: sheet-wide reconciliation). Three temporal concepts → three distinct words: `current_month` / `..._through_month` / `posted_*`.
8. **Trends forecast line** fed by 2 anchor cells per WP; no per-month helper grid.
9. **Summary `Forecast EUR`** is a single per-WP column; no per-year breakout (the previous per-year formula would have flat-lined the posted value into closed past years — semantically wrong).

## Out of scope

- ±i mirror consolidation / removal.
- Per-person EUR forecast surface (intentionally dropped).
- Replacement EU-compliance check on Summary.
- Reader "hard-fail until cells cleared" mode (manual migration is sufficient).
- Inactive-month greying rule changes.
- Trends chart restyling beyond rewiring data sources.

## Risks + pre-mortem mitigations

| Risk | Likelihood | Impact | Mitigation in design |
|---|:-:|:-:|---|
| User retains monthly-delta mental model after migration | Med | High | spec.md prose rewrite; explicit one-step manual migration in §6. |
| Avg-burn slope under-responds to genuine cadence shift | Low | Med | Acceptable for accountant cumulative-input cadence with quarterly catch-ups; revisit only if real-world data shows the failure mode. |
| EU compliance check loss goes unnoticed | Med | High | Documented in §4.C as accepted side effect; user explicitly opted in. |
| Variance/breach formulas regress when Posted source changes | Low | Med | Snapshot tests on demo workbook before/after (impl-time). |
| ±i mirrors fed by stale -i source after rewire | Low | Med | Audit every ROUND(×1.25) callsite during the builder rewrite. |
| Reader's "through-month" scan needs different formula | Low | High | Verified: same scan works under both old and new conventions. |
| Reader confuses sheet-wide `..._through_month` with per-WP `posted_*` | Low | Med | Distinct vocabulary (Decision 7); §5 explicitly notes the semantic-axis difference. |

## Resolved layout decisions

### 8. Slope column placement (Actuals)

`Sum -i` and `Sum +i` columns are dropped. Replaced 1-for-1 with the new helpers; net column delta is 0.

Per-WP row, -i side, in display order:

```
Budget -i | Posted -i | Slope -i | Variance -i | % Budget -i | Breach-In -i
                          ↑
                    black-formula (hidden via outline group)
                    = IF(posted_idx=0, "", posted_value / posted_idx)
```

+i mirror block: same shape, with `Slope +i = ROUND(Slope -i × 1.25, 0)`.

`posted_idx` and `posted_value` live as two hidden helper cells per WP-row (outline-grouped, black-formula). Per the spec preference for outline groups over `hidden=True` (`openpyxl.md` §Core patterns).

### 9. Forecast feeding Trends + Summary

**Trends chart — 2 anchor cells per WP, no helper grid.**

Trends owns two black-formula cells per WP that the chart series consumes:

```
anchor_now    : (posted_idx,   posted_value)                                          ' cross-sheet refs to Actuals
anchor_future : (project_end,  posted_value + slope * (project_end - posted_idx))     ' cross-sheet refs to Actuals
```

Chart series has 2 data points per WP for the forecast line. Excel/LO interpolates the linear segment between them. No per-WP × per-month helper grid.

**Why anchors not a per-month grid:** the forecast is a linear function of `m`; a line plotted from N evenly-spaced points equals the line plotted from its 2 endpoints. The grid was solving a non-problem.

**Summary — single `Forecast EUR` column per WP.**

One inline formula per WP row (no helper column on Actuals):

```
=IF(Actuals!<posted_idx>=0, "",
    MAX(0, Actuals!<posted_value> + Actuals!<slope> * (project_end - Actuals!<posted_idx>)))
```

Per-year breakout dropped: the previous per-year formula would have flat-lined the posted value into closed past years (semantically wrong — closed years should reflect actual EOY closed values, not a current-as-of projection). Single end-of-project projection per WP is the honest signal.

## Implementation plan

Phasing, module-by-module re-work audit, and test inventory: separate doc when impl starts. Standard project gating applies (plan-hygiene + pytest gate per `.claude/rules/plan-hygiene.md`).
