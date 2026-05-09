# Actuals "as-of" cumulative semantics — design

**Date:** 2026-05-09
**Status:** Design approved, awaiting implementation plan
**Scope:** Actuals + Actuals-Hours input convention shift; deletes Plan-Forecast tab, Actuals-Hours forecast-EUR band, EU TOTAL row, "Cumulative hours to date" row, "Actual PM to date" row.

## Goals

1. Make accountant-driven cumulative-to-date EUR snapshots the native input convention. Allow gaps in monthly data (snapshots arrive irregularly).
2. Derive the per-WP forecast automatically from the cumulative-input slope, removing the need for a separate hours-derived forecast pipeline.
3. Cut the derived/redundant surface area: one source of truth per WP × month for actuals, one slope cell per WP for forecast, one tab fewer.
4. Preserve `±i` indirect-mirror convention and the inactive-month greying rule unchanged.

## Chosen approach: Level C (maximal simplification) + EU TOTAL removal

Rejected alternatives:
- **Level A (semantic shift only)** — keeps Plan-Forecast and the Actuals-Hours forecast-EUR band even though both become redundant. Two surfaces computing the same forecast; not worth it.
- **Level B (delete Plan-Forecast only)** — keeps Actuals-Hours forecast-EUR band for person-level EUR audit. User explicitly said WP-level forecast is enough; the band carries no extra signal.
- **±i mirror simplification** — explicitly out of scope (user choice).

## Design details

### 1. Input convention

Both actuals tabs hold **as-of values per cell**.

| Tab | Cell convention |
|---|---|
| `Actuals` | `M[wp, m]` = WP total EUR as of month `m` (`-i` source; `+i` derived) |
| `Actuals-Hours` | `H[person, wp, m]` = person-on-WP total hours as of month `m` |

Rules:
- Blank cell = "no data point at this month." Not zero.
- A filled cell is the running total as of that month, regardless of neighbors.
- Monotonic-non-decreasing within a row's active window is expected but **not enforced** (correction postings can go backward).
- The latest filled cell at-or-before `current_month` IS the as-of total.
- Inactive-month greying unchanged. Filled cells inside the active window can have gaps.
- ±i mirrors unchanged: `+i = ROUND(-i × 1.25, 0)` per cell.

**Worked example, WP1:**

```
       Jan | Feb | Mar | Apr | May | Jun ... Dec
WP1          200   380         500
              ↑     ↑           ↑
              as-of as-of       as-of (latest)
              Feb   Mar         May = €500
```

Per-month spend (derived, never stored): Mar = 380 − 200 = 180. May = 500 − 380 = 120. Gap-spreading: between two filled points, the slope `(cum_n − cum_prev) / (m_n − m_prev)` spreads the delta evenly across the gap months.

### 2. Aggregation columns (Actuals)

| Column | Formula |
|---|---|
| ~~`Sum -i`~~ | **dropped** — collapses into As-of-at-project-end |
| `As-of -i` (through `current_month`) | latest filled cell at-or-before `current_month` |
| `Budget -i` | `=Config!F<r>` (unchanged) |
| `Variance -i` | `Budget − As-of` |
| `% Budget -i` | `As-of / Budget` |
| `Breach-In -i` | `IF(As-of=0,"—",(Budget−As-of)/(As-of/MAX(1,eur_as_of_month)))` |

Banned-pattern-safe formula for "latest filled month index" (per spec §"`MATCH`/`LOOKUP` banned"):

```
as_of_idx = SUMPRODUCT(MAX(
  (COLUMN(Jan:M_curr) - COLUMN($Jan$1) + 1)
  * ISNUMBER(Jan:M_curr)
))
as_of_value = INDEX(row, as_of_idx)   ' wrapped in IFERROR(..., "")
```

### 3. Forecast slope (Actuals)

Per-WP slope cell on Actuals (-i side; +i derives via ROUND(×1.25,0)):

| Filled cumulative points | Slope source | Forecast at project end |
|---|---|---|
| 0 points | blank — no forecast | forecast cells display "—" |
| 1 point | origin-anchored: `as_of_value / as_of_idx` | `as_of_value × proj_end / as_of_idx` |
| 2+ points | last-pair: `(as_of_value − prev_value) / (as_of_idx − prev_idx)` | `as_of_value + slope × (proj_end − as_of_idx)` |

**Per-WP manual override cell** (blue-input, optional). When filled, dominates the auto-slope. Catches one-off Q4 catch-up postings that would otherwise produce a runaway forecast.

`prev_idx` formula (banned-pattern-safe):

```
prev_idx = SUMPRODUCT(MAX(
  (COLUMN(Jan:Dec) - COLUMN($Jan$1) + 1)
  * ISNUMBER(Jan:Dec)
  * ((COLUMN(Jan:Dec) - COLUMN($Jan$1) + 1) < as_of_idx)
))
```

### 4. Tab + feature deletions

**A. Plan-Forecast tab deleted entirely.**
- `budget_planner/plan_forecast.py` (462 lines) removed.
- Workbook stage removed from `workbook.py`.
- Three downstream consumers rewired (all pulling from Actuals):
  - **Actuals forecast monitoring block** — sources from same-row slope, no cross-sheet ref.
  - **Summary `Forecast EUR` per-year column** — reads from a new per-year forecast helper column on Actuals (one cell per WP × year).
  - **Trends `Cumulative Forecast` chart line** — reads from a new per-month forecast helper block on Actuals (one row per WP, per-month projected values clamped at `[as_of_idx+1, project_end]`).

**B. Actuals-Hours forecast-EUR band deleted.** No replacement (WP-level forecast on Actuals is enough). Side benefit: removes the calendar-year-boundary EU-ceiling-CF mis-fire (spec.md §EUR cadence line 49).

**C. EU TOTAL row + EU per-year ceiling CF deleted.** `include_eu_total` opt-in on `person_section.plan_person_layout` becomes dead code; remove the parameter. **Compliance side effect**: live EU per-year ceiling check goes away. User checks compliance manually (or adds a Summary-level check later — out of scope).

**D. Actuals-Hours derived rows deleted.** "Cumulative hours to date" and "Actual PM to date" rows removed (cells themselves now hold cum-to-date — derived rows duplicate).

**E. Actuals `Sum -i` and `Sum +i` columns dropped** (already covered §2).

### 5. Preserved invariants (no behavior change)

- **Trust boundaries**: `eur_as_of_month` / `hours_as_of_month` named ranges still live at `Actuals!$D$1` and `'Actuals-Hours'!$D$1`. C1 auto-derivation = "largest month-index where ANY WP has a numeric value" (TOTAL + external rows excluded). Same scan as today; renamed.
- **Inactive-month rule** unchanged.
- **±i mirror convention** unchanged. New As-of -i is accountant-native source; As-of +i = ROUND(-i × 1.25, 0) per cell.
- **Font color convention (Rule 3)** unchanged: blue = blue input (as-of cells), black = formula (as_of_idx, as_of_value, slope, forecast columns), green = cross-sheet ref (Budget pulls from Config).

### 6. Reader and rebuild contract — migration

**Migration trap:** existing workbooks have monthly-delta values in cells that will be reinterpreted as cumulative. Rebuild MUST NOT silently reinterpret.

Plan:
1. **No automatic migration.** Rebuild treats existing cells as opaque — preserved verbatim.
2. **Schema version stamp** bumped to `actuals-as-of-v1` (extends the existing version stamp on Config).
3. **One-time warning on rebuild** when loading a pre-stamp workbook: "Existing Actuals/Actuals-Hours cells used monthly-delta convention. New convention is as-of cumulative. Clear cells and re-enter as as-of totals. Rebuild has preserved your numbers but they no longer have a valid interpretation." Stamp written on first rebuild.
4. **Preserved-data table** (spec.md §Reader and Rebuild contract → "What rebuild preserves") updated:
   - Removed: "EU TOTAL row" entry (cells deleted).
   - Removed: "Forecast-EUR band cells" entry (deleted).
   - Removed: "Plan-Forecast cells" entry (tab deleted).
   - Renamed: "Monthly EUR values per WP" → "As-of EUR values per WP". "Hour timelog per person × WP × month" → "As-of hours per person × WP × month".
5. **Reader extraction functions:** unchanged. The current "largest 1-based month-index with any numeric value" scan happens to correctly return the as-of-month under both old and new conventions.

### 7. Renames everywhere

| Current | New |
|---|---|
| Column header `Cum -i` / `Cum +i` | `As-of -i` / `As-of +i` |
| Forecast column `Cum forecast at project end -i` | `Forecast -i` |
| Helper cells `latest_idx`, `latest_value` | `as_of_idx`, `as_of_value` |
| Named range `eur_actuals_through_month` | `eur_as_of_month` |
| Named range `hours_actuals_through_month` | `hours_as_of_month` |
| Spec.md prose "cumulative-to-date" / "running cumulative" | "as-of total" / "as-of-month value" |

## Key decisions log

1. **Cumulative-to-date semantics on per-WP/per-person cells** (not new dedicated rows). One source of truth, no schema duplication.
2. **Last-pair slope** for forecast (recent-burn, two filled points). Origin-anchored fallback for 1 point. Blank for 0 points.
3. **Per-WP manual override cell** for slope. Pre-mortem mitigation against Q4 catch-up posting volatility.
4. **WP-level forecast only.** Drop person-level EUR forecast surface.
5. **±i mirror complexity preserved** as-is.
6. **EU TOTAL + EU per-year ceiling CF deleted.** User accepts manual compliance check.
7. **No automatic migration** — warn-don't-migrate. User clears cells and re-enters once.
8. **Renames consolidated to "As-of"** vocabulary (matches existing "As of:" header label on Actuals; accountant-native; time-anchored).

## Out of scope

- ±i mirror consolidation / removal.
- Per-person EUR forecast surface (intentionally dropped).
- Replacement EU-compliance check on Summary.
- Reader "hard-fail until cells cleared" mode (warning is sufficient).
- Inactive-month greying rule changes.
- Trends chart restyling beyond rewiring data sources.

## Risks + pre-mortem mitigations

| Risk | Likelihood | Impact | Mitigation in design |
|---|:-:|:-:|---|
| User retains monthly-delta mental model after migration | Med | High | Stamp + console warning on first rebuild (§6); spec.md prose rewrite. |
| Last-pair slope volatile on anomalous month | Med | Med | Per-WP manual override cell (§3). |
| Trends/Summary lose intra-year forecast granularity | Low | Med | Per-month forecast helper block on Actuals feeds Trends; per-year helper column feeds Summary (§4.A). |
| EU compliance check loss goes unnoticed | Med | High | Documented in §4.C as accepted side effect; user explicitly opted in. |
| Variance/breach formulas regress when As-of source changes | Low | Med | Snapshot tests on demo workbook before/after (Phase 6). |
| ±i mirrors fed by stale -i source after rewire | Low | Med | Audit every ROUND(×1.25) callsite during Phase 2/3 (Phase B in plan). |
| Reader's "through-month" scan needs different formula | Low | High | Verified: same scan works under both old and new conventions. |

## Phasing (re-work-audited per skill step 7.5)

| Phase | Scope | Modules touched |
|---|---|---|
| 1 | Spec.md + named-range rename + schema version bump | `docs/spec.md`, `budget_planner/workbook.py` (named-range registration), `config_sheet.py` (version stamp), `layout.py` if column constants shift |
| 2 | Actuals + Summary + Trends — full rewrite | `actuals.py`, `summary.py`, `trends.py` |
| 3 | Actuals-Hours — full rewrite | `actuals_hours.py`, `person_section.py` (drop `include_eu_total` param) |
| 4 | Delete Plan-Forecast tab | rm `plan_forecast.py`; `workbook.py` (remove stage) |
| 5 | Reader migration warning | reader/extractor module, `__main__.py` (rebuild entry) |
| 6 | Tests + manual demo verification | `tests/**`, demo workbook |

**Re-work audit:** no module is rewritten in two phases.
- `actuals.py`: Phase 2 only.
- `actuals_hours.py`: Phase 3 only.
- `workbook.py`: Phase 1 (named-range registration) + Phase 4 (Plan-Forecast stage removal) — distinct concerns, not re-write.
- `summary.py`, `trends.py`, `person_section.py`: each touched once.
- `spec.md`: Phase 1 only (terminology + named-range + preserved-table updates land together).

## Migration UX (user-facing)

1. Pull this branch.
2. Run rebuild on existing workbook → console warning fires.
3. Open workbook, clear Actuals data cells, clear Actuals-Hours data cells.
4. Re-enter latest accountant snapshot as as-of totals (one cell per WP × month for accountant data; gaps where no snapshot exists).
5. Re-enter Actuals-Hours as cumulative hours per person × WP × month (gaps OK).
6. Run rebuild again → schema stamp written; warning suppressed on subsequent rebuilds.

## Resolved layout decisions

### 8. Slope + override column placement (Actuals)

`Sum -i` and `Sum +i` columns are dropped. Replaced 1-for-1 with the new helpers; net column delta is +1.

Per-WP row, -i side, in display order:

```
Budget -i | As-of -i | Override -i | Slope -i | Variance -i | % Budget -i | Breach-In -i
                           ↑           ↑
                       blue input    black-formula (hidden via outline group)
                       (optional)    = IF(Override<>"", Override,
                                          IF(as_of_idx=0, "",
                                          IF(prev_idx=0, as_of_value/as_of_idx,
                                             (as_of_value-prev_value)/(as_of_idx-prev_idx))))
```

+i mirror block: same shape, with `Slope +i = ROUND(Slope -i × 1.25, 0)`. **No `Override +i` cell** — `+i` is a pure derivation. If the user wants to override, they do it on -i; +i follows.

`as_of_idx`, `as_of_value`, `prev_idx`, `prev_value` live as four hidden helper cells per WP-row (outline-grouped, black-formula). Per the spec preference for outline groups over `hidden=True` (`openpyxl.md` §Core patterns).

### 9. Per-month forecast block — lives on Trends, not Actuals

Trends chart owns its own per-WP × per-month forecast helper rows. Each cell:

```
=IF(m <= as_of_idx, "", as_of_value + slope * (m - as_of_idx))
```

References cross-sheet to `Actuals!<as_of_value>` and `Actuals!<slope>` (green-xref font). Clamped at project end by the column range itself.

**Why on Trends and not Actuals:** Actuals stays focused on actuals + single-cell slope. The chart-feeding grid is a Trends concern; placing it on Actuals would interleave forecast values with the input grid and re-introduce the "two surfaces" problem the simplification is meant to remove.

Summary's per-year `Forecast EUR` column reads inline (no helper column on Actuals): `=as_of_value + slope * (year_end_month - as_of_idx)`, clamped at `MAX(0, ...)` so a 0-point WP shows blank, not negative. One formula per WP × year cell, sourced directly from the two Actuals helper cells.

### 10. Schema version stamp — new named range on Config

Single new named range `actuals_schema` pointing to a fresh Config cell (sibling of the existing `version_stamp`, black-formula effective cell with literal value `"as-of-v1"`). Reader checks existence at rebuild start:

- Named range absent OR value ≠ `"as-of-v1"` → emit one-time console warning, then write the stamp at end of rebuild.
- Named range present with `"as-of-v1"` → silent rebuild as today.

**Why a new range, not extending `version_stamp`:** keeping `version_stamp` as a free-form human-readable string means parser logic touches a string-format. A dedicated boolean-ish named range is one `wb.defined_names` lookup. Lower-effort, lower-risk.

### 11. Test inventory

**Delete** (~7 files / ~2k LOC removed):
- `tests/test_plan_forecast*.py` — entire tab gone.
- `tests/test_actuals_hours_forecast_eur*.py` — band gone.
- `tests/test_actuals_hours_eu_total*.py` — row + CF gone.
- `tests/test_actuals_hours_cumulative_to_date*.py` — derived row gone.
- `tests/test_actuals_hours_actual_pm*.py` — derived row gone.
- `tests/test_actuals_sum_column*.py` (or whichever file targets `val_col_sum`) — column gone.

**Update** (existing tests, semantic + rename adjustments):
- `tests/test_actuals*.py` — change expected values to cumulative semantics; rename `Cum` → `As-of`, named-range references; update breach/variance/forecast-block source assertions to point at same-row slope instead of Plan-Forecast.
- `tests/test_actuals_hours*.py` (surviving files) — rename + cumulative semantics.
- `tests/test_summary*.py` — `Forecast EUR` source switch from Plan-Forecast to inline-from-Actuals.
- `tests/test_trends*.py` — chart series source switches to Trends-local helper rows referencing Actuals slope.
- `tests/test_reader*.py` — `eur_as_of_month` / `hours_as_of_month` named-range names; preserved-data assertions.
- Test fixtures in `tests/fixtures/` (or sample workbooks): regenerate with cumulative-to-date sample data.

**Add** (new files):
- `tests/test_actuals_as_of.py` — covers: latest-filled-month formula (banned-pattern-safe form), 0/1/2+ point slope branches, override cell dominance, +i mirror = ROUND(-i × 1.25, 0).
- `tests/test_actuals_forecast_block.py` — forecast columns (Forecast, Variance forecast, % Budget forecast, Breach-In forecast) sourced from same-row slope; correct values for the worked example in §1.
- `tests/test_reader_schema_stamp.py` — warning fires once on pre-stamp workbook; stamp written; subsequent rebuild silent.

Specific filename choices and dropped-test list confirmed via `grep -rn "Plan-Forecast\|val_col_sum\|EU TOTAL\|forecast_val_columns" tests/` during Phase 6.
