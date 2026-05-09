# Actuals posted-cumulative — implementation plan

**Date:** 2026-05-09
**Status:** Draft (Phase 1 spec rewrite already shipped — commit `06f122e`)
**Design ref:** `docs/plans/2026-05-09-actuals-posted-cumulative-design.md`

Implements the design's §"Implementation plan" deferred work: phasing, module-by-module re-work audit, test disposition. Spec is already ahead of code (§"Phase 1 SHIPPED" in the design doc) — phases below bring code to match.

## Phase ordering — dependency-driven, not nominal

Design lists Phase 2→3→4→6. Re-ordering here for dependency cleanliness:

1. **Phase 2** — Actuals + Summary + Trends rewrite. Decouples the 3 Plan-Forecast consumers; once shipped, Plan-Forecast is computed but no consumer reads it.
2. **Phase 4** — delete Plan-Forecast. Mechanical now (no consumers); also lets us delete `plan_forecast_layout=` kwargs from `build_actuals_sheet` / `build_trends_sheet` / `build_summary_sheet` cleanly.
3. **Phase 3** — Actuals-Hours rewrite. Independent of 2/4 (Actuals-Hours feeds Plan-Forecast, but Plan-Forecast is gone). Doing this last avoids touching 2 sheets in flight.
4. **Phase 6** — test sweep + demo regen + manual verification.

Phase 5 (reader migration) is dropped per the design doc's amendment.

## Stage boundaries (for `/impag` resumption)

| Stage | Phases | Rationale |
|---|---|---|
| **Stage A** | Phase 2 (Actuals rewrite only) | Largest single module (1042 LOC); ship + commit + retro before continuing. |
| **Stage B** | Phase 2 (Summary + Trends rewiring) | Thinner — wires Summary/Trends directly to Actuals' new `posted_*` / `slope` cells. |
| **Stage C** | Phase 4 (delete Plan-Forecast) | Mechanical delete of file, sheet stage, layout kwargs, cross-refs, tests. |
| **Stage D** | Phase 3 (Actuals-Hours rewrite) | Self-contained rewrite + `include_eu_total` kwarg drop. |
| **Stage E** | Phase 6 (tests + demo + manual) | Final sweep. |

Each stage = a fresh `/impag` session. Don't pile stages — the orchestrator runs out of context budget.

## Phase 2 — Actuals + Summary + Trends

### Phase 2.1 — `budget_planner/layout.py`

| Edit | Detail |
|---|---|
| Rename `VAL_OFFSET_CUM` → `VAL_OFFSET_POSTED` | line 176; same value (4); cum→posted across this module |
| Rename `VAL_OFFSET_CUM_PLUS_I` → `VAL_OFFSET_POSTED_PLUS_I` | line 186; same value (10) |
| Add `VAL_OFFSET_SLOPE` / `VAL_OFFSET_SLOPE_PLUS_I` | new offsets — replace the dropped `sum`/`sum_plus_i` slots (1, 7) so column count stays unchanged. Net column delta = 0 per design §8. |
| Drop `VAL_OFFSET_SUM` / `VAL_OFFSET_SUM_PLUS_I` | not needed — Posted absorbs the role |
| Add `VAL_OFFSET_POSTED_IDX` / `VAL_OFFSET_POSTED_VALUE` (-i side, hidden helpers) | two new offsets at the end of the -i band; outline-grouped per design §8; +i side does NOT need helper mirrors (Posted +i = ROUND(Posted -i × 1.25, 0) reads the -i Posted directly) |
| Update `val_columns()` keys | `sum`→removed; add `slope`, `slope_plus_i`, `posted_idx`, `posted_value`; rename `cum`→`posted`, `cum_plus_i`→`posted_plus_i` |
| Forecast-block `forecast_val_columns` | DELETE entirely. Forecast monitoring block on Actuals is gone — replaced by per-WP `slope`/`forecast` cells inline in the -i + +i rollup blocks. |

**Test impact:** `tests/test_layout.py` (assertions on offset names/values), every test referencing `val_col_sum*` / `val_col_cum*` / `forecast_val_columns`. Estimate: 15-25 test edits.

### Phase 2.2 — `budget_planner/actuals.py` rewrite

The 1042-LOC builder currently has these helpers (line numbers from current HEAD):

| Helper | Disposition |
|---|---|
| `_resolve_actuals_layout` (52) | Rewrite — drop `plan_forecast_layout` param; drop forecast-block keys; add `posted_idx_col`/`posted_value_col`/`slope_col` (& `_pi`); `_minus_i_band_cols` recomputed |
| `_write_asof_header_row` (138) | Rename label `"As of:"` → `"Posted through:"`; drop the duplicate at `col_cum_forecast` / `col_cum_forecast_pi` (forecast block deleted) |
| `_write_section_headers` (175) | Drop FORECAST MONITORING -i / +i banner merges; ACTUAL EUR -i / +i now spans the wider rollup block (Posted, Slope, Variance, %, Breach-In) |
| `_apply_header_fill_overrides` (207) | Update column lists to new -i/+i column ids |
| `_write_wp_rows` (249) | Major rewrite (covered below) |
| `_write_total_row_*` (471, 484, 533, 579, 626) | Rewrite -i + +i (drop forecast variants entirely) |
| `_write_cum_grey_cf` / forecast CF helpers | Drop forecast-band CF; keep monthly-grid grey-after-`eur_actuals_through_month` rule |

**`_write_wp_rows` rewrite plan** — per WP row:

```
-i block (collapsible, hidden by default):
  monthly cells (cols C..month_col_end)        [BLUE_INPUT, posted-cumulative]
  Posted -i  (was Cum -i)                      = IFERROR(INDEX(row, posted_idx), "")
                                                 — uses hidden posted_idx helper
  Slope -i  (NEW — replaces Sum -i)             = IF(posted_idx=0, "", posted_value/posted_idx)
  Budget -i                                     unchanged
  Variance -i                                   = Budget - Posted
  % Budget -i                                   = Posted / Budget
  Breach-In -i                                  unchanged shape

+i block (default-visible):
  Posted +i  = ROUND(Posted_-i × 1.25, 0)
  Slope +i   = ROUND(Slope_-i × 1.25, 0)
  Budget +i  unchanged
  Variance/% / Breach-In +i  same shape as -i

hidden helpers (outline-grouped, BLACK_FORMULA — design §8):
  posted_idx   = SUMPRODUCT(MAX((COLUMN(C{r}:M_curr)-COLUMN($C$1)+1)*ISNUMBER(C{r}:M_curr)))
                  where M_curr = column at index `current_month`
  posted_value = IFERROR(INDEX(row, posted_idx), "")
```

Key formula gotchas (per `.claude/rules/openpyxl.md`):
- `posted_idx` formula uses SUMPRODUCT-MAX-array form (banned-pattern-safe — design §2 quotes it verbatim).
- `posted_value` wraps `INDEX` in `IFERROR(..., "")` — empty string sentinel surfaces "no posting yet" to downstream consumers.
- Forecast monitoring block CFs (cum_forecast/var_forecast/pct_ceil_forecast/breach_forecast) — DELETE all corresponding `ws.conditional_formatting.add` calls.

**Drops:**
- All `plan_forecast_layout` references (param, dict keys, `_PF_QUOTED`, `pf_first_wp_row`, `pf_total_row`, `_PF_QUOTED + pf_range` formulas at lines 396-465 + TOTAL-row mirrors at 579-660+).
- Sum -i / Sum +i columns and headers.
- All forecast-monitoring formulas + section headers.

**Test impact:** `tests/test_actuals.py` (372L), `tests/test_actuals_breach_todate.py` (186), `tests/test_actuals_forecast_block.py` (251 — likely most of file deleted), `tests/test_indirect_view_invariants.py`, header-row tests in `test_actuals.py` (column-name asserts on `Sum +i`, `Cum +i`, etc.).

### Phase 2.3 — `budget_planner/summary.py`

Single section to rewire: `_write_wp_data_rows` (line 205).

Currently emits 5 columns per year per WP (Plan FTE, Plan EUR, Actual EUR, **Forecast EUR ← Plan-Forecast SUM**, Variance EUR). Per design §9: **drop per-year Forecast EUR**, replace with **single project-end Forecast EUR column** per WP (no year breakout — closed past years would be flat-lined).

Concrete edits:
- `_COLS_PER_YEAR` 5 → 4 (drops Forecast EUR slot).
- Year-group emits Plan FTE / Plan EUR / Actual EUR / Variance EUR (4 cols).
- Add a new "Forecast EUR (project-end)" column to the right of the per-year band, before PROJECT TOTAL. Single column per WP.
- Cell formula per WP row (per design §9):
  ```
  =IF(Actuals!<posted_idx_cell>=0, "",
      MAX(0, Actuals!<posted_value_cell> + Actuals!<slope_cell>
            * (project_months - Actuals!<posted_idx_cell>)))
  ```
  Cross-sheet refs — green. (`project_end` in design = `project_months` named range).
- Drop `plan_forecast_layout` param + `pf_first_wp_row` lookups (line 88, 209, 241-244).
- `_apply_year_forecast_cf` (311) — DELETE (no per-year forecast cells).
- `_apply_partial_coverage_cf` (336) — drop the `fcst_col_l` block (lines 358-362); keep the act + var blocks but switch their `_COLS_PER_YEAR` math to the new value.
- `_apply_total_forecast_cf` (374) — REPLACE with a single-cell rule on the new project-end Forecast EUR column (red if <0.8 × project total Plan EUR).

**Test impact:** `tests/test_summary.py` (`test_summary_forecast_eur_formulas_reference_plan_forecast` at line 71 — REPLACE with `references_actuals_posted_value` equivalent), `tests/test_summary_partial_year.py`.

### Phase 2.4 — `budget_planner/trends.py`

Forecast-EUR series + reconciliation section both consume Plan-Forecast.

Concrete edits:
- `_resolve_trends_layout` (53) — drop `plan_forecast_layout` param + downstream layout keys (`pf_first_wp_row` at 115, `pf_total_row` at 120; the None-fallback ternary becomes dead code).
- `_write_chart_helper_rows` (625) — currently: helper row 3 sums Plan-Forecast monthly TOTAL, helper rows feed the chart. Replace with **2-anchor-cells-per-WP** approach (design §9):
  - For each WP, 2 cells: `(posted_idx, posted_value)` now-anchor and `(project_months, posted_value + slope*(project_months − posted_idx))` future-anchor. Cross-sheet GREEN to Actuals.
  - Chart series for "Cumulative forecast" line: 2 data points per WP (or aggregate to project-level: 2 totals — sum of now-anchors, sum of future-anchors).
  - Decision point during impl: per-WP 2-anchor series (multi-line chart, busier) vs. project-aggregated 2-anchor (single forecast line, simpler). Design §9 says "the chart series consumes" — leans toward project-aggregated; verify against current chart legend before choosing.
- Drop the "Cumulative forecast EUR" helper row that SUMs `'Plan-Forecast'!<col><total_row>` (currently line 693-700 area).
- `_write_reconciliation_section` (378-510 area) — Col B "Cum Forecast EUR" currently SUMs `'Plan-Forecast'!` (line 458-484). Replace with per-WP slope projection: `MAX(0, Actuals!posted_value + Actuals!slope*(reconciled_through_month - Actuals!posted_idx))`.
- Remove `SHEET_PLAN_FORECAST` import (line 32) once all formulas re-pointed.

**Test impact:** `tests/test_trends.py` (820L — multiple `'Plan-Forecast'` formula assertions at 609-755), `tests/test_trends_reconciliation.py` (143-149, 370 — Col B Plan-Forecast reference).

### Phase 2.5 — `budget_planner/workbook.py` (transitional)

Phase 2 doesn't delete Plan-Forecast yet (Phase 4 does). But Plan-Forecast layout is no longer consumed:
- Keep `plan_forecast_layout = build_plan_forecast_sheet(...)` call at line 412.
- Drop `plan_forecast_layout=plan_forecast_layout` from `build_actuals_sheet` (429), `build_trends_sheet` (442), `build_summary_sheet` (456) calls. The 3 builders' signatures already lost the param in Phase 2.1-2.4.
- Drop wiring docstring stubs at lines 384-385, 422 referring to "consume Plan-Forecast's layout".

Verify the workbook still builds (Plan-Forecast renders but feeds nothing) before commit.

## Phase 4 — delete Plan-Forecast

Mechanical delete now that no consumer remains.

| Edit | Detail |
|---|---|
| `budget_planner/plan_forecast.py` | `git rm` |
| `budget_planner/layout.py` | drop `SHEET_PLAN_FORECAST = "Plan-Forecast"` (line 42) |
| `budget_planner/workbook.py` | drop `from budget_planner.plan_forecast import build_plan_forecast_sheet` (93), `SHEET_PLAN_FORECAST` import (103), Stage 6b block (lines 409-415), plan_forecast_layout var; update build-order docstring (38, 57-61). |
| `budget_planner/actuals_hours.py` comments | line 235 "Plan-Forecast SUMs match", 749/774 "Plan-Forecast" mentions — re-word to remove dangling refs. |
| `budget_planner/config_sheet.py` | line 853 ("Plan-Forecast"), 964 (comment) — re-word. |
| `tests/test_plan_forecast.py` | `git rm` (498 LOC) |
| `tests/test_actuals_forecast_block.py` | likely `git rm` (251 LOC — entire file is forecast-monitoring block tests) |
| `tests/test_workbook.py` | line 29 (sheet list), 283 (cross-ref dict) — drop "Plan-Forecast" entries |
| `tests/test_indirect_view_invariants.py` | drop `SHEET_PLAN_FORECAST` from sheet list (37, 51) |
| `tests/test_rebuild_hours_forecast_e2e.py` | this file's whole purpose is the Actuals-Hours → Plan-Forecast leaf-aggregation invariant. Either `git rm` or rescope to "rebuild-hours preserves Actuals-Hours data" only. Recommend `git rm` and rely on existing `test_rebuild_actuals_hours.py` (the leaf invariant goes away with Plan-Forecast). |
| `tests/test_actuals_hours.py` line 337 | drop sub-cent comment |
| `tests/test_summary.py` | `_find_plan_forecast_first_wp_row` helper (40-45) — DELETE; `test_summary_forecast_eur_formulas_reference_plan_forecast` (71-93) — REPLACE per Phase 2.3 |
| `docs/spec.md` | grep `Plan-Forecast` and remove all references; build graph diagram update |
| `CLAUDE.md` | grep `Plan-Forecast` (mostly in trigger-table content links) |
| `output/budget_plan_demo.xlsx` | regen in Phase 6 |

**Pre-archive sweep**: `grep -rn "Plan-Forecast\|plan_forecast\|SHEET_PLAN_FORECAST" .` — every result must resolve before commit. Current count: ~80 hits across 30+ files.

## Phase 3 — Actuals-Hours rewrite

Independent of Phase 2/4. Drops 4 features in one pass:

| Drop | Where |
|---|---|
| Forecast EUR band (per-month forecast cells per person × WP) | `actuals_hours.py` `forecast_col_start..forecast_col_end`, `forecast_plus_i_col`; person_section.plan_person_layout the band columns |
| EU TOTAL row + per-year ceiling CF | `actuals_hours.py` lines 406-490; `person_section.plan_person_layout` `include_eu_total` param + `eu_total_row` field |
| "Cumulative hours to date" row | `actuals_hours.py` lines 334-351; `plan_person_layout.cum_hours_todate_row` |
| "Actual PM to date" row | `actuals_hours.py` lines 353-360; `plan_person_layout.actual_pm_todate_row` |

Then add posted-cumulative semantics:
- Per-cell convention shifts from "hours-this-month" to "posted-cumulative-hours-through-month" per design §1. Inactive-month grey unchanged. Blank-allowed-inside-active-window now applies.
- Drop the `=SUM(per-month cells)` style TOTAL — TOTAL row's per-month value becomes "latest posted cumulative across data rows for this month" (or just keep SUM and let the user mentally interpret — design doesn't explicitly cover this; defer to user during impl).
- Sheet header `"As of:"` → `"Posted through:"` per design §7. Search `actuals_hours.py` for `"As of:"` literals.
- `hours_actuals_through_month` named range trust-boundary scan unchanged (design §5 — same scan works under both conventions).

**`person_section.plan_person_layout` cleanup** — drop dead code:
- Remove `include_eu_total` param entirely (Capacity already passes False → it's dead in capacity.py:169 too). The dataclass field `eu_total_row: int | None` becomes dead → remove.
- Remove `include_hours_todate_summary` param (was True for Actuals-Hours, no other caller).
- Remove `eu_total_row`, `cum_hours_todate_row`, `actual_pm_todate_row` fields.

**Reader impact** — `budget_planner/readers/actuals_hours.py`:
- Line 30 docstring + line 158 "EU TOTAL" stop-row check — drop EU TOTAL handling.
- Line 155 — TOTAL stop-row remains the sole terminator.
- Verify: under new posted-cumulative cells, the existing reader just returns whatever the user typed. No semantic change.

`budget_planner/readers/capacity.py:222-223` — `("TOTAL", "REMAINING", "EU TOTAL")` → `("TOTAL", "REMAINING")` (EU TOTAL dropped 2026-04-28 already on Capacity, but the reader's tuple still lists it — confirmable dead-list cleanup).

**Test impact:**
- `tests/test_actuals_hours.py` (534L) — header-row asserts (115-127), EU TOTAL formula (216-245), forecast-band tests (test_actuals_hours_forecast_shading.py whole file).
- `tests/test_actuals_hours_todate_rows.py` — DELETE entire file (todate rows removed).
- `tests/test_actuals_hours_forecast_shading.py` — DELETE.
- `tests/test_rebuild_actuals_hours.py` — line 54 "EU TOTAL" stop check.
- `tests/test_person_section.py` (lines 10-47 — `include_eu_total=True/False` cases). Drop both true cases; keep one minimal layout test.
- `tests/test_capacity.py:1020`, `tests/test_capacity_actual_view.py`, `tests/test_capacity_todate.py` — EU TOTAL grep hits, mostly comments.
- `tests/test_rebuild_hours_forecast_e2e.py` — already slated for delete in Phase 4.

## Phase 6 — tests + demo + manual verification

1. **Test sweep checklist** — every file flagged in Phases 2-4 above; count delta should be net -3 to -8 tests (deletions partially offset by new posted-cumulative coverage):
   - NEW: `tests/test_actuals_posted.py` — `posted_idx`/`posted_value` formula shape, slope formula, IFERROR-empty behavior, mid-window-gap handling.
   - NEW: `tests/test_summary_forecast_project_end.py` — single-column project-end Forecast EUR formula references Actuals slope cells.
   - NEW: `tests/test_trends_anchor_cells.py` — 2-anchor chart series shape.
2. **Pyright clean** — `uv run pyright` after each stage's commit.
3. **Demo regen** — `python -m budget_planner` then `python scripts/recalc.py output/budget_plan_demo.xlsx`. Verify pytest still green.
4. **Manual smoke** — open demo workbook in LibreOffice + Excel:
   - Posted +i/-i cells render, gaps allowed inside window.
   - Slope cells black-formula, hidden via outline.
   - Summary Forecast EUR project-end column shows sane numbers.
   - Trends Cumulative-Forecast chart line drawn cleanly between 2 anchors.
   - No dangling Plan-Forecast tab.
5. **CLAUDE.md / spec.md final pass** — sweep `Plan-Forecast` / `Cum -i` / `Sum -i` / "Cumulative hours to date" / "EU TOTAL" / "As of:" — confirm no surviving references on Actuals/Actuals-Hours pages of spec.md (Capacity still has EU TOTAL prose for the Capacity-only context, that stays).

## Risks specific to implementation

| Risk | Mitigation |
|---|---|
| Forecast monitoring block deletion breaks `test_actuals_breach_todate.py` ASSERT-by-position helpers | Rewrite helpers to use new column-key names from `val_columns`; assert on `posted`/`slope` semantics, not `cum_forecast`. |
| Slope formula `posted_value/posted_idx` divides by 0 when no postings | `IF(posted_idx=0, "", …)` guard per design §3 — bake into BOTH slope and forecast cells. |
| Mid-window gaps confuse SUMPRODUCT-MAX scan | Test fixtures with explicit gap patterns: `[200, blank, 380, blank, 500, blank...]` — `posted_idx` should return the LATEST filled index ≤ current_month. |
| Layout-shift breaks hardcoded column counts in tests | `grep -rn "base_cols\s*=\|val_col_sum\|val_col_cum\|forecast_val_columns" tests/` before each commit; mirror the timelog Stop-reason layout-shift gotcha (`memory/project_state.md` § 2026-04-26 stop-reason). |
| `_PERSON_SUMMARY_RATE_COL` analog in `actuals.py` — consumers of `val_col_cum` | Sweep every `val_col_cum`/`val_col_sum` consumer for unit semantics during the rename, mirroring the rate-col-flip lesson (`project_state.md` § 2026-04-26 rate split). |
| Reader misreads posted-cumulative as monthly delta | Per design §6 the existing scan works under both conventions. Spot-check `tests/test_reader_actuals.py` round-trip with posted-cumulative input. |

## Open decisions deferred to implementation

1. Trends "Cumulative forecast" line: per-WP 2-anchor series (multi-line) vs. project-aggregated 2-anchor (single line). Pick during Phase 2.4 by inspecting current chart.
2. Actuals-Hours TOTAL row's per-month value under posted-cumulative semantics: keep `=SUM` (sums posted-cumulative cells across persons → "all-time posted hours through this month, summed across team") or rewrite to `=MAX` semantics. Defer to user during Phase 3.
3. New posted-cumulative test fixture inputs — keep current demo-seed shape (random distribution) or seed an explicit posted-cumulative pattern with a Q1-and-Q3-catchup gap profile. Recommend the explicit pattern so tests pin gap-handling.
