# Timelog Plan enhancements — monthly totals, yearly totals/budget, vacation-month exclusion

**Status:** Validated via brainstorming 2026-04-19. Ready for phased implementation.
**Scope:** Four Timelog Plan-sheet enhancements. Follows v2 (`docs/plans/archived/2026-04-19-timelog-plan-v2-design.md`) which shipped the merged grid + resolve toggle. This doc does not revisit v2's architecture — it extends the tab with totals rows, a per-year budget reconciliation row, drops the `Largest gap` placeholder, and introduces a single shared vacation-month exclusion.

**Not addressed here (separate design doc):** WP gap-logic removal is out of scope — will be tackled in its own brainstorm. This spec is written to remain correct whether gap logic stays or is later removed (gap-aware `_wp_sched_months` is retained; swap-in is internal).

---

## 1. Motivation

After using v2 for one planning cycle the user surfaced four enhancements:

1. **TIMELOG GRID lacks a monthly cost line.** To eyeball "is this month thin?" today you read each row independently. A single EUR-total row beneath the grid closes that gap.
2. **PER-WP SUMMARY needs a yearly reconciliation view.** The `Y{yyyy} €` columns show per-WP × per-year cost, but there's no cross-row sum (total cost per year) and no comparison against per-year budget. Budget reconciliation is the user's primary acceptance gate.
3. **`Largest gap` placeholder still empty, naming now confusing.** Deferred in v1, still deferred in v2, no user ask since. The incoming WP-gap-removal work makes "gap" unambiguously refer to schedule gaps — a person-level metric called `Largest gap` becomes a name hazard.
4. **Vacation months must not be logged to the project.** programme rules forbid logging vacation months. Most people take July. The current tool has no enforcement — the solver can (and does) place WP codes in July.

---

## 2. Goals / Non-goals

### Goals
- TIMELOG GRID: one new **Monthly total €** row directly below the last person row, inside the same outline group.
- PER-WP SUMMARY: two new rows beneath the last WP row — **Yearly total €** and **Yearly budget €** — each with a col-C cross-check and per-year cells.
- Drop `Largest gap` from PER-PERSON SUMMARY; collapse from 9 cols to 8; Rate shifts from col I to col H.
- Introduce a single **Vacation month** Config scalar (default `"Jul"`). Uniform across all persons. Solver never places WP codes in vacation months. Grid columns corresponding to vacation months get a distinct (yellow) fill tier and are locked. Pool formula adjusts to `max_fte × (project_months − num_vacation_months)`.
- Forward-compat: later replacement of the scalar with a per-(person, year) Vacation Schedule table (Option A from brainstorm) must require only additive reader/solver changes — no cell relocation on this sheet.

### Non-goals (YAGNI)
- **Per-(person, year) vacation overrides.** Uniform July is explicitly enough for current use; heterogeneous vacation is deferred to a future design if the need materialises.
- **Vacation reflected on Plan / Capacity / Summary.** Only TIMELOG GRID + PER-PERSON SUMMARY Pool need vacation awareness in this doc. Plan / Capacity already approximate vacation via `max_fte`; extending them is a separate concern.
- **Vacation marker text ("VAC") inside cells.** Yellow fill + locked suffices; adding text clutters the grid and complicates the `grid_range<>""` pattern used by OOW / Assigned € SUMPRODUCTs.
- **Rebuild-time audit of pre-existing WP codes in vacation cells.** Solver prevents on init. On rebuild with `resolve_flag=FALSE`, the preserved grid's cells are written verbatim; if the user had typed a WP code into a vacation cell in a prior version (pre-vacation-awareness), that cell will become locked-blank on first rebuild under this doc. Acceptable one-shot migration.
- **WP gap-logic removal.** Separate design doc.
- **Reinstating `Largest gap` under a new name (`Longest idle run`).** Evaluated during brainstorm; dropped. The monthly-total row will visually surface thin months, which covers 80% of the "lumpy allocation" signal the metric would have given.

---

## 3. Final layout specs

### 3.1 TIMELOG GRID — Monthly total € row

| Row | Col A | Cols B…(B + months − 1) |
|---|---|---|
| grid_total_row | `Monthly total €` (label, BLACK_FORMULA) | `=ROUND(SUMPRODUCT((<month_col>:<month_col><>"")*RATE_col_range), 0)` per month |

- Row position: `grid_person_last_row + 1` (no blank spacer — same outline group).
- Styling: BLACK_FORMULA font on both label and values; `€#,##0` format on value cells; same horizontal alignment as month-header row (centered).
- Vacation column fills: each month cell corresponding to a vacation month in the project span gets VACATION_FILL (yellow) applied as a direct fill. Value formula stays identical — the SUMPRODUCT evaluates to 0 (all vacation cells above are empty).
- Outline group: `ws.row_dimensions[grid_total_row].outline_level = 1` — folds with the rest of the grid.

### 3.2 PER-WP SUMMARY — Yearly total € / Yearly budget € rows

Two new rows appended below `wp_summary_last_row`, no spacer (same outline group):

| Row | A | B | C | D-H | I… (per year) |
|---|---|---|---|---|---|
| yearly_total_row | `Yearly total €` | — | `=SUM(C{wp_first}:C{wp_last})` | — | `=SUM(<year_col>{wp_first}:<year_col>{wp_last})` |
| yearly_budget_row | `Yearly budget €` | — | `=SUM(B{wp_first}:B{wp_last})` | — | see 5.3 |

- **Col-C cross-checks**: yearly total € col C = total Assigned € across WPs; yearly budget € col C = total Ceiling € across WPs. Divergence between row-C-sum and (sum of year cols on that row) flags formula drift.
- **Row styling**: BLACK_FORMULA font; `€#,##0` on col C and year cells; col A label bolded (HEADER_FONT); cells B, D-H intentionally blank (no fill, no value).
- **Outline group**: both rows at `outline_level = 1`.

### 3.3 PER-PERSON SUMMARY — Largest gap dropped, Rate shifts

Before (v2):

| A | B | C | D | E | F | G | H | I |
|---|---|---|---|---|---|---|---|---|
| Person | Assigned | Pool | Over-pool? | Switches | First log | Last log | Largest gap | Rate |

After (this doc):

| A | B | C | D | E | F | G | H |
|---|---|---|---|---|---|---|---|
| Person | Assigned | Pool | Over-pool? | Switches | First log | Last log | Rate |

- Total cols: 9 → 8.
- `_PERSON_SUMMARY_METRIC_HEADERS` shrinks: drop `"Largest gap"`; `"Rate"` remains the last entry.
- `_PERSON_SUMMARY_RATE_COL` recomputes to col H.
- Every formula that referenced the v2 Rate col (PER-WP SUMMARY Assigned €, Y{yyyy} € cells, Burn-rate red-CF `MAX(…)`) retargets from col I to col H. Row ranges unchanged.
- Placeholder cell + comment in the old col H removed entirely — no residual styled "" cell.

### 3.4 Vacation month model

**Config scalar** (new):
- Named range: `vacation_month` → `Config!B{row}`.
- Value: calendar-month abbreviation string ("Jan"…"Dec"). Default: `"Jul"`.
- Font: `BLUE_INPUT` (user-editable). Validation: cell comment listing valid values; no formal data validation rule (keeps LO-compat).
- Placement: adjacent to existing scalars `project_start`, `fte_rate` in the Config PROJECT PARAMETERS block. Exact row determined by `config_sheet.py` — new row appended under existing params. Reader reads by named range, not row index — row shift is invisible to downstream.

**Config derived literal** (new):
- Named range: `num_vacation_months` → `Config!B{row2}`.
- Value: build-time integer = count of calendar-vacation-months (e.g. Julys) falling inside `[project_start, project_start + months − 1]`.
- Font: `BLACK_FORMULA` (Rule 3 carve-out — builder-written literal; documented in the builder comment + in-cell comment per the `version_stamp` precedent).
- Comment: "Recomputed on rebuild from vacation_month + project_start + Project Months."

**Solver extension**:
- `Person` dataclass gains `vacation_months: frozenset[int] = frozenset()` (project-1-indexed month numbers). Default empty for backward-compat with any direct solver callers.
- Greedy pass skips candidates where `month in person.vacation_months`.
- Switch-reduction pass: swap eligibility requires neither month being in either person's vacation set. (For uniform-July the check is trivially symmetric, but the per-person field keeps the contract clean for the future Option-A extension.)

**Builder extension**:
- On build, compute `vacation_months_project = frozenset(project_month_index_for(vac_month, y) for y in project_years if project_month_index_for(vac_month, y) ∈ [1, project_months])`. Pass to every Person.
- TIMELOG GRID cell writer: if `month in vacation_months_project`, write `""` + `VACATION_FILL` + `Protection(locked=True)` + no BLUE_INPUT font (cell is not user-editable — keep default font).
- Monthly-total row: same VACATION_FILL on the total-row cells corresponding to vacation months (formula still evaluates to 0).

**Pool formula** (PER-PERSON SUMMARY col C):
- Before: `=ROUND(Config!B{cfg_row}*project_months, 0)`.
- After: `=ROUND(Config!B{cfg_row}*(project_months-num_vacation_months), 0)` — `num_vacation_months` resolved via its named range.

**Styles addition**:
- `styles.py` gains `VACATION_FILL = PatternFill(fill_type="solid", fgColor="FFF2CC")` (pale yellow; distinct from existing fills). Imported into `timelog_plan.py`.

**Reader extension**:
- `read_config` parses `vacation_month` named range → str. Converts to calendar month int (`"Jul"` → 7) via `datetime.datetime.strptime(s, "%b").month`. Passes as `project_params["vacation_month_num"]` (int) to the builder.
- If the named range is missing (pre-this-doc workbook), default to 7 — one-shot migration.

### 3.5 Fill-tier hierarchy (for reference)

Cells in the TIMELOG GRID after this doc:

| Fill | Meaning | Precedence |
|---|---|---|
| `VACATION_FILL` (yellow) | Vacation month — uneditable | Direct fill; no CF can override (no CF rule targets vacation cells) |
| WP pastel (CF) | Cell contains a WP code | CF rule matches; overrides any empty-cell fill |
| `CF_RED_FILL` (CF) | WP code outside its active window | CF rule matches at higher priority than pastel |
| `GREY_FILL` (direct) | In-window-empty | Direct; no CF matches an empty cell |
| `DARK_GREY_FILL` (direct) | Out-of-window-empty (no person-WP window covers this month) | Direct; no CF matches |
| no fill | — | Default |

Vacation cells are always empty-string; no CF rule fires on them. Direct yellow shows.

---

## 4. Forward-compat with per-(person, year) vacation (Option A)

The scalar model in §3.4 is designed so the later Option A (Vacation Schedule table on Config) slots in additively:

- Keep the `vacation_month` scalar as the **default**.
- Add a `Vacation Schedule` table on Config: rows = persons (from Team Roster), cols = calendar years spanned, cells = blank or month abbreviation.
- Reader resolves per-(person, year) with **table-cell overrides scalar**. Blank → scalar. Table absent → scalar applies uniformly (current behaviour).
- `Person.vacation_months` stays a `frozenset[int]` — builder computes per-person rather than uniform.
- `num_vacation_months` on Config would then become per-person or be retired in favour of a Pool formula that sums per-person: `=ROUND(Config!B{cfg_row}*(project_months-COUNTIF(<vacation row>,"<>")), 0)`.
- No cell moves on Timelog Plan. No solver-contract change. Pool formula may rewire; everything else holds.

This doc commits to the scalar path only. Option A is mentioned solely to rule out foreclosing moves.

---

## 5. Formula specification (consolidated)

Reference key (unchanged from v2):
- `R1, Rn` = first / last person row of merged grid.
- `S1, Sn` = first / last person row of PER-PERSON SUMMARY.
- `W1, Wn` = first / last WP row of PER-WP SUMMARY.
- `M1col, MNcol` = first / last month col letters on the grid.
- `RATE_col` = col H on PER-PERSON SUMMARY (shifted from col I).
- `hdr` = month-header row.
- `midx` = hidden month-index helper row.

### TIMELOG GRID Monthly total € row (row `grid_total_row`)

| Col | Formula | Notes |
|---|---|---|
| A | literal `"Monthly total €"` | HEADER_FONT |
| B…MN | `=ROUND(SUMPRODUCT((<col>{R1}:<col>{Rn}<>"")*H{S1}:H{Sn}),0)` per month col `<col>` | BLACK_FORMULA, `€#,##0`; vacation col cells also get VACATION_FILL |

### PER-WP SUMMARY Yearly total € row (row `yearly_total_row`)

| Col | Formula | Notes |
|---|---|---|
| A | literal `"Yearly total €"` | HEADER_FONT |
| B | blank | no value |
| C | `=SUM(C{W1}:C{Wn})` | BLACK_FORMULA, `€#,##0`; cross-check |
| D–H | blank | — |
| I… (per year) | `=SUM(<year_col>{W1}:<year_col>{Wn})` | BLACK_FORMULA, `€#,##0` |

### PER-WP SUMMARY Yearly budget € row (row `yearly_budget_row`)

| Col | Formula | Notes |
|---|---|---|
| A | literal `"Yearly budget €"` | HEADER_FONT |
| B | blank | — |
| C | `=SUM(B{W1}:B{Wn})` | BLACK_FORMULA, `€#,##0`; grand project ceiling |
| D–H | blank | — |
| I… (per year) | see below | BLACK_FORMULA, `€#,##0` |

**Yearly budget € per-year cell formula** (year col `y`, WPs indexed `k = 1..num_wps`):

```
=<prop_1y>*Config!E{wp1_row} + <prop_2y>*Config!E{wp2_row} + … + <prop_Ny>*Config!E{wpN_row}
```

Where `<prop_ky>` is a build-time float literal:
```
prop_ky = months_of_wpk_in_year_y / total_sched_months_for_wpk    (if WPk overlaps year y)
prop_ky = 0                                                         (otherwise)
```

- `months_of_wpk_in_year_y` = count of 1-indexed project months `m` such that: `m ∈ WPk active segment(s)` AND `calendar_year(project_start + m − 1) == y`.
- `total_sched_months_for_wpk` = `_wp_sched_months(wpk.code, wp_schedule)` — respects multi-segment WPs today; if WP gap logic is removed later (Doc B), this collapses to `end - start + 1` with no formula change.
- WPs with `prop_ky = 0` across every year are still included in the term sum with a literal `0*Config!E{row}` — keeps the formula structure uniform and readable. (Alternative: omit zero terms; minor diff.) **Choice: omit zero terms.** A WP that doesn't overlap a year contributes nothing; the formula lists only the overlapping WPs.
- Formula stays **live to Config FTE / rate changes** on each WP's `Config!E{row}` (which is `=FTE × Rate`). Rebuild required if WP schedule changes (standard convention).

**Worked example**: Project starts M1 = Jan 2026, 36 months. WP1 active M1-M18 (18 months total). Year 2026 overlaps M1-M12 (12 months). Year 2027 overlaps M13-M18 (6 months). Assume Config!E{wp1_row} = €100,000.

- Yearly budget row 2026 cell includes: `12/18 * Config!E{wp1_row}` term → €66,666 when recalced.
- Yearly budget row 2027 cell includes: `6/18 * Config!E{wp1_row}` term → €33,333.
- Sum across 2026+2027 terms for WP1 = €100,000 (the full ceiling). Other WPs contribute their own terms.
- Col C = `=SUM(B{W1}:B{Wn})` = sum of all WP ceilings. Total per-year cells (summed across years) = same value. Cross-check passes.

### PER-PERSON SUMMARY Pool formula (col C)

```
=ROUND(Config!B{cfg_row}*(project_months-num_vacation_months),0)
```

- `project_months` and `num_vacation_months` both resolved via named ranges.
- Green xref font (unchanged — still references Config).

### Solver vacation exclusion (pseudo-diff)

```python
# timelog_solver.py — greedy pass
for month in range(1, project_months + 1):
    if month in person.vacation_months:        # NEW
        continue
    # existing placement logic
```

---

## 6. Phased implementation plan

Each phase sized for one focused session; ends with `pytest tests/ -q` green + recalc pass. Ordered to minimise re-work — `_PERSON_SUMMARY_RATE_COL` shifts in Phase 1, after which every downstream reference lands in its final position on first write.

### Phase 1 — Drop `Largest gap`, shift Rate

**Scope:** smallest change, unblocks Phase 2+ references.

- `timelog_plan.py`:
  - `_PERSON_SUMMARY_METRIC_HEADERS`: remove `"Largest gap"`.
  - `_PERSON_SUMMARY_RATE_COL` recomputes automatically (now col H).
  - `_write_person_summary_metrics`: delete the col-H placeholder write path; ensure Rate is now written at `c0 + 6` (which resolves to col H). Current code writes Rate at `c0 + 7`; decrement.
  - Every internal reference to `_PERSON_SUMMARY_RATE_COL` stays symbolic — no hardcoded col letter to grep.
- `reader.py`: no change (reader reads by named range / header text, not col index).
- Tests (`test_timelog_plan.py`):
  - Existing `test_layout_col_shift` docstring + assertion on col 9 (v2 Rate col I) → retarget to col 8 (H). Header-list assertion at `_PERSON_SUMMARY_METRIC_HEADERS` line drops `"Largest gap"`.
  - New `test_person_summary_shape`: 8-col shape; no `"Largest gap"` header.
  - New `test_rate_col_h`: Rate formula lands at col H; Assigned € SUMPRODUCT on PER-WP SUMMARY resolves to col H range.
  - **Sibling-file sweep** (per `.claude/rules/python-refactors.md` "Removing exported constants"): `grep -rn "Largest gap\|col.*[Rr]ate\|column=9" tests/` before committing Phase 1 — `test_timelog_plan_roundtrip.py` currently clean per 2026-04-19 audit, but re-check to catch any new additions.

**Verify:** `pytest tests/ -q` green; recalc clean.

### Phase 2 — Vacation month model

**Scope:** Config scalar + derived literal + solver extension + grid fill tier + Pool formula.

- `styles.py`: add `VACATION_FILL`.
- `layout.py`: no change — vacation doesn't alter column offsets.
- `config_sheet.py`:
  - Append `Vacation month` param row with default `"Jul"`, BLUE_INPUT.
  - Append `Num vacation months` param row as builder-written literal (BLACK_FORMULA), comment explaining derivation.
  - Register named ranges `vacation_month` and `num_vacation_months`.
- `reader.py`: parse both named ranges; fallback `vacation_month → 7` when missing.
- `workbook.py` (orchestrator): compute `vacation_months_project` from `vacation_month_num` + `project_start` + `project_months`; write back `num_vacation_months` literal on Config; pass `vacation_months_project` to every Person.
- `timelog_solver.py`:
  - `Person` gains `vacation_months: frozenset[int] = frozenset()`.
  - Greedy pass: skip `(person, month)` when `month in person.vacation_months`.
  - Switch-reduction pass: swap-eligibility guard.
- `timelog_plan.py`:
  - `_write_grid_cell`: if `month in vacation_months_project`, write `""` + `VACATION_FILL` + `Protection(locked=True)`; skip BLUE_INPUT font.
  - `_write_person_summary_metrics`: Pool formula updated to use `num_vacation_months` named range.
- Tests:
  - `test_vacation_config_scalar`: `vacation_month` named range resolves, default `"Jul"`.
  - `test_num_vacation_months_literal`: literal matches count of July months in project span.
  - `test_vacation_cells_locked`: July cells across every person row are `""`, have VACATION_FILL, locked=True.
  - `test_solver_skips_vacation`: solver output never places a WP code in a vacation month.
  - `test_pool_formula_accounts_for_vacation`: recalced Pool value = `max_fte × (project_months − num_vacation_months)` rounded.

**Verify:** `pytest tests/ -q` green; recalc clean; open in Excel/LO, eyeball July columns yellow across the grid, verify a July cell rejects manual WP-code entry (locked).

### Phase 3 — Monthly total € row

**Scope:** single new row at bottom of TIMELOG GRID.

- `timelog_plan.py`:
  - Introduce `grid_total_row = grid_person_last_row + 1`.
  - Extend grid outline group to include the total row.
  - Write label at col A; write `=ROUND(SUMPRODUCT(…),0)` per month col, referencing the shifted Rate col (H) from Phase 1.
  - Apply `VACATION_FILL` to total-row cells in vacation months (matches Phase 2's visual continuity).
- Tests:
  - `test_monthly_total_row_present`: row exists with label + formulas.
  - `test_monthly_total_vacation_zero`: recalced value in a vacation month = 0.
  - `test_monthly_total_non_vacation_nonzero`: for a month with ≥1 assigned person, value = sum of those persons' rates.
  - `test_monthly_total_row_vacation_fill`: total-row cells in vacation cols have VACATION_FILL.

**Verify:** `pytest tests/ -q`; recalc; eyeball sums.

### Phase 4 — Yearly total € + Yearly budget € rows

**Scope:** two new rows at bottom of PER-WP SUMMARY.

- `timelog_plan.py`:
  - Introduce `yearly_total_row = wp_summary_last_row + 1` and `yearly_budget_row = wp_summary_last_row + 2`.
  - Extend PER-WP SUMMARY outline group to include both rows.
  - Write Yearly total € row (col A label, col C `=SUM(...)`, per-year `=SUM(...)`).
  - Write Yearly budget € row (col A label, col C `=SUM(...)`, per-year term-sum formula built from `wp_schedule` × Config WP refs).
- Tests:
  - `test_yearly_rows_present`: both rows exist; outline group correct.
  - `test_yearly_total_col_c_crosscheck`: col-C value = sum of Assigned € across WPs = sum of per-year cells on that row.
  - `test_yearly_budget_col_c_crosscheck`: col-C value = sum of Ceiling € across WPs = sum of per-year cells on that row.
  - `test_yearly_budget_single_wp_single_year`: for a WP fully contained in one calendar year, that year's term = 100% of the WP's Config!E{row}.
  - `test_yearly_budget_multi_segment_wp`: for a multi-segment WP, proportion uses `_wp_sched_months` (active segments summed, gap months excluded).

**Verify:** `pytest tests/ -q`; recalc; eyeball yearly budget adds up to grand total.

### Phase 5 — Docs + final sweep

**Scope:** close out the plan; update living docs per CLAUDE.md critical rule #6.

- `docs/design-invariants.md`: add entries for (a) Monthly total € / Yearly total € / Yearly budget € cross-check rule (col-C sum equals row-sum across year cols), (b) vacation-month Pool formula invariant, (c) vacation cells are always locked and empty.
- `docs/internals.md` § Sheet Builders → Timelog Plan: add monthly-total row, yearly-total/budget rows, vacation fill tier, Pool formula shift.
- `docs/architecture.md`: if Mermaid shows Config params, add `vacation_month` + `num_vacation_months` nodes.
- `CLAUDE.md` trigger table: add "vacation month, July exclusion, Pool reduction" row pointing at this doc (once archived, update path).
- `.claude/rules/openpyxl-patterns.md`: no change — no new formula pattern introduced.
- Archive this doc to `docs/plans/archived/2026-04-19-timelog-plan-enhancements-design.md` after Phase 5 verifies.

**Verify:** `pytest tests/ -q` full suite; recalc; eyeball workbook.

---

## 7. Tests to add / update

| Test | Phase | Asserts |
|---|---|---|
| `test_person_summary_shape` | 1 | 8 cols (was 9); no `"Largest gap"` header |
| `test_rate_col_h` | 1 | Rate formula at PER-PERSON SUMMARY col H; PER-WP SUMMARY Assigned € SUMPRODUCT references col-H range |
| `test_vacation_config_scalar` | 2 | `vacation_month` named range resolves; default `"Jul"` |
| `test_num_vacation_months_literal` | 2 | Literal matches count of July months in `[project_start, project_start + months)` |
| `test_vacation_cells_locked` | 2 | Every July cell across person rows is `""`, has VACATION_FILL, Protection locked=True |
| `test_solver_skips_vacation` | 2 | `solve(...)` output never contains a (person, month) with month in vacation_months |
| `test_pool_formula_accounts_for_vacation` | 2 | Recalced Pool = `round(max_fte × (project_months − num_vacation_months))` |
| `test_monthly_total_row_present` | 3 | Row at `grid_person_last_row + 1` with label + per-month formulas |
| `test_monthly_total_vacation_zero` | 3 | Recalced: July cell = 0 |
| `test_monthly_total_non_vacation_nonzero` | 3 | For a month with k assigned persons, cell = sum of those persons' recalced rates |
| `test_monthly_total_row_vacation_fill` | 3 | July cells on the total row have VACATION_FILL |
| `test_yearly_rows_present` | 4 | Both rows below last WP row; in outline group |
| `test_yearly_total_col_c_crosscheck` | 4 | Col C = SUM(Assigned € across WPs) = SUM(year-cells of this row) |
| `test_yearly_budget_col_c_crosscheck` | 4 | Col C = SUM(Ceiling € across WPs) = SUM(year-cells of this row) |
| `test_yearly_budget_single_wp_single_year` | 4 | WP fully in year Y: year-Y cell term = 100% × Config!E{wp_row} |
| `test_yearly_budget_multi_segment_wp` | 4 | Multi-segment WP: proportion respects `_wp_sched_months` |
| `test_vacation_preserved_under_resolve_flag_false` | 2 | Edit grid cell via openpyxl (non-vacation), set flag FALSE, rebuild → vacation cells remain locked-yellow; preserved edits intact |

All tests must exercise **recalculated** workbooks per CLAUDE.md Rule 2.

---

## 8. Docs to update (per CLAUDE.md rule #6)

| File | Change |
|---|---|
| `docs/design-invariants.md` | Add: (a) col-C cross-check rule for Yearly total € / Yearly budget €. (b) Pool formula invariant: `Pool = max_fte × (project_months − num_vacation_months)`. (c) Vacation cells are locked and empty across every person row. |
| `docs/internals.md` | § Sheet Builders → Timelog Plan: describe Monthly total € row, Yearly total/budget rows, vacation fill tier, Pool formula shift. |
| `docs/architecture.md` | Update Config params block if Mermaid enumerates them. |
| `CLAUDE.md` trigger table | New row for "vacation month, Pool reduction" → this doc. |
| `docs/grant-domain.md` | Short entry: "Vacation months — uniform default (July); never logged to project." |
| `docs/common-tasks.md` | "Add vacation exception for a specific person" — placeholder noting current model is uniform; forward reference to future Option A if/when it ships. |

Grep before archive (per "Superseded docs: consume before archive" global rule): pull load-bearing content from this doc into `design-invariants.md` + `internals.md` before the `git mv`.

---

## 9. Decisions log (2026-04-19 brainstorm)

| # | Decision | Rationale |
|---|---|---|
| D1 | Monthly total € row below grid, inside grid's outline group | Standard "totals under data" idiom; scrolls with grid under frozen panes |
| D2 | Monthly total formula = `SUMPRODUCT((col<>"") × rate_vector)` | No WP-code disambiguation needed — every non-empty cell is 1.0 FTE at the person's rate. Simpler than per-WP enumeration |
| D3 | Yearly total row & Yearly budget row layout: **B** (col A label + col C cross-check sum + per-year cells) | Col-C cross-check doubles as a formula-drift sentinel |
| D4 | Yearly budget per-year value = build-time term-sum formula with literal proportions × live Config!E{wp_row} refs | Matches existing "WP sched months" pattern (rebuild-dependent on schedule, live on Config FTE/rate) |
| D5 | Yearly budget proportion = `active_months_in_year / total_sched_months`, both segment-aware | Consistent with `_wp_sched_months` convention; survives later WP gap removal without formula rewrite |
| D6 | Drop `Largest gap` entirely (do not rename to `Longest idle run`) | Placeholder since v1; no user ask; monthly total row surfaces thin months visually; name collision with incoming WP gap removal |
| D7 | Uniform vacation month via Config scalar (default `"Jul"`); no per-(person, year) overrides in this doc | Matches current workflow; per-person deferred to future Option A; forward-compat preserved |
| D8 | Pool formula option **A**: `max_fte × (project_months − num_vacation_months)` | Keeps `max_fte` semantically clean as "fractional allocation during available months"; vacation is a separate axis |
| D9 | `num_vacation_months` as Config-level build-time literal with named range | Same pattern as `version_stamp` on Config; Rule 3 carve-out; referenced by Pool formula |
| D10 | Vacation cells: yellow fill + `Protection(locked=True)` + empty string; **no "VAC" text** | Keeps `grid<>""` SUMPRODUCT semantics unchanged for Assigned €, OOW, Switches, etc. |
| D11 | 4th fill tier (`VACATION_FILL`, pale yellow `#FFF2CC`) distinct from `GREY_FILL`, `DARK_GREY_FILL`, `LIGHT_BLUE_FILL`, `LIGHT_GREEN_FILL` | Visual disambiguation from in-window-empty / out-of-window-empty |
| D12 | Vacation exclusion applied in solver greedy AND switch-reduction passes | Prevents Pass 2 from swapping a legal placement into a vacation slot |
| D13 | Apply VACATION_FILL to monthly-total-row cells as well as grid cells | Visual column continuity — yellow runs top-to-bottom for vacation months |
| D14 | Phase 1 drops `Largest gap` first so `_PERSON_SUMMARY_RATE_COL` shifts once, before Phase 3 writes monthly-total formulas that reference it | Re-work audit: avoids writing the monthly-total SUMPRODUCT twice (once against col I, once against col H) |
| D15 | Doc A (this doc) merged with vacation-month design when user's simplification collapsed Option A → uniform scalar; Doc B (WP gap removal) remains separate | Scope check: merged Docs A+C share grid + PER-PERSON SUMMARY surfaces; Doc B is cross-cutting and independent |
| D16 | Yearly budget formula omits WP terms with `prop_ky = 0` (WP doesn't overlap year) rather than writing `0*Config!E{row}` terms | Compact; the self-check (col-C sum = year-sum) still holds |
