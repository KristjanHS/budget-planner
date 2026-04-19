# Timelog Plan v2 — Column Surface Redesign & LO Recalc Bug Fixes

**Status:** Validated via brainstorming 2026-04-19. Ready for phased implementation.
**Scope:** Restructure the Timelog Plan tab into a single editable grid + richer PER-WP summary, add EUR-denominated per-WP views (Buffer €, Burn rate, yearly breakdown, WP sched months), fix two silent LibreOffice recalc bugs (First log, Last log), and reorder tabs so Timelog Plan + Actuals-Hours sit immediately after Config.

**Supersedes:** `docs/plans/archived/2026-04-18-timelog-plan-design.md` (v1 — shipped). This is a column-surface evolution of v1, not a replacement of the solver or the builder architecture.

---

## 1. Motivation

The v1 Timelog Plan tab landed in April 2026 with a solver + locked PROPOSED band + editable ASSIGNED band (each ASSIGNED cell defaults to `=<Proposed>`). In daily use the user surfaced five frictions:

1. **Two grids is one too many.** PROPOSED and ASSIGNED carry the same shape and almost-identical content. The `=<Proposed>` default was chosen to preserve user edits across rebuild, but in practice the user re-solves often — they would rather have one editable grid and a workbook-level toggle controlling whether the next rebuild re-solves.
2. **"Window months remaining" is wallclock — not a planning metric.** A WP can have 8 months left but already be full, or be near-closing with big shortfall. The number doesn't tell you what to do next.
3. **Over/Undershoot in two columns.** A signed column is easier to scan when "positive = buffer = good" and "negative = shortfall = bad" are well-established domain semantics.
4. **Underload (months) is rate-blind.** The user optimizes for EUR; a count of remaining free month-slots doesn't reflect that.
5. **Two stale-looking cells.** First log always showed `""` even when a person had assignments; Last log always showed the project's last month for everyone. Both are LibreOffice recalc pitfalls around `MATCH`/`LOOKUP` with array coercion.

Also two additive asks:
- **Yearly EUR per WP** — budgeting is on calendar-year boundaries; a per-WP × per-year EUR table (extension of PER-WP SUMMARY) gives the user the view they reconcile against.
- **WP sched months** — pairs with "Assigned months" to surface slot coverage (`Assigned / Sched`), the rate-blind planning view that complements the EUR views.

And a tab-order change: Timelog Plan is the user's first-touch sheet in any planning session, so it and its adjacent source (Actuals-Hours) should sit immediately after Config.

---

## 2. Goals / Non-goals

### Goals
- One editable grid (no side-by-side PROPOSED band).
- Workbook-level "Re-run solver on next rebuild" toggle — user-controlled, no auto-reset.
- PER-WP SUMMARY becomes a dashboard: Buffer €, Burn rate needed, WP sched months, per-year EUR columns.
- Merged grid starts at col B (no per-row metric col).
- Rate helper moves to the end of PER-PERSON SUMMARY (single source).
- First log + Last log formulas rewritten to the SUMPRODUCT-forcing-array pattern (LO-stable).
- Tab order: `Config → Timelog Plan → Actuals-Hours → (rest unchanged)`.

### Non-goals (YAGNI)
- **Solver algorithm changes.** The two-pass greedy + switch-reduction stays exactly as `timelog_solver.py` ships today.
- **Config table layout changes.** No new Config columns or tables. The toggle lives on Timelog Plan itself.
- **Round-trip preservation of individual cells.** Under the new toggle semantics, `TRUE` = rebuild re-solves and overwrites the grid unconditionally; `FALSE` = rebuild preserves the grid verbatim. No per-cell sentinel or "solver-controlled?" flag.
- **"Last solver run" audit cell.** Considered during brainstorm; skipped for simplicity. Can be added later if the toggle's effect proves opaque in use.
- **Cumulative-EUR-per-person row metric.** Considered as a replacement for Underload; dropped in favour of leaving the merged grid metric-free and keeping Assigned count in PER-PERSON SUMMARY only.
- **Distinct-WPs-per-person metric.** Not asked for.

---

## 3. Final layout specs

### 3.1 Merged editable grid (was PROPOSED + ASSIGNED)

| col A | col B … col (B + months − 1) |
|---|---|
| Person (green xref to Config Team Roster) | M1, M2, …, MN — solver-written initial WP codes, blue-input editable |

- No per-row metric column. `TIMELOG_MONTH_COL_START`: 4 → 2. `TIMELOG_RATE_HELPER_COL`: removed.
- Freeze panes: was `D3`, becomes `B3` (rows 1-2 + col A frozen).
- Cells hold literal values, not formulas. Empty cells outside a person's WP windows carry `DARK_GREY_FILL`; in-window-but-empty cells carry `GREY_FILL`. Both visuals stay on hand-typed literals (solver writes the fills; user edits don't disturb them unless they blank / re-enter values).
- Pastel per-WP CF and red OOW CF rules carry over to the merged grid; the rules that referenced both PROPOSED and ASSIGNED halves now reference only one grid.

### 3.2 Row 2 — traffic lights + toggle

Five `(label, value)` pairs across cols 1-10:

| cols | label | value semantics |
|---|---|---|
| 1, 2 | Any undershoot? | `COUNTIF(Buffer_col, "<0")` over PER-WP SUMMARY (replaces the current undershoot-col >0 count) |
| 3, 4 | Any out-of-window cell? | `SUM(OOW_col)` over PER-WP SUMMARY |
| 5, 6 | Any over-pool person? | `COUNTIF(Over-pool_col, TRUE)` over PER-PERSON SUMMARY |
| 7, 8 | Total switches | `SUM(Switches_col)` over PER-PERSON SUMMARY |
| **9, 10** | **Re-run solver on next rebuild?** | TRUE/FALSE blue input, default TRUE on init, **no auto-reset**, named range `timelog_resolve_flag` |

CF on first three value cells: red fill when > 0. Toggle cell is TRUE/FALSE only; styled BLUE_INPUT; not flagged by CF.

### 3.3 PER-WP SUMMARY (rows = WPs)

| A | B | C | D | E | F | G | H | I … (num_years) |
|---|---|---|---|---|---|---|---|---|
| WP code | Ceiling € | Assigned € | **Buffer €** | Assigned months | **WP sched months** | **Burn rate needed** | Any OOW? | **Y2026 €, Y2027 €, Y2028 € …** |

Total cols: `8 + num_years`. No col-C spacer (removed — col-C-collapse logic no longer needed).

### 3.4 PER-PERSON SUMMARY (rows = persons)

| A | B | C | D | E | F | G | H | **I** |
|---|---|---|---|---|---|---|---|---|
| Person | Assigned | Pool | Over-pool? | Switches | First log | Last log | Largest gap | **Rate** |

Total: 9 cols. Underload removed from col B; Assigned (count) moves to col B; remaining metrics shift left by one. Rate appended as the last col — no longer duplicated on each person row of the merged grid.

### 3.5 Tab position

Current creation order in `budget_planner/workbook.py`: `Config → Capacity → Plan → Trends → Timelog Plan → Actuals-Hours → Plan-Forecast` (Actuals EUR + Summary created elsewhere in the orchestrator; verify during Phase 4).

Target order: `Config → Timelog Plan → Actuals-Hours → (everything else in current relative order)`.

---

## 4. Solver semantics under toggle

**On init (workbook being built for the first time):**
1. Solver runs. Merged grid filled with WP codes + empty cells.
2. Toggle cell written as `TRUE` (default).

**On rebuild (workbook being regenerated from an existing file):**
1. Reader reads the toggle cell's current value (via `timelog_resolve_flag`). If the cell is missing (pre-v2 workbook), default TRUE.
2. Reader reads the current merged-grid state into a `dict[(person_name, month_num), wp_code | ""]`.
3. Builder receives both values:
   - If toggle = **TRUE**: solver runs, merged grid overwritten with the solver's output; existing grid state discarded. **Toggle left at TRUE** (no auto-reset — explicit per user spec).
   - If toggle = **FALSE**: solver does NOT run; grid cells written back from the dict verbatim; cells for persons / months no longer in scope are dropped; new persons / months get empty cells (not solver-filled) — user must flip the toggle and rebuild to populate them.

### Semantics the user is signing off on
- "Re-solve often, preserve occasionally" — match the user's workflow.
- The user is aware that `TRUE` wipes manual edits; they flip to `FALSE` when they want to keep edits through a rebuild (typically the final-draft stage of a review).

### Reader backward-compatibility
- Pre-v2 workbooks have the old PROPOSED + ASSIGNED shape and no toggle cell. Reader detects missing toggle → defaults to TRUE → solver runs → grid rebuilt in v2 shape. One-shot migration; no separate migration script needed.
- The existing `read_timelog_plan` function keeps its signature but returns `{"grid": {...}, "resolve_flag": bool}` instead of `{"assigned_overrides": {...}}`. Consumers (workbook orchestrator) updated accordingly.

---

## 5. Per-item design details

### 5.1 Merge ASSIGNED ← PROPOSED (Item 1)

- The `=<Proposed>` formula-vs-literal distinction (which was the v1 round-trip mechanism) disappears. There's only one grid.
- Solver output is still produced by `solve(persons, wps, windows, project_months)` — same call, same types, same determinism. The builder just writes the result to the merged grid directly instead of to PROPOSED and synthesising `=<Proposed>` formulas in a sibling ASSIGNED grid.
- `Protection(locked=False)` on every cell (was locked on PROPOSED).
- `BLUE_INPUT` font on every cell (unchanged from v1 ASSIGNED).
- Fills: in-window-empty → `GREY_FILL`; out-of-window-empty → `DARK_GREY_FILL`; WP-assigned → no fill (pastel CF applies). Same rules as v1.

### 5.2 Hide/move Rate helper (Item 2)

- `_write_rate_helper_cell` removed from merged-grid person rows.
- New: Rate cell on each PER-PERSON SUMMARY row, written by a new helper `_write_summary_rate_cell(ws, row)` — same VLOOKUP formula (`=VLOOKUP(A{row}, roster_lookup, 4, FALSE)`), rendered on the PER-PERSON SUMMARY row's last col (I).
- Per-WP Assigned € SUMPRODUCT formulas target the PER-PERSON SUMMARY's rate col. Row-alignment works because persons appear in the same order in both bands (already guaranteed via `person_rows_by_band`). SUMPRODUCT takes vectors-of-same-length from anywhere; it doesn't require shared row indices.

### 5.3 Burn rate needed (Item 3)

Formula for each WP row:
```
=IF(AND(<buffer_cell> < 0, <wallclock_remaining> > 0),
    -<buffer_cell> / <wallclock_remaining>,
    "")
```
where `<wallclock_remaining>` = `MAX(0, wp.active_end - current_month)` — the literal expression that was the old `Window months remaining` column.

- Blank (empty string) when buffer ≥ 0 (no shortfall → nothing to burn).
- Blank when wallclock = 0 (WP window closed → no time to close the gap by spending).
- Numeric (EUR/month) when buffer < 0 AND wallclock > 0.
- Red CF: `{burn_rate_cell} > MAX(all eligible persons' rates)` — the gap is mathematically unreachable with the current team. Formula for the max rate uses MAX over the PER-PERSON SUMMARY rate col (I).

`current_month` is an existing named range (Config scalar). `wp.active_end` is a build-time literal (matches the existing `Any out-of-window?` pattern).

### 5.4 Remove Underload + consolidate Assigned (Item 4)

- Drop the col-B Underload formula from merged-grid person rows (effectively: merged grid has no col B metric).
- Drop the col-B Underload formula from PER-PERSON SUMMARY rows.
- Drop the existing "Assigned" col C of PER-PERSON SUMMARY; **re-add it at col B** (same formula: `=COUNTIF(<row_range>, "WP*")`).
- Shift Pool, Over-pool?, Switches, First log, Last log, Largest gap left by one col (now C-H).
- Rate newly at col I.
- Remove Rule 6 (the Underload font-colour CF).

### 5.5 Merge Overshoot / Undershoot → Buffer € (Item 5)

- Drop cols D (Overshoot €) and E (Undershoot €) from PER-WP SUMMARY.
- Add col D: `Buffer €` with formula `=C{row} - B{row}` (Assigned − Ceiling). Positive = buffer/overshoot (good); negative = shortfall (bad).
- CF: `Buffer € > 0` → green fill (`LIGHT_GREEN_FILL` or similar from `styles.py`). `Buffer € < 0` → red fill (`CF_RED_FILL`). At zero: default fill.
- Traffic light #1 (`Any undershoot?`) re-wires from "count Undershoot > 0" to `COUNTIF(D{first}:D{last}, "<0")` — same semantic, new column.

### 5.6 Per-WP × per-year EUR breakdown (Item 6)

- Calendar-year columns appended to PER-WP SUMMARY right of "Any OOW?".
- Number of year cols = number of distinct calendar years spanned by `project_start` through `project_start + months`. Computed via the existing `_calendar_year_groups(project_start, project_months)` helper in `timelog_plan.py`.
- Each year col: header `Y2026 €`, `Y2027 €`, etc. written on the WP-summary header row.
- Each cell formula (WP row × year col):
  ```
  =IF(<overlap_check>,
       SUMPRODUCT((grid[year_start_col:year_end_col] = "WPk") * rate_col),
       "")
  ```
  where `<overlap_check>` = the WP's active window overlaps the year's month range, hardcoded as a build-time boolean (no formula — either the column shows for that WP or it's blank).
- Sum across year cols equals `Assigned €` (col C). Built-in cross-check; no separate "Total" column.
- `rate_col` targets the PER-PERSON SUMMARY rate col (I), same vector as the Assigned € formula uses.

### 5.7 WP sched months column (added 2026-04-19)

- New col F in PER-WP SUMMARY.
- Value: build-time integer literal = `sum(s["end"] - s["start"] + 1 for s in wp_schedule if s["id"] == wp.code)`. Respects multi-segment WPs (e.g., WP3's gap counts only the active segments).
- Written as `BLACK_FORMULA` font even though it's a literal — same pattern as `version_stamp` on Config (Rule 3 carve-out documented in `.claude/rules/dev-workflow.md`). Cell comment: "Number of months the WP SCHEDULE marks as active. Recomputed on rebuild."
- No CF — purely informational alongside "Assigned months".

### 5.8 Tab reorder (added 2026-04-19)

- In `budget_planner/workbook.py:_build_all_sheets` (or wherever sheet creation is orchestrated), move `wb.create_sheet(SHEET_TIMELOG_PLAN)` to slot 2 and `wb.create_sheet(SHEET_ACTUALS_HOURS)` to slot 3.
- Worksheets must still be built in **dependency order** (builders depend on other builders' outputs, e.g., Summary depends on Plan + Actuals). Creation order and build order are distinct — `create_sheet` just positions the tab; `_build_all_sheets` invokes builders in their existing order.
- On rebuild of a pre-v2 workbook, reorder existing sheets by reassigning `wb._sheets` or using `wb.move_sheet()` after all builders run. Idempotent — running on an already-v2 workbook is a no-op.

### 5.9 First log + Last log bug fixes (added 2026-04-19)

**Root cause:** `MATCH(1, (row<>"")*1, 0)` and `LOOKUP(2, 1/(row<>""), headers)` both depend on array-argument evaluation. openpyxl writes formulas as regular-entry (no CSE marker). LibreOffice silently returns `#N/A` (for MATCH) or drifts to a false-but-valid-looking neighbour (for LOOKUP). The `IFERROR` wrapper swallows MATCH's `#N/A` into `""` (First log always empty), and LOOKUP's drift surfaces as "last month" for everyone.

**Fix — First log:**
```
=IFERROR(INDEX(headers, SUMPRODUCT(MIN((row<>"")*month_idx + (row="")*99999))), "")
```
- Empty cells contribute 99999 (sentinel), non-empty contribute their month index.
- `MIN` picks the smallest, which is the first non-empty month index.
- `SUMPRODUCT` wraps `MIN` to force array evaluation (the one documented-stable LO pattern in `.claude/rules/openpyxl-patterns.md`).
- All-empty row → MIN = 99999 → INDEX out of range → IFERROR returns `""`.

**Fix — Last log:**
```
=IFERROR(INDEX(headers, SUMPRODUCT(MAX((row<>"")*month_idx))), "")
```
- Empty cells contribute 0, non-empty cells contribute their month index.
- `MAX` picks the largest = last non-empty month index.
- All-empty row → MAX = 0 → INDEX out of range → IFERROR returns `""`.

Both are arithmetic-only (no IF, no LOOKUP, no MATCH). SUMPRODUCT forces the inner array to evaluate. Pattern survives LO recalc across the versions the user runs.

---

## 6. Formula specification (consolidated)

Reference key:
- `R1, Rn` = first / last person row of merged ASSIGNED grid.
- `S1, Sn` = first / last person row of PER-PERSON SUMMARY (same persons, different rows).
- `M1col, MNcol` = first / last month col letters (= `B, ...`).
- `RATE_col` = col I on PER-PERSON SUMMARY.
- `hdr` = month-header row (on merged grid's year/month axis).
- `midx` = hidden month-index helper row (on merged grid).

### PER-WP SUMMARY (WP row `r`)

| Col | Formula | Notes |
|---|---|---|
| A `WP code` | literal | `BLACK_FORMULA` |
| B `Ceiling €` | `=Config!E{cfg_wp_row}` | `GREEN_XREF`, `€#,##0` |
| C `Assigned €` | `=ROUND(SUMPRODUCT((M1col{R1}:MNcol{Rn}="WPk") * RATE_col{S1}:RATE_col{Sn}), 0)` | `BLACK_FORMULA`, `€#,##0` |
| D `Buffer €` | `=C{r} - B{r}` | `BLACK_FORMULA`, `€#,##0`; CF green>0, red<0 |
| E `Assigned months` | `=SUMPRODUCT(1*(M1col{R1}:MNcol{Rn}="WPk"))` | `BLACK_FORMULA`, `0` |
| F `WP sched months` | literal int | `BLACK_FORMULA`, comment explains source |
| G `Burn rate needed` | `=IF(AND(D{r}<0, MAX(0,{wp.active_end}-current_month)>0), -D{r}/MAX(0,{wp.active_end}-current_month), "")` | `BLACK_FORMULA`, `€#,##0`; CF red when `> MAX(RATE_col)` |
| H `Any OOW?` | `=SUMPRODUCT((M1col{R1}:MNcol{Rn}="WPk")*((M1col{midx}:MNcol{midx}<{wp.active_start})+(M1col{midx}:MNcol{midx}>{wp.active_end})))` | `BLACK_FORMULA`, `0`; CF red when >0 |
| I … | `=IF(<overlap at build time>, SUMPRODUCT((year_subrange="WPk")*RATE_col), "")` | `BLACK_FORMULA`, `€#,##0`; one col per calendar year |

### PER-PERSON SUMMARY (person row `s`)

| Col | Formula | Notes |
|---|---|---|
| A `Person` | `=Config!A{cfg_row}` | `GREEN_XREF` |
| B `Assigned` | `=COUNTIF(M1col{r_grid}:MNcol{r_grid}, "WP*")` | `r_grid` = matching row in merged grid; `BLACK_FORMULA`, `0` |
| C `Pool` | `=ROUND(Config!B{cfg_row}*project_months,0)` | `GREEN_XREF`, `0` |
| D `Over-pool?` | `=B{s}>C{s}` | `BLACK_FORMULA`; CF red when TRUE |
| E `Switches` | `=SUMPRODUCT(1*(M1col{r_grid}:M(N-1)col{r_grid}<>M2col{r_grid}:MNcol{r_grid}))` | `BLACK_FORMULA`, `0` |
| F `First log` | `=IFERROR(INDEX(hdr_range, SUMPRODUCT(MIN((M1col{r_grid}:MNcol{r_grid}<>"")*midx + (M1col{r_grid}:MNcol{r_grid}="")*99999))), "")` | `BLACK_FORMULA` |
| G `Last log` | `=IFERROR(INDEX(hdr_range, SUMPRODUCT(MAX((M1col{r_grid}:MNcol{r_grid}<>"")*midx))), "")` | `BLACK_FORMULA` |
| H `Largest gap` | placeholder `""` | deferred from v1; comment in cell |
| I `Rate` | `=VLOOKUP(A{s}, roster_lookup, 4, FALSE)` | `BLACK_FORMULA`, `€#,##0` |

### Merged grid — no formulas
Every cell holds a literal. Initial fill via solver; user-edit via direct input. Empty-but-in-window cells hold `""`; empty-out-of-window cells hold `""` + `DARK_GREY_FILL`. Pastel/red/grey fills via CF + direct fill (same rules as v1).

### Row 2 traffic-light formulas
- `Any undershoot?` (col 2): `=COUNTIF(D{wp_first}:D{wp_last}, "<0")`
- `Any out-of-window cell?` (col 4): `=SUM(H{wp_first}:H{wp_last})`
- `Any over-pool person?` (col 6): `=COUNTIF(D{p_first}:D{p_last}, TRUE)` — Over-pool? moved to col D
- `Total switches` (col 8): `=SUM(E{p_first}:E{p_last})` — Switches moved to col E
- `Re-run solver on next rebuild?` (col 10): literal `TRUE` on init; read as-is on rebuild

---

## 7. Phased implementation plan

Each phase sized to ≈ one focused session. Ordered to minimise re-work — every formula lands in its final form on first write. Each phase ends with `pytest tests/ -q` green + recalc pass; session can hand off cleanly at the boundary.

**Re-work avoidance**: the biggest reshuffle (PER-WP SUMMARY schema) is pulled into Phase 1 so Undershoot / `Window months remaining` / First-log / Last-log never get written in their interim-v1 form during v2 implementation. Phase 3 narrows to year cols (purely additive); Phase 4 narrows to tab reorder only.

### Phase 1 — Big layout refactor (merged grid + PER-WP SUMMARY redesign + log-formula fix)

**Goal:** final-form Timelog Plan tab, minus year cols and toggle-gated rebuild. Every formula written once, in its final shape.

Scope is deliberately large to prevent re-write: `_write_wp_summary_row`'s rewrite touches col D (Buffer € vs. old Overshoot), col E (Assigned months shifts in from F), col F (new WP sched months), col G (Burn rate vs. old `Window months remaining`). Doing these piecewise across Phases 1 + 3 would have meant writing col E / F / G twice.

**Files touched:**
- `budget_planner/layout.py` — `TIMELOG_MONTH_COL_START: 4 → 2`; remove `TIMELOG_RATE_HELPER_COL`.
- `budget_planner/timelog_plan.py`:
  - Delete PROPOSED-band code path.
  - Rename ASSIGNED → merged grid; every cell is a literal (no `=<Proposed>` formula).
  - Remove `_write_underload_cell`, `_write_summary_underload_cell`, `_UNDERLOAD_RED_FONT`, `_UNDERLOAD_GREEN_FONT`, and Rule 6 CF.
  - Remove `_write_rate_helper_cell`; add `_write_summary_rate_cell` writing at PER-PERSON SUMMARY col I.
  - Reorder `_PERSON_SUMMARY_METRIC_HEADERS`: Assigned → col B; Pool/Over-pool/Switches/First/Last/Gap shift left by one; Rate at col I.
  - `_PERSON_SUMMARY_FIRST_METRIC_COL: 3 → 2`.
  - Freeze pane `D3 → B3`.
  - `_WP_SUMMARY_HEADERS` → `["WP code", "Ceiling €", "Assigned €", "Buffer €", "Assigned months", "WP sched months", "Burn rate needed", "Any OOW?"]`.
  - `_write_wp_summary_row` rewritten end-to-end:
    - Col C Assigned € retargets rate-col to PER-PERSON SUMMARY col I (rows `S1:Sn`).
    - Col D **Buffer €** = `=C{r} - B{r}` (replaces old Overshoot + Undershoot double-col).
    - Col E Assigned months (shifts from old col F).
    - Col F **WP sched months** = literal int `sum(s["end"] - s["start"] + 1 for s in wp_schedule if s["id"] == wp.code)`.
    - Col G **Burn rate needed** = IF-wrapped formula with build-time literals (replaces old `Window months remaining`).
    - Col H OOW formula stays (targets merged-grid cols).
  - CF rules on PER-WP SUMMARY:
    - Green-positive / red-negative on col D (Buffer €).
    - Red-excess on col G (Burn rate) — formula `=G{first} > MAX(I{p_first}:I{p_last})`.
    - Red on col H (OOW) — unchanged semantic, col unchanged.
    - **Delete** old Rule 4 (Undershoot>0 red CF on col E) — never written in v2.
  - Row-2 traffic lights written in final form:
    - `Any undershoot?` = `COUNTIF(D{wp_first}:D{wp_last}, "<0")` (never references the interim Undershoot col).
    - `Any out-of-window cell?` unchanged.
    - `Any over-pool person?` = `COUNTIF(D{p_first}:D{p_last}, TRUE)` — Over-pool? at new col D.
    - `Total switches` = `SUM(E{p_first}:E{p_last})` — Switches at new col E.
  - **Toggle cell** (row 2, cols 9-10) written as literal `TRUE` + styled `BLUE_INPUT`; named range `timelog_resolve_flag` added. Phase 1 writes the cell but rebuild doesn't yet gate on it — solver always runs (Phase 2 wires the gate). Clean interim state: cell visible, but flipping it does nothing until Phase 2 ships.
  - **First log / Last log formulas** written in final SUMPRODUCT-MIN / SUMPRODUCT-MAX form on first write (never in MATCH / LOOKUP form). Tests catch the LO-specific correctness from day 1.
- `budget_planner/reader.py` — `read_timelog_plan` returns `{"grid": {...}, "resolve_flag": bool}`; grid reader scans the merged grid's literal values; `resolve_flag` read but ignored by orchestrator until Phase 2.
- `budget_planner/workbook.py` — orchestrator consumes new reader return shape; solver still always runs (gate added in Phase 2).
- Tests — `test_timelog_plan.py` rewritten:
  - Layout assertions for final col positions (merged grid M1 at col B, Rate at col I of PER-PERSON SUMMARY, PER-WP SUMMARY 8-col final shape).
  - Buffer € sign semantics (positive when Assigned > Ceiling; negative when Ceiling > Assigned; CF green>0 / red<0).
  - WP sched months literal correctness (multi-segment WP respected).
  - Burn rate conditional: blank when Buffer ≥ 0; blank when wallclock = 0; numeric otherwise.
  - First log returns first non-empty month header on recalced workbook; `""` for empty row.
  - Last log returns last non-empty month header (NOT project last month) for partial-row persons; `""` for empty row.
  - Toggle cell exists at row 2 col 10 with value TRUE; named range resolves; no behaviour assertion (Phase 2 owns that).
  - Merged grid: every cell is a literal (not a formula).

**Verify:** `pytest tests/ -q` green (`timeout: 300000` per dev-workflow rule); `python scripts/recalc.py output/budget_plan.xlsx`; open in Excel/LO, eyeball Buffer € CF colours, Burn rate conditional, First/Last log cells for a partial-assignment row (THIS is the LO-specific case pytest on openpyxl alone cannot catch — follow the dev-workflow "Excel-specific bugs" protocol).

### Phase 2 — Toggle gate on rebuild

**Goal:** `timelog_resolve_flag` cell (already present from Phase 1) now actually gates rebuild behaviour.

**Files touched:**
- `budget_planner/workbook.py:_build_all_sheets` — branches on `resolve_flag`:
  - `True`: solver runs, grid written from solver output (Phase-1 behaviour).
  - `False`: solver skipped, grid written from reader's preserved dict; new persons / months (added to Config since last solve) get empty cells, not solver-filled.
- Tests: new `test_timelog_resolve_flag.py`:
  - init → edit grid cell via openpyxl → set toggle FALSE → rebuild → edit preserved.
  - same + set toggle TRUE → rebuild → edit wiped.
  - Pre-v2 workbook (no toggle cell, no named range) → reader defaults to TRUE → solver runs.

**Verify:** Phase-2 tests green + full suite. Small phase.

### Phase 3 — Year cols (additive PER-WP SUMMARY extension)

**Goal:** Y{yyyy} € cols appended to PER-WP SUMMARY. Purely additive — **no Phase 1 formulas are modified**.

**Files touched:**
- `budget_planner/timelog_plan.py`:
  - `_WP_SUMMARY_HEADERS` extended with dynamic year labels computed from `project_start` + `project_months` (via existing `_calendar_year_groups`).
  - `_write_wp_summary_row` appends year-col formulas after col H — IF-wrapped SUMPRODUCT with build-time overlap check.
  - `_write_section_title(last_col=...)` bumps from 8 to `8 + num_years`.
- Tests: extend `test_timelog_plan.py`:
  - Year col count = distinct calendar years spanned by project.
  - Year col header format (`Y2026 €`, `Y2027 €`, …).
  - Non-overlap cells are `""`.
  - Sum of year cells across a WP row = that WP's Assigned € (cross-check).

**Verify:** `pytest tests/ -q`; recalc; eyeball year breakdown adds up to Assigned €.

### Phase 4 — Tab reorder

**Goal:** Timelog Plan → position 2; Actuals-Hours → position 3.

**Files touched:**
- `budget_planner/workbook.py` — after all builders run, `wb.move_sheet(SHEET_TIMELOG_PLAN, offset=<delta>)` and `wb.move_sheet(SHEET_ACTUALS_HOURS, offset=<delta>)`. Idempotent.
- Pre-check: grep `docs/**`, `tests/**`, `budget_planner/**` for tab-index lookups (`wb.worksheets[N]`, `wb.sheetnames[N]`) — patch any that would now resolve to a different sheet. Expected: none (project convention is name-based access).
- Tests: new `test_tab_order.py` — `assert wb.sheetnames[0:3] == [SHEET_CONFIG, SHEET_TIMELOG_PLAN, SHEET_ACTUALS_HOURS]`.

**Verify:** `pytest tests/ -q`; recalc; eyeball workbook tab order in Excel/LO.

### Phase 5 — Tests, docs, final sweep

**Goal:** close out the plan; update living docs per CLAUDE.md critical rule #6.

**Files touched:**
- `docs/design-invariants.md` — add entries:
  - "Timelog Plan's merged grid holds literal WP codes, not formulas. Rebuild behaviour is governed by `timelog_resolve_flag` (on the Timelog Plan tab, row 2 col 10)."
  - "`MATCH(1, arr, 0)` and `LOOKUP(2, 1/arr, headers)` patterns are banned — they rely on CSE evaluation which openpyxl-written workbooks don't carry. Use `SUMPRODUCT(MIN(…))` or `SUMPRODUCT(MAX(…))` instead."
- `docs/internals.md` § Sheet Builders — update Timelog Plan entry to note the merged-grid shape + toggle.
- `docs/directory-map.md` — no new files; just confirm `timelog_plan.py` / `timelog_solver.py` entries remain accurate.
- `docs/architecture.md` Mermaid tab-order diagram — update to reflect Timelog Plan / Actuals-Hours as positions 2/3.
- `CLAUDE.md` trigger table — update Timelog Plan entry to point at this design doc, not the archived v1.
- `.claude/rules/openpyxl-patterns.md` — add explicit note about the MATCH/LOOKUP pitfall (with "see Timelog Plan First log / Last log" reference) if not already there.
- Archive this design doc after Phase 5 completes (move to `docs/plans/archived/`).

**Verify:** `pytest tests/ -q` full suite green; `python scripts/recalc.py output/budget_plan.xlsx` clean; open in Excel/LO, eyeball all three summary bands + merged grid + toggle.

---

## 8. Tests to add / update

| Test | Phase | Asserts |
|---|---|---|
| `test_timelog_plan.py::test_layout_col_shift` | 1 | `TIMELOG_MONTH_COL_START == 2`; merged-grid M1 at col B; Rate at last col (I) of PER-PERSON SUMMARY |
| `test_timelog_plan.py::test_underload_removed` | 1 | No col-B Underload formula in merged grid or PER-PERSON SUMMARY |
| `test_timelog_plan.py::test_merged_grid_literal` | 1 | Every merged-grid cell is a literal value, not a formula |
| `test_timelog_plan.py::test_toggle_cell_present` | 1 | Row 2 col 10 holds literal TRUE; named range `timelog_resolve_flag` resolves to that cell |
| `test_timelog_plan.py::test_buffer_signed` | 1 | Buffer € = C - B; green CF on positive, red on negative |
| `test_timelog_plan.py::test_wp_sched_months_literal` | 1 | Literal matches `sum(segment_lengths)`; WP3 multi-segment respected |
| `test_timelog_plan.py::test_burn_rate_conditional` | 1 | Blank when buffer ≥ 0; blank when wallclock = 0; numeric = -buffer/wallclock otherwise |
| `test_timelog_plan.py::test_first_log_first_non_empty` | 1 | Recalced workbook: for a person with assignments starting at Mk, First log returns that Mk month-header value |
| `test_timelog_plan.py::test_last_log_last_non_empty` | 1 | Recalced workbook: for a person with assignments ending at Mj < MN, Last log returns Mj's header (NOT MN) |
| `test_timelog_plan.py::test_first_last_log_empty_row` | 1 | Person with no assignments → both return `""` |
| `test_timelog_plan_roundtrip.py::test_init_solver_output_preserved` | 1 | Init + recalc + reread → grid reflects solver output |
| `test_timelog_resolve_flag.py::test_flag_true_wipes_edits` | 2 | Edit grid cell; rebuild with flag=TRUE; edit overwritten |
| `test_timelog_resolve_flag.py::test_flag_false_preserves_edits` | 2 | Edit grid cell; rebuild with flag=FALSE; edit preserved |
| `test_timelog_resolve_flag.py::test_flag_missing_defaults_true` | 2 | Pre-v2 workbook (no toggle cell) → reader defaults flag to TRUE → solver runs |
| `test_timelog_plan.py::test_year_cols_count_and_headers` | 3 | Year col count = `num_calendar_years`; headers match `Y{yyyy} €` |
| `test_timelog_plan.py::test_year_cols_overlap_blanks` | 3 | Cells where WP active window doesn't overlap year are `""` |
| `test_timelog_plan.py::test_year_cols_sum_matches_assigned` | 3 | Sum across year cols on each WP row = that WP's Assigned € |
| `test_tab_order.py::test_config_timelog_actuals_hours_first` | 4 | `sheetnames[0:3] == [SHEET_CONFIG, SHEET_TIMELOG_PLAN, SHEET_ACTUALS_HOURS]` |

All tests must exercise **recalculated** workbooks per CLAUDE.md Rule 2 — openpyxl round-trip without recalc masks the LO-specific bugs Phase 4 fixes.

---

## 9. Docs to update (per CLAUDE.md rule #6)

| File | Change |
|---|---|
| `docs/design-invariants.md` | Add: "Timelog Plan merged grid = literal values; toggle governs re-solve." + "Banned: `MATCH(1, arr, 0)` / `LOOKUP(2, 1/arr, hdr)` — use `SUMPRODUCT(MIN/MAX(…))`" |
| `docs/internals.md` | Sheet Builders § Timelog Plan — describe merged grid + row 2 toggle |
| `docs/architecture.md` | Mermaid tab-order diagram: Config → Timelog Plan → Actuals-Hours → … |
| `docs/directory-map.md` | Confirm timelog_plan.py / timelog_solver.py entries still accurate |
| `CLAUDE.md` trigger table | Timelog Plan row points at this doc (once archived, point at `docs/plans/archived/2026-04-19-timelog-plan-v2-design.md`) |
| `.claude/rules/openpyxl-patterns.md` | Reference MATCH/LOOKUP ban with Timelog Plan as concrete example |

Grep before archive (per "Superseded docs: consume before archive" global rule): pull load-bearing content from this doc into `design-invariants.md` + `internals.md` before the final `git mv`.

---

## 10. Decisions log (2026-04-19 brainstorm)

| # | Decision | Rationale |
|---|---|---|
| D1 | Merge PROPOSED + ASSIGNED into one editable grid | v1's `=<Proposed>` layer added visual complexity for round-trip preservation the user rarely exercises; re-solve is the common path |
| D2 | Workbook-level "Re-run solver?" toggle, TRUE by default, no auto-reset | User solves often; manual flip to FALSE is the expected gesture when preserving edits |
| D3 | Toggle lives on Timelog Plan row 2 (not Config) | Context-local; visible every time the user opens the tab |
| D4 | Col-C collapse abandoned in favour of moving Rate to PER-PERSON SUMMARY | Rate needed only once per person; previous design duplicated it per band; moving it to the summary band removes the hidden-column clutter entirely |
| D5 | Merged grid starts at col B (no per-row metric col) | Underload removed as rate-blind; no replacement metric earns the column |
| D6 | "Months to fill" rejected; `Burn rate needed` chosen instead | Month-count drifts when persons on one WP have different rates; burn rate stays in EUR units the user actually optimizes |
| D7 | Buffer € signed, green+/red− | Single column conveys the two states with domain-native semantics (overshoot = buffer = good) |
| D8 | WP sched months as build-time literal | Matches existing literal pattern (active_end hardcoded); user rebuilds anyway when Config changes |
| D9 | Per-WP × per-year EUR added to PER-WP SUMMARY (Option A) vs. standalone band (Option B) | "Add to some existing table" per user preference; WPs already on rows there; sum-across-years sanity-checks against Assigned € |
| D10 | First log / Last log fixed in this plan (not deferred) | LO-specific; current formulas silently drift; fix is one-line per formula with the SUMPRODUCT pattern |
| D11 | Tab reorder: Timelog Plan → 2, Actuals-Hours → 3 | Timelog Plan is the first-touch sheet each session; Actuals-Hours is its source |
| D12 | No "Last solver run" timestamp cell | Over-engineering given no-auto-reset toggle; can add later if opacity becomes a real issue |
| D13 | `MATCH(1, arr, 0)` / `LOOKUP(2, 1/arr, hdr)` globally banned for openpyxl-written workbooks | Root cause of First log / Last log is generic; document the pattern to prevent recurrence |
| D14 | Phase 1 absorbs PER-WP SUMMARY schema redesign + First/Last log formula fix (vs. original plan that deferred them to Phase 3 / Phase 4) | Re-work audit found 6 cells / formulas / CF rules that would be written twice under the original ordering. Pulling the schema redesign forward collapses those to one write each. Phase 1 grows modestly; Phases 3-4 shrink to purely-additive / isolated-change scope. |
