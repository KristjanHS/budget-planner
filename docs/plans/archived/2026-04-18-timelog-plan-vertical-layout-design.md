# Timelog Plan — Vertical-Stacked Layout Design

Date: 2026-04-18
Status: Approved — ready for implementation

## Problem

The Timelog Plan sheet currently renders three data sections side-by-side across a very wide canvas:

```
[A=Person | B=Rate] | PROPOSED grid (months) | gap | ASSIGNED grid (months) | gap | PER-PERSON SUMMARY (8 cols)
```

For a long project this produces a wide canvas. Users must scroll horizontally to compare PROPOSED vs ASSIGNED vs SUMMARY for the same person; the PER-PERSON SUMMARY sits far right and is often off-screen even on wide monitors. There is no way to hide a section the user isn't currently looking at.

## Goals

1. Stack PROPOSED, ASSIGNED, and PER-PERSON SUMMARY as three vertical row-bands (one above the other), so the sheet reads top-to-bottom rather than left-to-right.
2. Make each band a collapsible Excel row-group (outline level 1) so the user can fold away sections they aren't currently working on.
3. Surface an at-a-glance **Underload** indicator next to the Person name, frozen together with col A so it stays visible when scrolling month columns.
4. Move PER-WP SUMMARY to the top of the sheet (directly below the traffic-lights row) so "overview / health" info is gathered together and data bands are all below.
5. Keep all formula shapes (SUMPRODUCT, COUNTIF, INDEX/MATCH) and cross-sheet semantics unchanged — only row/column coordinates move.

## Non-goals

- No change to the solver (`budget_planner/timelog_solver.py`). Proposed cells still come from the same greedy two-pass pipeline.
- No change to the Config-side trust boundary, blue-input conventions, or round-trip override preservation contract.
- No merging of PROPOSED and ASSIGNED into a single visual band (e.g. toggleable overlay) — both remain discrete grids.
- No new metrics beyond Underload; the existing 7 per-person summary columns (after dropping duplicated Remaining) and 8 per-WP summary columns stay.

## In scope (files touched)

- `budget_planner/timelog_plan.py` — core builder restructure
- `budget_planner/layout.py` — add `TIMELOG_MONTH_COL_START`, retire `timelog_assigned_col_start/end`, `timelog_summary_col_start`, `TIMELOG_GAP_WIDTH`; change `TIMELOG_RATE_HELPER_COL` from 2 → 3
- `budget_planner/reader.py` — `read_timelog_plan` switches from offset-based lookup to section-title anchor
- Tests: `test_timelog_plan.py`, `test_reader_v2.py` (if it touches Timelog round-trip)
- Docs: `docs/internals.md` (Timelog Plan layout paragraph + extraction-functions table rows), `docs/directory-map.md` (if it cites row/col positions)

## Out of scope

- Migrating any stored user data. This is a layout change only; a rebuild writes the new layout and re-parses any ASSIGNED-grid overrides from the previous workbook's (horizontal) layout via the reader's anchor-based logic — no offline migration script required because rebuild reads the old sheet cell-by-cell.
- Cross-sheet freeze-pane coordination with Plan / Capacity / Actuals. Those sheets remain horizontal.
- Splitting `build_timelog_plan_sheet` into per-band helper functions. Helper extraction can be a separate follow-up commit once this layout has stabilised.

---

## Section 1 — Row & column geometry

New column constant:

```python
TIMELOG_MONTH_COL_START = 4   # col D — one right of standard MONTH_COL_START
TIMELOG_RATE_HELPER_COL = 3   # col C (shifted right from old col B)
```

The standard `MONTH_COL_START = 3` (col C) applies to every other sheet; Timelog Plan is the only sheet with an extra left-column metric.

### Column layout per band

| Band | A | B | C | D … end | Further cols |
|---|---|---|---|---|---|
| PROPOSED | Person (green xref to Config roster) | Underload (formula) | Rate helper (VLOOKUP; collapsible via `outline_level=1`) | month grid | — |
| ASSIGNED | Person (green xref) | Underload (formula) | Rate helper (VLOOKUP; collapsible) | month grid | — |
| PER-PERSON SUMMARY | Person (green xref) | Underload (= Pool − Assigned) | Assigned | Pool | E=Over-pool?, F=Switches, G=First log, H=Last log, I=Largest gap |
| PER-WP SUMMARY | WP code | Ceiling € | Assigned € | Overshoot € | E=Undershoot €, F=Assigned months, G=Window months remaining, H=Any out-of-window? |

**Column-count note:** PER-PERSON SUMMARY metrics drop to 7 (removed duplicated `Remaining`, since col B's Underload is semantically identical).

### Row map (10-person / 14-month / 5-WP example)

| Row(s) | Content | Outline |
|---|---|:---:|
| 1 | Banner (merged A..AV) | 0 |
| 2 | Traffic lights (A..H, 4 label+value pairs) | 0 |
| 3 | blank spacer | 0 |
| 4 | **"PER-WP SUMMARY"** title (merged A..H) | 0 |
| 5 | WP summary column headers | 1 |
| 6–10 | 5 WP rows | 1 |
| 11 | blank spacer | 0 |
| 12 | **"PROPOSED (solver-written, locked)"** title (merged A..AV) | 0 |
| 13 | Year header | 1 |
| 14 | Month header | 1 |
| 15 | Hidden month-idx helper (also `hidden=True`) | 1 |
| 16–25 | 10 person rows | 1 |
| 26 | blank spacer | 0 |
| 27 | **"ASSIGNED (edit here — default =Proposed)"** title (merged A..AV) | 0 |
| 28 | Year header | 1 |
| 29 | Month header | 1 |
| 30 | Hidden month-idx helper | 1 |
| 31–40 | 10 person rows | 1 |
| 41 | blank spacer | 0 |
| 42 | **"PER-PERSON SUMMARY"** title (merged A..I) | 0 |
| 43 | Metric column headers | 1 |
| 44–53 | 10 person rows | 1 |

**Fully collapsed view** (all four outline groups folded): rows 1, 2, 4, 12, 27, 42 — banner + traffic lights + 4 section titles.

### Freeze panes

```python
ws.freeze_panes = ws.cell(row=3, column=TIMELOG_MONTH_COL_START).coordinate   # "D3"
```

Freezes rows 1–2 (banner + traffic lights) + cols A–C (Person + Underload + Rate helper). Month grid scrolls horizontally; Underload stays visible next to Person at all times.

---

## Section 2 — Formula rewiring

The builder tracks band membership via:

```python
person_rows_by_band: dict[str, dict[str, int]] = {
    "proposed": {name: row, ...},
    "assigned": {name: row, ...},
    "summary":  {name: row, ...},
}
```

### ASSIGNED cell default

Each in-window month cell in ASSIGNED row `{a}` col `{c}` defaults to:
```
={col_letter(c)}{proposed_row}
```
Where `proposed_row = person_rows_by_band["proposed"][name]`. Round-tripped user overrides (from `assigned_overrides`) are written as `BLUE_INPUT` literals, same as today.

### Underload (col B, all three bands)

- PROPOSED / ASSIGNED band, person row `{r}`:
  ```
  =$D{summary_row} − COUNTIF(D{r}:AV{r}, "WP*")
  ```
  Forward-references PER-PERSON SUMMARY's col D (Pool) for that person. Excel/LO resolve forward refs at recalc — no chain issue.
- PER-PERSON SUMMARY band, person row `{s}`:
  ```
  =D{s} − C{s}
  ```
  (Pool − Assigned, using this band's own cols.)

Font colour CF on col B in each band:
- Red when value < 0 (over-pool)
- Green when value > 0 (underload)
- Black at 0 (fully loaded — default)

### Per-person summary formulas (cols C..I, PER-PERSON SUMMARY band)

Row-offset change only — formula shapes unchanged. Each targets the ASSIGNED band row for the same person:

| Col | Metric | Formula (schematic) |
|---|---|---|
| C | Assigned | `=COUNTIF(D{a}:AV{a}, "WP*")` |
| D | Pool | `=ROUND(Config!B{cfg_row} * project_months_ref, 0)` (green xref) |
| E | Over-pool? | `=C{s} > D{s}` (reordered — Pool now at D, Assigned at C) |
| F | Switches | `=SUMPRODUCT(1*(D{a}:AU{a}<>E{a}:AV{a}))` |
| G | First log | `=IFERROR(INDEX(month_headers, MATCH(1, (D{a}:AV{a}<>"")*1, 0)), "")` |
| H | Last log | `=IFERROR(LOOKUP(2, 1/(D{a}:AV{a}<>""), month_headers), "")` |
| I | Largest gap | `""` placeholder (existing deferred work; see builder docstring) |

Where `{a}` = ASSIGNED row for the person, `{s}` = this summary row.

### PER-WP SUMMARY formulas

- `Assigned €` SUMPRODUCT: rate column = **ASSIGNED band's col C** (must align row-for-row with the assigned grid). PROPOSED band's col C becomes informational/cosmetic — documented with a builder-code comment on the PROPOSED rate helper write.
- `Assigned months`, `Any out-of-window?`: target ASSIGNED band's month grid + its hidden month-idx row.
- `Ceiling €`, `Overshoot €`, `Undershoot €`, `Window months remaining`: unchanged.

### Traffic lights (row 2)

Formula shapes unchanged. Row/column ranges update to new band positions:

| # | Cell | Formula |
|---|---|---|
| 1 | Any undershoot? | `=COUNTIF(E{wp_first}:E{wp_last}, ">0")` |
| 2 | Any out-of-window cell? | `=SUM(H{wp_first}:H{wp_last})` |
| 3 | Any over-pool person? | `=COUNTIF(E{sum_first}:E{sum_last}, TRUE)` (Over-pool? now in col E) |
| 4 | Total switches | `=SUM(F{sum_first}:F{sum_last})` (Switches now in col F) |

### Conditional formatting

All rules re-target the new ranges; rule shapes unchanged. Added:

- **New Rule 6 — Underload font colour** (3 range-scoped rules, one per band):
  - `col_B_range` → `FormulaRule(formula=["B{first}<0"], font=red)`
  - `col_B_range` → `FormulaRule(formula=["B{first}>0"], font=green)`

Existing rules re-targeted:
- Rules 1–2 (out-of-window red + pastel per WP) — applied per-grid, now on two distinct row ranges for PROPOSED and ASSIGNED, cols D..AV.
- Rule 3 (Over-pool? = TRUE → red fill) — PER-PERSON SUMMARY col E range.
- Rules 4–5 (Undershoot € > 0 / Any out-of-window? > 0 → red) — new PER-WP SUMMARY row range.

---

## Section 3 — Outline grouping, layout.py cleanup, reader round-trip

### Outline grouping mechanics

- Section-title rows stay at `outline_level=0` (default) — always visible.
- All rows inside each band (sub-headers, hidden idx, data) get `ws.row_dimensions[r].outline_level = 1`.
- Blank spacer rows between bands stay at level 0 so they act as group separators.
- Empty-band guard: if `person_first_row > person_last_row` or `wp_first_row > wp_last_row`, skip the `outline_level=1` assignment for that band (openpyxl quirk with zero-row groups).

### `layout.py` cleanup

**Remove:**
```python
TIMELOG_GAP_WIDTH
timelog_assigned_col_start()
timelog_assigned_col_end()
timelog_summary_col_start()
```

**Add:**
```python
TIMELOG_MONTH_COL_START = 4
```

**Change:**
```python
TIMELOG_RATE_HELPER_COL = 3   # was 2
```

**Pre-removal grep:** `rg "timelog_assigned_col_start|timelog_assigned_col_end|timelog_summary_col_start|TIMELOG_GAP_WIDTH"` must return hits only inside `budget_planner/timelog_plan.py` (callsites rewritten in this same change) and `budget_planner/layout.py` (the definitions themselves).

### Reader round-trip (`read_timelog_plan`)

Switch from hardcoded row-offset lookup to section-title anchor search:

1. Scan col A downward; the ASSIGNED band starts at the first row whose value equals the exact title string `"ASSIGNED (edit here — default =Proposed)"` (or whatever the current builder writes — read from a shared constant to avoid drift).
2. Month header row = anchor + 2. Hidden-idx row = anchor + 3. First data row = anchor + 4.
3. Last data row = scan downward until an empty col A row. Person name comes from col A (it's a cross-sheet xref formula, so evaluate via `data_only=True` to get the resolved string).
4. Columns D..`D + project_months − 1` parsed the same way as today: open workbook with `data_only=False` to get formulas; any cell whose value does not match the expected `=<col_letter>{proposed_row}` default pattern is a blue-input override.
5. Return shape unchanged: `dict[tuple[str, int], str]` mapping `(person, month_num) → wp_code`.

Anchoring on the section title (not on row numbers) follows the drift-tolerant pattern from `docs/plans/archived/2026-04-12-rebuild-shape-drift-design.md`. If the banner, traffic lights, or PER-WP SUMMARY height changes in a future edit, the reader still finds ASSIGNED.

### Tests

**`test_timelog_plan.py`** — update cell-address assertions to match the new row/column layout. Specifically:
- Section titles land on new rows (derive expected rows from the builder's returned layout dict, not hardcoded).
- Per-person summary metric ordering changed — `Remaining` removed, Over-pool? now at col E, Switches at col F.
- Underload col B present in all three bands with correct formula shape.
- Outline levels: assert `ws.row_dimensions[r].outline_level == 1` for rows inside each group; `== 0` for section-title and spacer rows.
- Freeze panes == `"D3"`.

**`test_reader_v2.py`** (or equivalent Timelog round-trip test) — round-trip coverage:
- Init → openpyxl edit one ASSIGNED cell to a different WP code (blue-input override) → `rebuild_workbook` → reopen and assert the override survives at its new band location.

**Post-generation verify**: per `dev-workflow.md`, run `python scripts/recalc.py output/budget_plan.xlsx` before `pytest`, so `data_only=True` reads surface real formula results rather than `None`.

### Known edge cases

- **Empty team (persons=0):** each data band has zero person rows. Guard `outline_level=1` loop with a `person_first_row <= person_last_row` precondition.
- **Single person:** adjacent section-title rows land close together (title + 3 header rows + 1 data + blank spacer). Collapse still works; cosmetic only.
- **PER-WP SUMMARY with 0 WPs:** existing `_build_wps` filter leaves `wps == []` → `wp_first_row > wp_last_row` → skip grouping for that band. PER-WP SUMMARY section title still renders (acceptable — user sees an empty placeholder band rather than a missing one).

---

## Key decisions recap

| Decision | Chosen | Alternative considered | Why |
|---|---|---|---|
| Band ordering | PER-WP SUMMARY → PROPOSED → ASSIGNED → PER-PERSON SUMMARY | Traffic lights + PER-WP SUMMARY both at top | Groups all "overview / health" info together above the data bands |
| Header repetition | Each band self-contained (Person col + year/month headers repeated) | Shared headers at top only | Each band readable standalone when scrolled |
| Collapse semantics | Section-title row stays visible, everything else collapses | Whole band (incl. title) collapses | Compact fully-collapsed overview = 4 visible section bars |
| Underload bands | All three (PROPOSED, ASSIGNED, PER-PERSON SUMMARY) | ASSIGNED only | Self-contained bands per Q1; user accepts duplicate-with-Remaining trade-off |
| `Remaining` column | Dropped from PER-PERSON SUMMARY metrics | Kept (duplicated) | DRY — col B Underload carries the same value |
| Rate helper col | Kept at col C (shifted from B) | Drop and inline VLOOKUP into PER-WP SUMMARY SUMPRODUCT | Preserves collapsible rate column; minimal disruption |
| Rate helper canonical band | ASSIGNED (row-aligned with consumer) | PROPOSED | PER-WP SUMMARY's SUMPRODUCT consumes rate alongside assigned grid — row alignment required |
| Builder shape | Restructure monolithic function | Split into per-band helpers | Reduces review surface; helper extraction can be separate follow-up |

## Follow-ups (explicitly deferred)

- Extract per-band emitter functions (`_emit_wp_summary_band`, `_emit_proposed_band`, `_emit_assigned_band`, `_emit_person_summary_band`). Do after this layout change has stabilised.
- Implement `Largest gap` summary column (placeholder today). Needs hidden streak-length helper row + MAX. Deferred until concrete user need surfaces — documented in existing builder docstring.
- Consider per-grade tinting of Underload col B (e.g. amber halo for senior grades with large under-utilisation). Speculative; not in this scope.
