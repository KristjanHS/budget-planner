# Timelog Plan — Design

**Status:** Validated via brainstorming 2026-04-18. Ready for implementation plan.
**Scope:** New sheet `Timelog Plan` + solver module. Separate allocation surface from the existing Plan sheet.

## 1. Purpose & scope

A prescriptive, whole-month, 1.0-FTE timesheet allocation plan. For each person and each calendar month, the sheet tells the person which WP to log 100% against. The goal is to fill every WP to (or slightly past) its EUR budget ceiling before its active window closes, so auditor haircuts and partner spillover stay absorbable.

Deliberately decoupled from the existing Plan sheet:
- **Plan** = actual effort (what people *do*).
- **Timelog Plan** = timesheet-compliant allocation (where people *log*).
- The two may legitimately diverge.

**Non-goals (YAGNI):**
- No sub-month granularity — whole-month 1.0-FTE blocks are the point.
- No write-back to Actuals or Plan.
- No multi-project view.
- No saved what-if scenarios (use existing `scripts/snapshot.py`).

## 2. Sheet layout

Tab: `Timelog Plan`, colour orange. Placement between `Plan` and `Actuals`.

- **Row 1:** merged title banner.
- **Row 2:** traffic-light cells — `Any undershoot?`, `Any out-of-window cell?`, `Any over-pool person?`, `Total switches`.
- **Rows 4-5:** year/month header band (same cadence as Plan).
- **Proposed grid** (left): section header `PROPOSED (Python-computed, locked)`. Rows = persons, ordered by priority ascending then roster order. Col A = person name (green cross-ref to Team Roster). Cells = WP codes, locked, pastel fill per WP via CF.
- **Gap column** (separator).
- **Assigned grid** (right, mirrors Proposed shape): section header `ASSIGNED (edit here — start as =Proposed)`. Cells start as `=<Proposed cell>` (black text on blue-input fill); user overrides become hard blue-input values.
- **Per-person summary block** (right of Assigned): `Assigned count, Pool, Remaining, Over-pool?, Switches, First log, Last log, Largest gap`.
- **Per-WP summary block** (below both grids): `WP code, Ceiling €, Assigned €, Overshoot €, Undershoot €, Assigned months, Window months remaining, Any out-of-window?`.

Freeze panes: rows 1-5, col A.

## 3. Config changes

**Exactly one change to Config** — the PERSON-WP ASSIGNMENTS table stays completely untouched (existing Plan autofill / Capacity / reader all keep working unchanged).

**Team Roster: new column `Priority`** (integer 1-3, blue input).
- Default on init / fresh rebuild: `2` (neutral).
- Header comment explains the solver's use.
- Named range: `team_priorities`.
- `reader.py` Team Roster reader: add `priority: int` to returned rows; missing/blank → `2` (backward-compat).

Other Config inputs reused as-is:
- **Team Roster.Max FTE** → per-person pool.
- **WP Budget.Ceiling €** → per-WP target.
- **WP Schedule** (start / end month per WP) → hard WP active window.
- **PERSON-WP ASSIGNMENTS** (Person, WP, Start Month, End Month) → per-(person,WP) hard window; solver search space.
- **Grade rate table** → per-person €/month via existing lookup.

## 4. Solver algorithm (two-pass greedy)

Module: `budget_planner/timelog_solver.py`. Pure functions, deterministic, pure-Python (no new deps).

**Derived inputs:**
- `pool[p] = round(max_fte[p] × project_months)`
- `rate[p] = grade_rate[p] × hours_per_fte_month`

**Pass 1 — Greedy fill, WPs in earliest-close order:**

```
assigned[p][m] = None
for wp in sorted(wps, key=w.active_end_month):
    for prio in (1, 2, 3):
        candidates = [p for p in persons
                      if p.priority == prio
                      and has_person_wp_row(p, wp)     # hard constraint
                      and pool[p] > 0]
        candidates.sort(key=(-is_running(p, wp), p.name))  # continue-a-run
        for p in candidates:
            window = intersect(wp.active, person_wp_window(p, wp))
            for m in window:
                if assigned[p][m] is None and pool[p] > 0:
                    assigned[p][m] = wp.code
                    pool[p] -= 1
                    if wp_filled(wp): break   # allows slight overshoot
            if wp_filled(wp): break
```

Check-after-placing yields exactly-one-month overshoot when it happens — matches the "buffer for auditor haircut" requirement.

**Pass 2 — Switch-reduction local improvement:**

```
repeat until no swap found:
  for each person p:
    for each adjacent month pair (m, m+1):
      if assigned[p][m] != assigned[p][m+1]:
        find a legal swap with another (p2, m2) that
        preserves all WP sums, windows, pools
        and reduces total switches
```

Deterministic, bounded, no randomness → reproducible tests.

**Output:** `dict[person_name][month_idx] → wp_code | None`, written to the Proposed grid.

## 5. Feedback formulas + CF

**Per-WP summary:**
| Column | Formula |
|---|---|
| `Ceiling €` | green cross-ref to Config WP Budget ceiling cell |
| `Assigned €` | `=SUMPRODUCT((assigned_grid="WP2")*person_rate_col)` |
| `Overshoot €` | `=MAX(0, Assigned - Ceiling)` — no CF |
| `Undershoot €` | `=MAX(0, Ceiling - Assigned)` — **red CF if > 0** |
| `Assigned months` | `=SUMPRODUCT((assigned_grid="WP2")*1)` |
| `Window months remaining` | `=MAX(0, wp_end - current_reporting_month)` |
| `Any out-of-window?` | `=SUMPRODUCT(...)` counting cells outside `[wp_start, wp_end]` — **red CF if TRUE** |

**Per-person summary:**
| Column | Formula |
|---|---|
| `Assigned count` | `=COUNTIF(person_row, "WP*")` |
| `Pool` | `=ROUND(max_fte × project_months, 0)` |
| `Remaining` | `=Pool - Assigned count` |
| `Over-pool?` | `=Assigned > Pool` — **red CF** |
| `Switches` | `=SUMPRODUCT(1*(person_row_shifted <> person_row))` — counts ALL adjacent diffs (empty↔WP included) |
| `First log` | `=INDEX(month_headers, MATCH(TRUE, person_row<>"", 0))` array |
| `Last log` | `=LOOKUP(2, 1/(person_row<>""), month_headers)` |
| `Largest gap` | hidden streak-length helper row + `=MAX(...)` |

**Canvas CF rules:**
- One rule per WP: `cell = "WPk"` → pastel fill (light-blue, light-green, light-yellow, …). Constants live in `styles.py`.
- Per-row rule: `cell = "WPk" AND (month_col < wp_k.start OR month_col > wp_k.end)` → red border.
- Empty cell inside person's project window → light-grey fill.
- Empty cell outside project window → dark-grey fill.

Respect openpyxl/LO gotchas from `.claude/rules/openpyxl-patterns.md`:
- `CF_*_FILL` constants (8-char AARRGGBB) for all CF.
- Per-row rules when external refs vary (WP end months differ).
- Most critical fill added first (red before pastels).
- SUMPRODUCT + equality test is the one SUMPRODUCT pattern that survives LO recalc.

## 6. Build / read / rebuild integration

**New files:**
- `budget_planner/timelog_plan.py` — stateless builder `build_timelog_plan_sheet(...)`; returns layout dict.
- `budget_planner/timelog_solver.py` — pure `greedy_pass(...)`, `switch_reduction_pass(...)`.

**Modified files:**
- `budget_planner/layout.py` — add `TIMELOG_PLAN_SHEET`, column offsets, section labels.
- `budget_planner/workbook.py:_build_all_sheets` — call builder after `build_plan_sheet`.
- `budget_planner/reader.py` — add `read_timelog_plan(wb)` returning `assigned_overrides` dict; add `priority` to Team Roster reader.
- `budget_planner/config_sheet.py` — Team Roster: add Priority column + named range.

**Rebuild flow:**
1. Reader extracts `assigned_overrides` from the existing workbook (cells the user typed literal values into, not `=Proposed!...` formulas).
2. Builder runs solver → writes Proposed grid (always fresh).
3. Builder writes Assigned grid: `=<Proposed cell>` by default; if `(person, month) ∈ assigned_overrides`, write the override value as blue-input literal instead.

Missing sheet / missing named range on read → empty dict. Backward-compatible with pre-feature workbooks.

## 7. Tests

Per `.claude/rules/dev-workflow.md` round-trip rule.

- `tests/test_timelog_plan.py` — structural: sheet exists, row/col count from config, summary formulas present, CF rule count/pattern match (by fill colour + formula pattern, not index).
- `tests/test_timelog_solver.py` — solver determinism on fixture inputs; prio-1 fills earliest-close WP rows first; pool never exceeded; undershoot formula fires when budget under-filled; overshoot tolerated.
- `tests/test_timelog_plan_roundtrip.py` — init → edit one Assigned cell via openpyxl → `rebuild_workbook` → reopen → assert override preserved, Proposed regenerated.
- Post-build: `python scripts/recalc.py output/budget_plan.xlsx` then `pytest tests/ -v` (CLAUDE.md critical rule #2).

## 8. Docs to update before merge

Per CLAUDE.md critical rule #6 + dev-workflow invariants:
- `docs/design-invariants.md` — "Timelog Plan is a separate allocation surface from Plan; both are editable inputs and may legitimately diverge."
- `docs/internals.md` — add builder to Sheet Builders section; add new § Solvers.
- `docs/directory-map.md` — `timelog_plan.py`, `timelog_solver.py`.
- `docs/architecture.md` — add sheet to Mermaid tab-order diagram.
- `CLAUDE.md` trigger table — `solver, timelog, per-person 1.0 FTE logging → timelog_plan.py + timelog_solver.py`.

## 9. Design decisions log

Resolved during 2026-04-18 brainstorming:

| # | Decision | Rationale |
|---|---|---|
| Q1 | Whole-month 1.0-FTE blocks, one WP per person-month | Indivisible month = simple grid, clear timesheet rule |
| Q2 | Pool = `max_fte × project_months` from Team Roster | Only limiter; no per-person date windows today |
| Q3 | Hybrid: canvas + Python solver | Formula feedback, solver provides starting point |
| Q4 | Persons × months grid, cell = WP code | Most compact; fewest-switches becomes visual |
| Q5 | Side-by-side Proposed (locked) + Assigned (=Proposed default) | Preserves user edits across rebuild |
| Q6 | High-level aggregate feedback; per-WP + per-person summary blocks | Matches Config helper-row style |
| Q6a | Cell-level red CF marker for out-of-window | User wants visual at cell level too |
| Q7 | Two-pass heuristic (greedy + swap) | Near-optimal, pure-Python, deterministic |
| Q8 | Reuse existing WP Budget Ceiling € column | Already in Config |
| Q9 | Pastel fill per WP; light-grey empty in-window; dark-grey out-of-window | Switches = colour changes, visually obvious |
| Q10 | Availability = `max_fte × project_months` only, no per-person start/end | Team Roster doesn't carry those today |
| Q10+ | Team Roster gets new `Priority` column (1-3) | New input for solver ordering |
| Q11 | Prio-1 WP preferences auto-derived from existing PERSON-WP ASSIGNMENTS | Reuse table; no duplicate input |
| Q12 | All priorities use PERSON-WP ASSIGNMENTS as hard constraint; priority orders fill sequence | Keeps solver search space aligned with effort-plan graph |
| Q13 | Earliest-close WP dominates; priority is secondary tiebreaker within a WP pass | Stated optimization objective |
| Q14 | Slight overshoot allowed by design | Auditor haircut + partner spillover buffer |
| Q15a | Switch count includes empty↔WP transitions | "Any break in streak" semantics |
| Q15b | Undershoot is red (unused budget = real problem); overshoot is neutral | Money left on the table is the concern |
| Q16 | Pass-2 swap primitive is fixed at adjacent-pair `(p1, m, m+1) ↔ (p2, m, m+1)`; 3+ month block swaps and 3-way same-month-window rotations are not considered | Accepted as a known shape limitation — determinism + bounded runtime > optimality. Documented in `timelog_solver.py` module docstring (Stage 2B, 2026-04-18). |
