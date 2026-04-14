# Simplification / Refactoring Audit — 2026-04-13

Project-wide audit of `budget_planner/` (15 modules, ~7700 LOC) conducted via three
parallel explorer agents: reuse, quality, efficiency.

Goal: identify simplifications that **retain current functionality** (all 372
tests passing, same output workbook structure).

---

## Prior-art warning

A 2026-04-13 simplification sweep (see
`~/.claude/rules/python-refactors.md` "Dual-caller gotcha") broke ~15 fixtures
when `build_config_sheet` fallbacks were removed. **Do not retry** fallback
tightening in `build_config_sheet` / `build_team_sheet` / `build_reader_v2`
without grep-checking bare test callers first. Orchestrator-only builders are
fine.

---

## Tier A — Quick wins, low risk (~1 hour, ~50 LOC net reduction)

Apply directly; zero architectural change.

### A1. Precompute `wp_by_id` / `wp_index` maps in person-grid builders
- **Files:** `capacity.py:125–128, 167` and `actuals_hours.py:106–113` (mirrored)
- **Now:** Each person-assignment row does two linear scans over
  `work_packages` — `any(w["id"] == wp_id ...)` and
  `next((w for w in work_packages if w["id"] == wp_id), None)`.
- **Fix:** Hoist `wp_by_id = {w["id"]: w for w in work_packages}` and
  `wp_idx = {w["id"]: i for i, w in enumerate(work_packages)}` once at function
  entry. Replace scans with `.get(wp_id)` / `.get(wp_id, None)`.
- **Impact:** O(N²) → O(N) per person; ~10 LOC simpler. Collapses findings
  efficiency #1, #3, #10, #11.

### A2. De-duplicate `_wp_config_row` helper
- **Files:** `capacity.py:29–34` ≡ `actuals_hours.py:75–80` (verbatim).
- **Fix:** Keep one copy (put in `person_section.py` or `layout.py`); import
  from the other site. **But:** if A1 lands, this helper is largely replaced by
  the precomputed `wp_idx` map — consider deleting both copies and inlining
  `config_first_wp_row + wp_idx[wp_id]` at the 2–3 callsites.
- **Impact:** ~6–12 LOC.

### A3. Use `cell_ref()` in `plan.py:129` named range
- **File:** `plan.py:129` — `attr_text=f"{SHEET_PLAN}!$G$1"`
- **Fix:** `attr_text=cell_ref(SHEET_PLAN, 1, 7)` — matches the pattern used
  elsewhere (`workbook.py:281`, `config_sheet.py`).
- **Impact:** 1 LOC; removes a stringly-typed outlier.

### A4. Remove task-narrative / Phase-N comments
- **Files:** grep for comments mentioning "T7", "T13", "Task N", "Stage 1",
  "Phase 1" across `budget_planner/`. CLAUDE.md forbids these: comments explain
  *why*, not *when added*. `layout.py:64–65` calls out "Phase 1 readers" which
  no longer exist.
- **Fix:** Delete task-reference comments. If the underlying *why* is still
  load-bearing, rewrite as a content-focused comment.
- **Impact:** ~20 LOC.

### A5. Collapse `SUMMARY_COLS_PER_YEAR` / `TOTAL_COLS` redundancy
- **File:** `summary.py:75–80` — two constants with the same value 5 used for
  different meanings.
- **Fix:** One `SUMMARY_COLS_PER_SECTION = 5` constant (or move to `layout.py`
  near the Summary-sheet layout).
- **Impact:** 4 LOC; removes drift risk.

---

## Tier B — Medium-leverage consolidation (~3–4 hours, ~100–150 LOC reduction)

Worth doing, but warrants its own review commit. Run full suite + pyright after
each.

### B1. Extract `month_range_ref(row, start_col, end_col)` helper
- 8 near-identical f-string builders in `plan.py`, `actuals.py`, `capacity.py`,
  `actuals_hours.py`. Move to `layout.py`.
- **Impact:** ~20 LOC, removes a stringly-typed idiom.

### B2. Extract section-reader skeleton in `reader.py`
- `read_team_from_config`, `read_assignments_from_config`,
  `read_externals_from_config` share structure: find-header → named-range or
  bounds → sentinel loop with `_parse_config_float` guard.
- **Risk note:** The openpyxl rule about mixed-sentinel break semantics (bool
  subclass trap, formula strings, empty-string vs None) is load-bearing. The
  extracted helper MUST preserve `_parse_config_float`'s contract on all
  callers.
- **Impact:** ~80–100 LOC consolidated; improved testability.

### B3. `PersonGridContext` dataclass for 14-param builders
- `_write_person_grid` (capacity) and `_write_person_hours_grid`
  (actuals_hours) each take 13–14 kwargs with identical shapes.
- **Fix:** Introduce a `PersonGridContext` dataclass in `person_section.py`;
  orchestrator builds it once per person; both builders consume it.
- **Risk:** Mechanical but wide — touches both builders + any direct test
  callers. Grep before acting.
- **Impact:** ~30 LOC signature reduction, big readability win.

### B4. Year-column-index helper in `summary.py`
- 7 loops at `summary.py:159–192` build parallel `plan_year_cols`,
  `plan_eur_cols`, `act_eur_cols`, `fcst_eur_cols` with offsets 0/1/2/3.
- **Fix:** One helper `_year_col_indices(offset, num_years, start, stride)`;
  4 one-line calls.
- **Impact:** ~20 LOC reduction; reduces offset-typo risk.

### B5. Remove `col_letter()` thin wrapper (optional)
- `styles.py:70–72` is literally `return get_column_letter(c)`.
- **Fix:** `from openpyxl.utils import get_column_letter as col_letter` —
  or a scripted rename across ~80 sites.
- **Verdict:** Defer. The aliasing is cheap, wide-reach rename is noisy, and
  the wrapper actually does provide a useful single import point if we ever
  want to validate bounds.

---

## Tier C — Larger restructure (defer; needs discussion first)

### C1. Split `reader.py` (1263 LOC) into `reader/` submodules
- Suggested split: `reader/config.py`, `reader/capacity.py`,
  `reader/actuals.py`, `reader/_parse.py` (shared helpers), `reader/__init__.py`
  for the public API.
- **Why not yet:** touches import paths across ~10 test files. No functional
  win, only ergonomic. Do this when we next need to add a reader.

### C2. Conditional-formatting helper family
- Per-row CF rules in capacity/actuals are correct (per the
  openpyxl-patterns rule about per-row rules referencing per-WP end months) —
  **do not batch them into range-wide rules.** The per-row design is
  load-bearing. A helper `add_row_cf(ws, row, col, formula, fill)` is safe;
  batching is not.

### C3. Builder-orchestrator strict mode / fallback audit
- **Explicit don't-do note:** The 2026-04-13 simplification sweep already
  tried to tighten `build_config_sheet` fallbacks and broke ~15 test
  fixtures. Do not repeat without grepping `tests/` for bare builder callers
  first.

---

## Explicit rejections (flagged but I disagree)

1. **Narrowing `except Exception` in `reader.py:188–191`** — this wraps
   `list(dn.destinations)` on third-party openpyxl `DefinedName` whose
   internals change across versions. Broad catch + `None` return is the right
   defensive posture here.
2. **`_strict=True` kwarg on builders** — adds complexity for a hypothetical
   test-authoring bug. The existing orchestrator-covers-it pattern is fine.
3. **Session-scoped session-wide workbook copy fixture** — claimed ~22s suite
   saving is speculative; measure first. 82s / 372 tests = 220ms/test, already
   fast enough.
4. **Merged-cell header helper** — the existing `write_section_header` in
   `styles.py:75` already does this; most remaining copy-paste is at the
   sub-header level which varies meaningfully between builders.

---

## Recommendation

Land **Tier A** as a single commit (`refactor(simplify): Tier A quick wins`)
after running full `pytest` and `pyright`. Defer Tier B to follow-up commits
each with their own test + review pass. Leave Tier C unless a separate need
drives it.

Estimated total safe reduction: ~200 LOC across A + B, zero functional change.
