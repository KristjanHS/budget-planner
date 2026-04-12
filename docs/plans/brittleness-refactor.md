# Brittleness Refactor Plan

**Date:** 2026-04-12
**Scope:** `budget_planner/` — risk-ranked hardening plan targeting parts most likely to break under realistic user edits.
**Status:** Plan only. No code changes in this document.
**Inputs:** Architecture audit (Python Backend Architecture Review skill + targeted verification of top findings against the current source).

---

## 1. Executive summary

The workbook is well-structured around a single layout module and a stateless-builder contract. The brittleness lives at the **seams**, not inside any one module:

1. **Reader → workbook round-trip** is the riskiest surface. `reader.py` uses substring-based header discovery plus fixed `+2` row offsets to locate each section, and silently coerces formula strings to numeric defaults. A user who renames a section header, inserts a blank row, or writes a custom rate formula will see data silently disappear on rebuild.
2. **Cross-sheet formula coupling** leaks layout knowledge into builders. A handful of builders embed hardcoded Config column letters (`$B$`, `E{n}`) and row offsets instead of deriving from `layout.py`. When Config shape shifts, those formulas point at the wrong cells without raising.
3. **Implicit build order** in `workbook.py` is load-bearing but undocumented. Plan must run after Config and Capacity; Actuals must run after Trends-stub. Reordering or parallelising builders crashes the build.
4. **_META drift** is a known hazard (Critical Rule 6) but `meta.py` still lazy-imports from `config.py` as a fallback and hardcodes several row predictions that are only true for the default layout.

The plan below is ordered by **risk-reduction ÷ effort**. Tasks 1–3 close the highest-probability data-loss paths; tasks 4–6 remove silent-drift vectors; tasks 7–8 are hygiene. Nothing here changes the user-visible workbook.

---

## 2. Risk matrix (top findings)

| # | Finding | Severity | File:line | Failure mode | Task |
|---|---------|----------|-----------|--------------|------|
| F1 | `_safe_float` silently drops formula rates to default (6500/7500) | **HIGH** | `reader.py:19–26`, `86`, `120` | User's custom per-WP rate formula vanishes on rebuild; Ceiling is wrong | T1 |
| F2 | Hardcoded `section_row + 2` offset in every section reader | **HIGH** | `reader.py:107, 128, 194, 312, 355, 395, 432` | Extra blank row after header → first data row read as header; last row dropped | T2 |
| F3 | Substring header match can collide with user-renamed sections | **HIGH** | `reader.py:62` (used ×7) | Renaming "WP BUDGET" → "Legacy WP Budget" still matches | T2 |
| F4 | `cfg_wp_first` fallback to literal row 8 if regex parse fails | **HIGH** | `reader.py:721, 723` | Legacy workbooks with Config at different rows → rebuild writes to wrong rows | T3 |
| F5 | `current_month` named range hardcoded to `Trends!$B$1` | **HIGH** | `workbook.py:172–182` | If Trends layout shifts, every Actuals % Ceiling formula silently zeroes | T4 |
| F6 | Build order in `_build_all_sheets` is load-bearing, undocumented | **HIGH** | `workbook.py:118–223` | Reordering builders → KeyError on `capacity_layout`; silent skip on move_sheet | T5 |
| F7 | `meta.py` lazy-imports `config.py` as fallback when params are None | **HIGH** | `meta.py:39–60, 90–100` | _META documents seed data, not actual workbook state | T6 |
| F8 | Hardcoded `Config!E{row}`, `Config!$B${row}` in builders | **MED** | `actuals.py:112`, `capacity.py:340, 345` | Reordering Config columns → formulas fetch wrong values | T7 |
| F9 | `meta.py` row predictions hardcoded (`row 9`, `+3` offsets) | **MED** | `meta.py:96–100, 110–113` | Config or Summary layout change → _META points at wrong rows | T7 |
| F10 | Preserved Actuals dropped when all monthly values are None | **MED** | `reader.py:779–785` | Intentional zero-spend WP looks like "never entered" on rebuild | T8 |

---

## 3. Tasks (execute in order; each is independently shippable)

### T1 — Preserve formula rates across round-trip

**Problem:** `_safe_float` (reader.py:19–26) returns a numeric default whenever `float(value)` fails. Formula strings (`"=SUMPRODUCT(...)"`) fail conversion → rate silently becomes 6500 or 7500. The workbook loses the formula.

**Change:**
- Extend `work_packages` dict to carry `rate_raw` alongside `rate`: if the Config cell holds a formula string, keep it verbatim; if numeric, store as float.
- Update builders that write WP rate back to Config (`config_sheet.py`) to emit the formula string verbatim when present, else the numeric value.
- Apply the same treatment to `FTE Rate` in project_params (`reader.py:86` already attempts this — clean up the triple-nested ternary into an explicit branch).

**Files:** `reader.py`, `config_sheet.py`, `workbook.py` (pass `rate_raw` through).

**Tests:**
- New: `tests/test_reader_v2.py::test_formula_rate_roundtrip` — seed a workbook with `=AVERAGE(B9:B11)` in the FTE Rate cell, rebuild, assert cell still contains the formula string.
- New: per-WP formula rate round-trip.
- Regression: all existing reader tests must still pass.

**Effort:** M (half-day). **Rollback:** revert three files; existing tests already cover numeric path.

---

### T2 — Replace substring header discovery with anchored-label contract

**Problem:** `_find_header_row` (reader.py:43–68) scans column A for a case-insensitive *substring* match. Seven call sites rely on it. Combined with the `+2` offset (F2), any layout drift silently misaligns the read.

**Change:**
- Introduce `SECTION_HEADER_LABELS` in `layout.py` — a single dict mapping canonical section keys (`"wp_budget"`, `"wp_schedule"`, `"assignments"`, `"externals"`, `"roster"`, `"grade_rates"`, `"reporting_periods"`, `"actuals_wp_id"`) to their exact header strings.
- Change `_find_header_row` to require **exact match** (after strip + upper-case normalization), not substring. Raise a clear error identifying the expected vs. found label when not found.
- Replace each `row + 2` with a named helper `_data_row_after(header_row, header_rows=2)` — still magic-numbered, but localised. The two-header-row assumption becomes a single constant in `layout.py` (`SECTION_HEADER_ROWS = 2`).
- When a section is missing, builders must opt into `required=False` explicitly.

**Files:** `layout.py`, `reader.py`; no builder changes required if headers already match the canonical labels (verify this in T2a below).

**Sub-task T2a:** grep-audit every builder that writes a section header label; confirm they all match the new canonical dict. If any builder writes "Work Package Budget" instead of "WP BUDGET", fix the builder — do NOT loosen the match.

**Tests:**
- New: `test_reader_v2::test_exact_header_match_required` — rename "WP BUDGET" to "Legacy WP Budget" in a fixture workbook; assert `ValueError` with a useful message.
- New: `test_reader_v2::test_blank_row_between_header_and_data_detected` — insert an extra blank row; assert the reader detects the layout drift instead of silently reading header as data.

**Effort:** M (half-day, driven mostly by T2a grep cascade). **Rollback:** revert `layout.py` + `reader.py`; restore substring match as a temporary escape hatch guarded by an env var if needed.

---

### T3 — Remove hardcoded row-8 fallback for Config WP rows

**Problem:** `reader.py:708–724` tries to parse the first Plan WP cell as `=Config!A{n}`. If the regex fails, it searches Config for "WP BUDGET" and falls back to `wp_section + 2`, and *if that also fails*, hardcodes `cfg_wp_first = 8`. The literal 8 is a silent time bomb for legacy workbooks generated before the current layout.

**Change:**
- Delete the `cfg_wp_first = 8` fallback entirely. If both the formula parse AND the Config section-header scan fail, raise `ValueError` with the diagnostic: "Cannot locate Config WP Budget section. Regenerate workbook with `python -m budget_planner init`."
- If T2 lands first, the section-header scan becomes reliable enough that this fallback is never hit in practice. Keep it as the only discovery path.

**Files:** `reader.py:708–724`.

**Tests:**
- New: `test_reader_v2::test_no_wp_budget_section_raises` — workbook missing the Config!WP Budget section raises with clear message.
- Regression: ensure the happy path (formula-based Plan!A cells) still resolves.

**Effort:** S (1–2h). **Rollback:** trivial.

---

### T4 — Make cross-sheet formula coupling layout-derived, not hardcoded

**Problem:** Three hotspots embed cell addresses that layout.py should own:
- `workbook.py:172–182` — `current_month` named range → `Trends!$B$1`
- `actuals.py:112` — `=Config!E{config_row}` for Ceiling
- `capacity.py:340, 345` — `Config!$B${roster_row}` for Max FTE

Each is correct today. Each breaks silently if layout.py changes the underlying column.

**Change:**
- Add a small helper `layout.cell_ref(sheet, row, col_const)` that returns `"Sheet!$L$row"` using `openpyxl.utils.get_column_letter(col_const)`.
- Replace all three hotspots with `cell_ref(SHEET_CONFIG, config_row, COL_WP_CEILING)` etc.
- Define a new named range `wp_ceiling_col_letter` is **not** needed — the helper covers this.

**Files:** `layout.py` (+helper), `workbook.py`, `actuals.py`, `capacity.py`.

**Tests:**
- New: `test_layout::test_cell_ref_helper` — unit test.
- Regression: full E2E suite must pass unchanged (generated formulas are bytewise identical today).

**Effort:** S (1–2h). **Rollback:** trivial, pure refactor.

---

### T5 — Document and enforce build order

**Problem:** `_build_all_sheets` (workbook.py:118–223) runs builders in a specific order. Plan depends on Capacity's layout dict; Actuals depends on Trends stub; _META depends on all prior layouts. None of this is documented beside the code.

**Change:**
- Add a module-level docstring to `workbook.py` listing the build stages as a dependency diagram (Mermaid-or-ASCII, five lines max).
- Add inline comments at the two real coupling points: (a) where Plan consumes `capacity_layout`, (b) where Actuals consumes `trends_layout`.
- Introduce `BuildContext` dataclass (or just a typed dict) that each builder returns. The orchestrator threads it through explicitly — currently the data flows via local variables in `_build_all_sheets`, which works but obscures the DAG.
- Add `tests/test_workbook::test_build_order_is_stable` — a smoke test that imports `_build_all_sheets` and asserts the call sequence via mock spies. This is a "change detector" — if anyone reorders, the test fails with a pointer to this plan.

**Files:** `workbook.py` (docstring + BuildContext), `tests/test_workbook.py`.

**Tests:**
- New: build order spy test.
- Regression: `test_integration.py` full E2E unchanged.

**Effort:** M (half-day). **Rollback:** revert BuildContext; docstring-only change is safe.

---

### T6 — Remove `meta.py` lazy-import fallback to `config.py`

**Problem:** `meta.py:39–60, 90–100` lazy-imports `config.py` when its parameters are `None`. This means _META can document **seed data** instead of the actual workbook state when a caller forgets to pass the right args. Combined with the Critical Rule 6 hazard (meta drift), this is a footgun: meta silently "succeeds" with wrong content.

**Change:**
- Change `build_meta_sheet(ws, layout, project_params, work_packages, team, external_commitments, ...)` so that every data parameter is **required** (no `None` default).
- Delete the lazy-import block entirely.
- Update the only two callers in `workbook.py` (init path and rebuild path) to always pass the live values — they already have them.

**Files:** `meta.py`, `workbook.py`.

**Tests:**
- Update: any test that calls `build_meta_sheet` with fewer args must be updated (likely 2–3 tests in `test_workbook.py`, `test_rebuild.py`).
- New: `test_workbook::test_meta_requires_all_params` — assert TypeError on missing arg.

**Effort:** S (1–2h). **Rollback:** trivial.

---

### T7 — Collapse remaining hardcoded column letters and row predictions

**Problem:** Follow-up to T4, covering residual offenders:
- `meta.py:96–100` — `first_grade_row = 9` and friends, used only when `config_layout` is None. After T6, `config_layout` is always passed → these fallbacks become unreachable.
- `meta.py:110–113` — `sum_first_wp = 5` assumes Summary starts WPs at row 5.
- `plan.py:109` — relies on `config_layout["param_rows"]` dict keys matching specific labels.
- `capacity.py:19–24` — `_wp_config_row` re-derives layout locally.

**Change:**
- After T6, delete all unreachable "config_layout is None" branches in `meta.py`.
- Move `sum_first_wp` and the Summary row layout into `summary.py`, return it as a `summary_layout` dict threaded through `BuildContext` (T5). `meta.py` reads the dict, never guesses.
- Replace `capacity.py:_wp_config_row` with a call to the shared layout dict.
- Add `param_rows` schema validation in `layout.py` (TypedDict) so typos surface at pyright time.

**Files:** `layout.py`, `meta.py`, `summary.py`, `plan.py`, `capacity.py`.

**Tests:**
- Regression: full suite.
- New: `test_meta::test_meta_uses_live_summary_layout` — modify Summary to start at row 6, rebuild, assert _META reflects row 6.

**Effort:** M (half-day). **Depends on T5 and T6.**

---

### T8 — Distinguish "zero spend" from "never entered" in preserved Actuals

**Problem:** `reader.py:779–785` drops a WP from `preserved_actuals` when every monthly value is None. This conflates "user has not touched this WP" with "user entered zero for every month". On rebuild, both cases produce blank rows, losing the user's explicit-zero intent.

**Change:**
- Preserve a WP row if **any** monthly cell has been touched. Since openpyxl surfaces empty cells as None regardless, we need a different signal: either (a) preserve all WP rows that appear in the Actuals sheet (simplest), or (b) record a `touched` boolean by scanning cell formatting / comment / a hidden marker cell. Option (a) is preferred unless memory is a concern (it isn't — projects have <50 WPs).
- Document in the reader docstring that the round-trip contract is "every Actuals WP row is preserved, blank cells included".

**Files:** `reader.py:779–785`.

**Tests:**
- New: `test_reader_v2::test_zero_spend_wp_preserved` — WP with all blank monthly cells still appears in `preserved_actuals` and round-trips.

**Effort:** S (1h). **Rollback:** trivial.

---

## 4. Out of scope (deliberate non-goals)

- **Full migration to named-range–only formulas.** The audit flagged hardcoded `Config!$B$` references; T4 addresses the worst offenders via a `cell_ref` helper, not via a wholesale named-range overhaul. Named ranges already carry their own brittleness (deletion on sheet rename). Not worth the churn.
- **Replacing openpyxl with a domain model + render step.** Would be cleaner architecturally. 2-week effort. Defer until we hit a second major layout change that the current builders can't absorb.
- **Dynamic layout discovery in the reader (e.g. reading the _META sheet).** Tempting but circular — _META is generated by the same build order the reader must trust. Make the reader strict instead.

---

## 5. Sequencing and risk

**Execution order:** T1 → T2 → T3 → T4 → T5 → T6 → T7 → T8.

- T1–T3 are **data-loss fixes**; land these first and regenerate any shared workbooks after merge.
- T4 is a **pure refactor** with byte-identical output; safe to land any time.
- T5–T7 are **coupled** (T7 depends on the BuildContext from T5 and the required-params contract from T6). Review as one feature branch or three sequential PRs.
- T8 is independent, can land in parallel with any other task.

**Total estimated effort:** ~2.5 days of focused work, 8 PRs at most.

**Biggest risk during rollout:** T2 will break any user workbook where someone renamed a section header manually. Mitigation: land T2 behind a `strict_headers` kwarg on `read_config_sheet()` defaulting to True, with a migration note in the release commit. After one release cycle, remove the kwarg.

---

## 6. Invariants preserved

All tasks preserve the project's Critical Rules:
- Excel formulas remain authoritative; no calculated values added in Python.
- Font-color convention untouched.
- Named ranges `project_start` and `fte_rate` untouched.
- _META is now **more** accurate, not less (T6 + T7).
- `pytest tests/ -v` must pass after each task.
