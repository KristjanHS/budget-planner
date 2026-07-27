# Capacity Row Cleanup + Skip-and-Cumulate — Plan

**Status**: PLAN 2026-04-28 (resolved from brainstorm; ready for `/impag` Stage 1 in a fresh session).

Resolved from the prior brainstorm (renamed in place after user picks). The brainstorm's framing — Items 1–4, what's-there-today / what-changes / cross-sheet-coupling sections — remains the design rationale; this header captures the user's resolution and the implementation phases.

---

## Resolved decisions (user picks, 2026-04-28)

| ID | Decision |
|---|---|
| 1.1 | EU TOTAL: drop on Capacity only. Actuals-Hours keeps EU TOTAL + per-calendar-year ceiling CF (the enforcement surface). |
| 1.2 | REMAINING row: drop. Fold the per-person over-allocation CF onto the TOTAL row (red when month-cell on TOTAL exceeds Max FTE). |
| 1.3 | Capacity EU-ceiling preview loss: accepted. Reviewers consult Actuals-Hours for per-year ceiling pressure. |
| 2.1 | Zero-row definition: **(a) only** — `total_fte == 0` in Config PERSON-WP ASSIGNMENTS. The empty-window-intersection case (b) is **NOT in scope**. |
| 2.2 | Drop the row everywhere on rebuild even if preserved hours exist (A1). Data loss is silent and accepted; document in test docstring. |
| 2.3 | External rows: leave alone. |
| 2.4 | Empty-person-block: keep the 1-row spacer + header (preserves Plan/Trends section anchor). |
| 3.1 | Actuals (EUR) skip-and-cumulate: documented (D1) — formula already supports it; one-paragraph spec.md addition mirroring Hours wording. |
| Bonus | Hours calendar-year-boundary caveat on EU ceiling CF: documented in spec.md. |

---

## Stage split

8 implementation tasks → two stages (per `/impag` ≥5-task rule). Stage boundary lands after Stage 1's commit set: Capacity geometry shrinks by 2 rows per person, all tests green, no half-wired interfaces. Stage 2 then layers the zero-FTE filter and spec edits on top.

- **Stage 1 — Item 1** (Capacity row cleanup): drop REMAINING + EU TOTAL on Capacity, fold REMAINING CF onto TOTAL.
- **Stage 2 — Item 2 + spec edits**: filter `total_fte == 0` rows on Capacity + Actuals-Hours; spec.md doc edits for Items 3+4.

---

## Stage 1 tasks

### S1.T1 — capacity.py: drop REMAINING + EU TOTAL helpers, fold CF onto TOTAL
Files: `budget_planner/capacity.py`, `budget_planner/person_section.py`.

- Delete helpers (live in `capacity.py` post the 17-helper decomposition shipped in `019dfd8`):
  - `_write_remaining_row`, `_write_eu_total_row`
  - `_write_actual_remaining_band`, `_write_actual_eu_total_band`
  - `_apply_planned_total_remaining_cf`, `_apply_planned_eu_total_cf`
- Keep `include_remaining` / `include_eu_total` params on `_resolve_grid_state` and `person_section.plan_person_layout` — Actuals-Hours still consumes EU TOTAL. Capacity callers set both to `False`.
- New helper `_apply_planned_total_overalloc_cf` — red fill on TOTAL row when month cell `> Max FTE`. Replaces the dropped REMAINING-row CF.
- `_write_person_grid` orchestrator: remove dispatch to deleted helpers; wire the new TOTAL CF.
- `person_section.plan_person_layout`: when `include_remaining=False, include_eu_total=False`, geometry shrinks by 2 rows per person (subsequent persons shift up by `2 × person_index`).

### S1.T2 — Tests: prune REMAINING / EU-TOTAL assertions on Capacity
Files: `tests/test_capacity_todate.py` plus `grep -rn 'REMAINING\|EU TOTAL' tests/` sweep.

- Delete tests asserting REMAINING / EU TOTAL row presence on **Capacity** (Actuals-Hours tests stay).
- Update row-position assertions affected by the 2-row geometry shrink.
- Add test for the new TOTAL-row over-allocation CF (red when monthly TOTAL exceeds Max FTE).

### S1.T3 — Module docstring + spec.md (Item 1 only)
Files: `budget_planner/capacity.py` (top docstring), `docs/spec.md`.

- Capacity module docstring: drop REMAINING/EU TOTAL rows from the per-person geometry block.
- `docs/spec.md`: per-person Capacity geometry section + row-order contracts. Note that EU TOTAL persists on Actuals-Hours (enforcement surface).

### S1.T4 — Stage 1 closeout
- `python -m budget_planner --output output/budget_plan_demo.xlsx` (regenerate demo)
- `python scripts/recalc.py output/budget_plan_demo.xlsx`
- `pytest tests/ -q` (full suite — workbook output **changes by design**; structural-diff baseline must be regenerated)
- `pyright budget_planner/` and `ruff check budget_planner/ tests/`
- Commit per task (`/impag` step 4 commits one task at a time).

---

## Stage 2 tasks (queued for next session)

### S2.T1 — `_resolve_grid_state`: filter `total_fte == 0` WPs
File: `budget_planner/capacity.py`.

- New helper `_filter_zero_fte_wps(person_assignments, ordered_wp_ids) -> list[str]` removes WPs where the assignment's `total_fte` is 0.
- Apply inside `_resolve_grid_state` after the existing intersection filter.
- Empty-block edge case (every WP suppressed, no externals): keep the 1-row spacer + header per Q2.4=A.

### S2.T2 — Mirror the filter on Actuals-Hours
File: `budget_planner/actuals_hours.py`.

- Same suppression rule (no preserved-hours carve-out per Q2.2=A1).
- Factor the filter helper to a shared location (`budget_planner/person_section.py` is the natural home — both Capacity and Actuals-Hours already consume it).
- Confirm row-order mirror with Capacity (load-bearing per `docs/spec.md` § Row-order contracts).

### S2.T3 — Tests: zero-FTE suppression + rebuild contract
Files: `tests/test_capacity_todate.py`, `tests/test_actuals_hours.py`, `tests/test_rebuild_e2e.py`.

- Fixture: assignment with `total_fte == 0`. Assert WP row absent on **both** Capacity and Actuals-Hours.
- Rebuild test: previously-logged hours on a row that becomes `total_fte == 0` → row dropped, hours discarded silently. Test docstring explicitly documents this as **accepted A1 data loss** (per Q2.2 pick) so the failure mode isn't mistaken for a regression.
- Row-order mirror invariant: assert Capacity and Actuals-Hours have identical WP-row counts and identical WP IDs in identical order, post-filter.

### S2.T4 — `docs/spec.md`: Items 3+4 + Item 2 row-order amendment
File: `docs/spec.md`.

- Two-cadence actuals § EUR cadence: add skip-and-cumulate paragraph mirroring Hours wording (D1).
- Two-cadence actuals § Hours cadence: add calendar-year-boundary caveat — "if a skip crosses a calendar-year boundary, the dumped cumulative lands in the later year's annual SUM; the per-year EU ceiling CF on Actuals-Hours mis-fires on both years. Cumulative figure remains correct."
- Row-order contracts § "intersection filter": amend to "intersection filter, further reduced to WPs with non-zero `total_fte`".
- `capacity.py` + `actuals_hours.py` module docstrings: update if they cite the row-set rule.
- Watch: docs-only task BUT `tests/test_config_sheet.py::test_config_named_range_targets_match_spec_doc` parses spec.md tables. Run `pytest tests/ -q` before the closeout commit even though no Python changes (per `plan-hygiene.md` carve-out).

### S2.T5 — Stage 2 closeout
Same gates as S1.T4. Demo regen + full pytest + pyright + ruff. Commit per task.

---

## Risks / caveats

- **A1 silent data loss** (Q2.2): if a user previously logged hours on a WP row, then changes that assignment's `total_fte` to 0 and rebuilds, hours are silently discarded. No telemetry, no comment cell, no warning. `git log output/budget_plan*.xlsx` is the audit trail. Test docstring at `tests/test_rebuild_e2e.py` calls this out so future agents don't file the behavior as a bug.
- **(a)-only suppression** (Q2.1): empty-window intersection rows (assignment exists with non-zero `total_fte` but the WP active range and the assignment window have no overlap) still render as all-grey rows. This was the visually noisy case the brainstorm flagged; **NOT addressed by this plan**. Revisit if user feedback prompts a follow-up.
- **Capacity EU-ceiling preview gone** (Q1.3): only Actuals-Hours surfaces the per-year ceiling CF — and only after hours are logged. Auditors / reviewers must consult Actuals-Hours (not Capacity) for ceiling pressure. The plan-side preview is gone by design.
- **Workbook output changes by design**: `tests/test_xlsx_structural_diff.py` baseline will need regeneration after Stage 1 ships. This is **not** a byte-identical refactor (unlike Phase 3 of the architectural-audit plan).
- **`_write_person_grid` regression risk**: capacity.py recently shipped the 17-helper decomposition (commit `019dfd8`). Stage 1 deletes 6 of those helpers; verify the deletion order against the orchestrator dispatch chain so no dangling caller references survive (pyright + targeted pytest catch this).

---

## Test gates

- `pytest tests/ -q` green at each stage close.
- `pyright budget_planner/` clean.
- `ruff check budget_planner/ tests/` clean.
- `tests/test_rebuild_e2e.py` stays green at every commit (rebuild loop is a load-bearing invariant).
- `tests/test_xlsx_structural_diff.py`: baseline regenerated post-Stage 1; gate stays in force from Stage 2 onward.

---

## Out of scope / deferred

- Empty-window intersection suppression (Q2.1 option (b)). Revisit when a user reports the visual noise.
- Per-year EU ceiling CF migration to Trends or Summary (the Q1.3 = B / C alternatives the user declined). Revisit only if Capacity-preview loss proves painful in practice.
- Plan-hygiene archival of this plan: `git mv` to `docs/plans/archive/` after Stage 2 ships and commit references are in place.
