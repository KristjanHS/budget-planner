# Test Suite Audit — Efficiency & Value Cleanup

**Created:** 2026-04-26
**Origin:** `testing-anti-patterns` skill audit of all 56 test files / 609 tests / 15.2k lines.
**Baseline metrics at audit time:** prod 11,371 lines, tests 15,212 lines (1.34:1 ratio, healthy bulk). Suite runtime ~158s. Total: 609 tests collected.

## Why this exists

A thorough sweep applied the 5-check test value gate (`~/.claude/rules/testing.md`), the 6 anti-patterns (`testing-anti-patterns` skill), and project rules (`.claude/rules/testing-project.md`) to the suite. Findings landed in 7 phases ordered by ROI. This plan exists so the next session can execute incrementally without re-doing the audit.

**Estimated total impact:** ~20 tests removed/consolidated, ~500 lines deleted, 3–5% runtime drop, hardcoded-date silent-rot bug removed.

## Execution rules

1. **Each phase is independently shippable.** Run `pytest tests/ -q` after each phase; commit per phase.
2. **Verify before deleting.** Audit findings are recommendations; the executing session must read each flagged test and confirm the gate failure before removal. Don't trust the audit blindly.
3. **No bulk commits.** One commit per phase keeps git log useful for rollback.
4. **Per CLAUDE.md Critical Rule 2:** after any code change, `python scripts/recalc.py output/budget_plan.xlsx` then `pytest tests/ -v` (or targeted subset) before claiming done.
5. **Per `plan-hygiene.md`:** when fully shipped, `git rm` this plan (single-session shippable; commits + this section are the audit trail). Don't `git mv` to archived.

---

## Phase 1 — Delete low-value tests (5-check gate failures)

**Effort:** ~15 min. **Impact:** 5 tests removed, marginal runtime delta.

For each candidate: open the file, read the test, confirm the cited gate failure. If the test could plausibly catch a real-world edit error (esp. data-integrity regressions from constant edits), keep it.

| File | Test | Gate | Rationale to verify |
|------|------|------|---------------------|
| `tests/test_parse_config_float.py:85` | `test_unsupported_type_returns_none` | 5 (impossible input) | `_parse_config_float` is only called on openpyxl `Cell.value`, never custom objects/lists. |
| `tests/test_styles.py:16` | `test_style_cell_default_locked` | 1 (testing language) | Asserts openpyxl's default `cell.protection.locked`, not `style_cell` behavior. **Verify:** read `budget_planner/styles.py::style_cell` first — if it sets `locked` explicitly, this test is real; if not, delete. |
| `tests/test_team.py:3` | `test_team_person_names_are_strings` | 1 (data shape, not code) | `isinstance(m["name"], str)` on TEAM constants. |
| `tests/test_team.py:7` | `test_team_fte_positive` | 4 (no function under test) | TEAM is a module constant; this is data lint. **Caveat:** stakeholder-naming rule says individual names never in files — confirm TEAM names are role-based before deleting. |
| `tests/test_config.py:6` | `test_external_commitments_persons_in_team` | 3 (data lint) | **High risk to delete** — this catches a typo in `EXTERNAL_COMMITMENTS` referencing a person not in `TEAM`, which silently corrupts downstream rollup. **Recommend keep**; document the audit's recommendation but override. |

**Acceptance:** `pytest tests/ -q` green. `git diff --stat` shows ≤ 5 deleted test functions.

---

## Phase 2 — Consolidate reader test files

**Effort:** 1–2 hours. **Impact:** ~300 lines deleted, 3 files → 2, no coverage loss.

Three reader test files (`test_reader.py`, `test_reader_v2.py`, `test_reader_new.py`) accumulated incrementally during the v1→v2→new reader refactors. Scope overlaps but isn't redundant.

### Steps
1. **Read all three files end-to-end** before touching anything. Map each test to its scope (capacity, assignments, externals, team, config, actuals).
2. **Build a coverage matrix** (test name → what's actually under test). Identify true duplicates vs same-named-different-scope.
3. **Merge `test_reader_v2.py` + `test_reader_new.py` → `test_reader_capacity.py`** for capacity/assignments/externals/team coverage. Distinguish:
   - Synthetic `_make_config_wb()` tests for fresh-parse coverage
   - `assigned_wb_path` round-trip tests for builder→reader integration
4. **Extract Priority-field tests** (currently in `test_reader_v2.py`) into `test_reader_team_priority.py`.
5. **Delete duplicates flagged by audit** (verify each before removal):
   - `test_reader_new.py::test_read_config_sheet_roundtrip` — duplicate of `test_config_sheet.py::test_params_round_trip` (line 41). **Verify** the test_config_sheet version actually covers the same cells.
   - `test_reader_new.py::test_read_actuals_values_*` (5 tests) — synthetic actuals reading; integration coverage is stronger in `test_rebuild_e2e.py` + `test_actuals*.py`. **Verify** by grep before removing.
6. **Keep `test_reader.py`** unchanged (4 tests, legacy `GRANT_ACCEPT_CEILING_MIGRATION` env-guard — orthogonal scope).

**Project rule:** `testing-project.md` § "round-trip required" — when consolidating, prefer round-trip assertions over reader-only "does it parse?" tests.

**Acceptance:** `pytest tests/test_reader*.py tests/test_config_sheet.py -q` green. Sum of test counts before/after: no net coverage loss (count behaviors covered, not tests).

---

## Phase 3 — Delete redundant rebuild test file — **OVERRIDE: KEEP**

**Effort:** 10 min. **Impact:** 1 file / 2 tests removed, ~0.3s savings.

**2026-04-26 override during execution:** verified both audit claims are wrong; `tests/test_rebuild_followups.py` is NOT redundant. Kept the file.

- `test_rebuild_prunes_orphaned_hours_when_person_removed` covers `workbook.py` lines 781/789 (`Orphaned Hours: ...`) — the **Actuals-Hours** prune block, fired when a person is removed from the roster while having logged hours. `test_timelog_plan_roundtrip.py::test_orphan_wp_dropped_with_bystander` covers `workbook.py` lines 764/769 (`Orphaned Timelog Plan cell: ...`) — the **Timelog Plan grid** prune block, fired when a grid cell carries a now-missing WP code. Different `warnings.append` sites, different sheets, different stdout strings. Sibling code paths, not duplicates.
- `test_rebuild_falls_back_to_capacity_when_config_sections_absent` exercises blanking the Config `PERSON-WP ASSIGNMENTS` / `EXTERNAL COMMITMENTS` section headers so `_find_header_row(required=False)` returns None and the rebuild falls back to `read_capacity_sheet`. `test_rebuild_e2e.py` calls `read_capacity_sheet` directly in helpers but never blanks Config sections — the fallback branch is otherwise unexercised.

The audit conflated two `Orphaned ...` warnings as one and assumed direct `read_capacity_sheet` calls equaled fallback coverage; they don't.

**No deletion. Phase 3 closes as no-op with override recorded above.**

---

## Phase 4 — Fixture-duplication fixes

**Effort:** 30 min. **Impact:** 3–5s per full-suite run (~2–3% of 158s).

### 4a — `test_timelog_plan_roundtrip.py` module fixture
3 tests (lines 77, 122, 161) each call `init_workbook(..., DEMO_CAPACITY_ASSIGNMENTS)` + `subprocess recalc.py`. Per `~/.claude/rules/testing.md` § "Reducing test runtime" → migrate to module-scope.

```python
@pytest.fixture(scope="module")
def recalced_timelog_wb_path(tmp_path_factory):
    out = tmp_path_factory.mktemp("tl_rt") / "wb.xlsx"
    init_workbook(str(out), assignments=DEMO_CAPACITY_ASSIGNMENTS)
    subprocess.run(["python3", str(SCRIPTS_DIR / "recalc.py"), str(out)], check=True)
    return str(out)
```

**CRITICAL:** Per `testing.md`: "Session/module fixtures that mutate a file must return the mutated artifact, not just metadata." Return the path, not stats.

**CRITICAL:** Per `testing.md`: "When migrating a module fixture to a session fixture, verify via full-suite wall-clock delta (not `--durations` top-N)." Run `time pytest tests/ -q` before and after; the durations report attributes setup cost to the first triggering test, which can mislead.

If any of the 3 tests **mutate** the workbook (rebuild, write), they cannot share a module fixture — each needs isolation. Read the test bodies first.

### 4b — `test_rebuild_shapes.py` consume `fresh_wb_path`
Line 53 (`test_rebuild_roundtrip_shape_variant`) calls `init_workbook(str(path))` inline. The session fixture `fresh_wb_path` already exists in `tests/conftest.py:14`. Replace inline build with fixture consumption.

**Verify:** parameterized test mutates the workbook (rebuild + recalc). If yes, the test still needs a copy — use `shutil.copy(fresh_wb_path, tmp_path/...)` to copy the session-built file rather than re-building.

**Acceptance:** `time pytest tests/ -q` shows wall-clock improvement; targeted runs of the touched files green.

---

## Phase 5 — Refactor hardcoded dates

**Effort:** 30 min. **Impact:** No runtime delta; prevents 2027-and-later silent rot.

Per `~/.claude/rules/testing.md` § "Time-dependent expected values" — derive from the same formula as production code, never hardcode.

Six files embed `datetime(2026, 7, 1)` (or `2026, 1, 1`) in setup:

| File | Line | Replace with |
|------|------|--------------|
| `tests/test_reader.py` | 14, 42, 103 | `config.PROJECT_START` (or matching constant from `budget_planner/config.py`) |
| `tests/test_reader_v2.py` | 415 | same |
| `tests/test_trends.py` | 437 | same |
| `tests/test_rebuild.py` | 103 | same |
| `tests/test_config_sheet.py` | 41 | same |
| `tests/test_team.py` | 14 | same |

**Verify the constant exists** before sweeping: `grep -n "PROJECT_START\|DEFAULT_START" budget_planner/config.py`. If the project's seed date isn't a single constant, derive from `seed_demo_data` defaults or refactor `config.py` to expose one (this is a tiny but real refactor — note in commit message).

**Project rule reminder:** `.claude/rules/testing-project.md` says "Tests must not hardcode config.py seed values" — this includes dates. Use this rule as the commit-message anchor.

**Acceptance:** Run `git grep "datetime(2026"` in `tests/` after — should return no hits.

---

## Phase 6 — Anti-pattern cleanup

**Effort:** 1–2 hours. **Impact:** Maintainability; ~1 test removed.

### 6a — `tests/test_workbook.py:255::test_build_order_is_stable` (anti-pattern: testing mock behavior)

Test patches 9 builder functions with spies and asserts call order. Per `testing-anti-patterns` Anti-Pattern 1: this verifies that patching/spying works, not that the build order is meaningful.

**Replace with:** simple integration test — `init_workbook` then assert all expected sheets exist + cross-sheet formulas resolve (`Capacity!...`, `Plan!...`). DAG breaks fail at build time anyway; the spy chain is unmaintainable if builders ever refactor (e.g. parallelization).

**Caveat:** if the test is the only guard against accidentally re-ordering builders such that `Capacity` runs before `Plan`'s WP-list is ready, document that scenario in the replacement test's docstring. Don't drop coverage silently.

### 6b — `tests/test_config_sheet.py:734::test_resolve_version_stamp_fallback_on_git_failure` — KEEP

Audit flagged the monkeypatched `subprocess.run` as fragile, but this is testing a specific contract (fallback path exists when git fails). **Do not delete.** Optionally add an inline comment: `# Mock isolates the fallback branch; real git absence is caught by CI env setup.`

### 6c — `tests/test_reader.py::test_legacy_migration_sanity_guard_*` (lines 108, 122) — KEEP

Environment-variable monkeypatching is the right tool for env-gating tests. Audit flagged as boundary-fragile but no concrete fix proposed. **Leave as is.**

**Acceptance:** `pytest tests/test_workbook.py -v` green. The replacement integration test passes and (manually verified) catches the same class of regression the spy chain was guarding.

---

## Phase 7 — Reader-only → round-trip strengthening

**Effort:** 1 hour. **Impact:** Coverage strengthening (no test count delta).

`tests/test_reader_new.py:51::test_read_config_sheet_roundtrip` — currently reader-only (Config sheet → read → assert). Per `testing-project.md`: "A reader-only 'does it parse?' assertion is a false green — the builder may still clobber the preserved value."

**Decision tree:**
1. Read `tests/test_config_sheet.py:41::test_params_round_trip` — does it cover the same Config cells via the full `init → read → mutate → rebuild → read` loop?
   - **Yes:** delete `test_read_config_sheet_roundtrip` as duplicate (rolls into Phase 2 cleanup).
   - **No:** extend `test_read_config_sheet_roundtrip` to be a true round-trip (init → read → mutate via openpyxl → `rebuild_workbook` → reopen → assert).

**Acceptance:** Either the test is gone (and the deletion commit references `test_params_round_trip` as the surviving coverage), OR the test now mutates + rebuilds + reasserts.

---

## Closeout

After all 7 phases ship:

1. Confirm `pytest tests/ -q` green and `time` shows wall-clock improvement.
2. Update `pytest --collect-only | tail -1` count delta in commit message of final phase (target ~589 tests, down from 609).
3. `git rm docs/plans/2026-04-26-test-suite-audit.md` — single-session shippable plan; commits + closeout summary in the final commit message are the audit trail.

If a phase reveals the audit's recommendation was wrong (e.g. a "delete candidate" actually catches a real bug class), record the override inline in this plan **before archiving** so the audit's reasoning is corrected for future reference.

## Auditor's surprises (preserve for future test-discipline decisions)

These aren't action items but are worth carrying forward:

- **Zero external-process tests outside `tests/manual/`.** The `external-process-test-gate` hook is working — no `requests`, docker, kubectl, playwright sneak-ins. Keep enforcing.
- **Round-trip pattern usage is strong.** `init → mutate → rebuild → reread` is the dominant idiom. This is correct for openpyxl's clobber-on-write surface.
- **Session fixture chain is well-designed.** `demo_wb_path` → `demo_recalc` → `demo_recalc_path` correctly enforces recalc-before-read ordering via dependency edges. Don't disturb.
- **No data-structure mocking.** Reader/builder return values are constructed via real fixtures or synthetic Config workbooks, never mocked dicts. Maintain this.
- **Test naming is precise** ("returns_grade", "names_and_fte"). Preserve when consolidating in Phase 2 — don't generalize to "test_reader_returns_team".
