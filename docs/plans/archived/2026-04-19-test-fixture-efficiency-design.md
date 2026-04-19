# Test fixture efficiency

**Date:** 2026-04-19
**Goal:** Cut `pytest tests/` runtime by promoting shared setup into session scope and reusing existing conftest fixtures, instead of deleting tests. Scope narrowed per user direction: focus on fixture efficiency.
**Non-goal:** Deleting tests or introducing `@pytest.mark.slow` markers. (Marker splits hide cost without reducing it — see `~/.claude/rules/testing.md`.)

## Baseline (2026-04-19)

- **522 tests in 143.06s** (2m 23s)
- Prior audit baseline (2026-04-12): 307 tests in 63s
- Growth: +70% tests, +127% runtime — runtime outpaced test count

### Measured setup costs (one call, warm process)

| Operation | Cost |
|---|---|
| `init_workbook()` (no assignments) | 0.33s |
| `init_workbook(assignments=DEMO_CAPACITY_ASSIGNMENTS)` | 0.42s |
| `seed_demo_data()` | 0.32s |
| `scripts/recalc.py` via subprocess (LibreOffice) | ~7–12s |
| `openpyxl.load_workbook(data_only=False)` | ~0.05s |

### Durations top-5 (from `pytest --durations=30`)

| Rank | Time | Test / setup | Owner |
|---|---|---|---|
| 1 | 12.70s call | `test_hours_propagate_through_rebuild_and_recalc` | test_rebuild_hours_forecast_e2e.py (inline subprocess) |
| 2 | 12.29s setup | `test_demo_mode_recalc_zero_errors` | conftest.py `demo_recalc` (session, shared by 8+ tests) |
| 3 | 11.84s setup | `TestRebuildRecalc` | test_rebuild_e2e.py (class-scoped, shared by 2 tests) |
| 4 | 11.54s setup | `test_total_actual_sums_wp_actuals` | test_capacity_todate.py `cap_values` (**module-scoped, 1 test**) |
| 5–30 | <2s each | rebuild round-trips | various |

**Interpretation.** Four recalc subprocess calls account for ~48s (~34% of runtime). Of those:
- #2 is legitimately shared (8+ tests use the session `demo_recalc`). Good.
- #3 is class-shared (2 tests). Acceptable.
- #1 is used by one test but tests a distinct code path (edited grade rate → rebuild → recalc hours-EUR chain). Legitimately unshareable.
- **#4 is the redundancy target — redoes what session `demo_recalc` already did, for one test.**

## Waste map

### Waste #1: `test_capacity_todate.py::cap_values` duplicates session `demo_recalc`

`tests/test_capacity_todate.py:34-42`:

```python
@pytest.fixture(scope="module")
def cap_values(demo_wb):
    import subprocess
    subprocess.run(
        ["python", "scripts/recalc.py", demo_wb],
        check=True, capture_output=True,
    )
    wb = openpyxl.load_workbook(demo_wb, data_only=True)
    return wb["Capacity"]
```

And `demo_wb` on line 20:
```python
@pytest.fixture(scope="module")
def demo_wb(tmp_path_factory):
    out = tmp_path_factory.mktemp("td") / "test.xlsx"
    init_workbook(str(out), assignments=DEMO_CAPACITY_ASSIGNMENTS)
    seed_demo_data(str(out))
    return str(out)
```

This rebuilds the same demo workbook `conftest.py::demo_wb_path` already builds (session-scoped), then recalcs it — which `conftest.py::demo_recalc` already did on the same file.

**Cost: ~11.5s per suite run. Entirely redundant.**

**Catch:** the existing `demo_recalc` fixture returns the recalc JSON, not a path/workbook. After recalc runs, `demo_wb_path` is a mutated file with cached values. Adding a tiny session fixture that returns that path (after `demo_recalc` has forced the recalc) lets `cap_values` and future callers load it with `data_only=True` for free.

### Waste #2: Module-scoped fixtures re-run `init_workbook` / `seed_demo_data` that session fixtures already produced

Every sheet-specific test file (~22 files) has its own module-scoped `init_workbook(...)` + optional `seed_demo_data(...)` fixture. Examples (confirmed by reading):

| File | Module fixture | Dup of session fixture |
|---|---|---|
| `test_capacity_todate.py` | `demo_wb` = init + demo-assign + seed | `demo_wb_path` |
| `test_trends.py` | `trends_workbook` = init + seed | `demo_wb_path` (load) |
| `test_actuals_hours.py` | `wb_path` = init + demo-assign | `assigned_wb_path` |
| `test_plan_forecast.py` | `assigned_wb` = init + demo-assign | `assigned_wb_path` (load) |
| `test_plan_forecast.py` | `empty_wb` = init only | `fresh_wb_path` (load) |
| `test_plan_v2.py` | workbook = init + demo-assign | `assigned_wb_path` (load) |
| ...plus `test_actuals.py`, `test_summary.py`, `test_grouping.py`, `test_merged_cells.py`, `test_conditional_formatting.py`, `test_timelog_plan.py`, `test_gapped_wp.py`, `test_summary_partial_year.py`, `test_actuals_breach_todate.py`, `test_actuals_forecast_block.py`, `test_actuals_hours_forecast_shading.py`, `test_actuals_hours_todate_rows.py`, `test_trends_reconciliation.py`, `test_capacity.py`, `test_capacity_actual_view.py`, `test_team.py`, `test_workbook.py` | | |

**Conservative estimate.** ~20 redundant `init_workbook` calls and ~10 redundant `seed_demo_data` calls per suite run:
- 20 × 0.38s (mix of empty + assigned) = **~7.6s**
- 10 × 0.32s (seed) = **~3.2s**
- **Total: ~10–11s**

These fixtures predate the conftest session fixtures added in the 2026-04-12 audit. They were never migrated.

**Caveat.** Tests that MUTATE the workbook (rebuild, write, then re-read) must NOT share a session path — mutation poisons it for every other test. The audit in step 2.1 below must distinguish read-only from mutating fixtures. Rough scan: none of the listed module fixtures above mutate after init; they load read-only and inspect. Confirm per-file during execution.

### Waste #3: `test_rebuild_hours_forecast_e2e.py::test_hours_propagate_through_rebuild_and_recalc` inline recalc

One 12.70s test runs its own `subprocess.run(recalc.py)` inline, not in a fixture. Today it's a single-assertion test so sharing buys nothing. **No change proposed**, but document the "one test, unshareable" status so a future contributor who adds a second hours-chain assertion promotes it to class-scope instead of copy-pasting the recalc.

## Changes

### Change 1: Add a session-scoped `demo_recalc_path` to conftest

`tests/conftest.py` gets one new fixture after `demo_recalc`:

```python
@pytest.fixture(scope="session")
def demo_recalc_path(demo_wb_path, demo_recalc):
    """Path to the demo workbook AFTER session recalc ran.

    Depends on demo_recalc to enforce ordering: the recalc subprocess
    must have mutated demo_wb_path before callers open it with
    data_only=True. Returns the same path as demo_wb_path, just with
    a dependency edge that guarantees recalc has run.
    """
    return demo_wb_path
```

The `demo_recalc` dep isn't consumed but forces execution order. (pytest runs dependency fixtures even when their return value is unused.)

### Change 2: Migrate `test_capacity_todate.py::cap_values` to session recalc path

Replace the module-scoped `demo_wb` + `cap_values` fixtures with:

```python
@pytest.fixture(scope="module")
def cap_formulas(demo_wb_path):
    wb = openpyxl.load_workbook(demo_wb_path, data_only=False)
    return wb["Capacity"]

@pytest.fixture(scope="module")
def cap_values(demo_recalc_path):
    wb = openpyxl.load_workbook(demo_recalc_path, data_only=True)
    return wb["Capacity"]
```

**Savings: ~11.5s** (eliminated duplicate recalc + init + seed).

### Change 3: Audit & migrate module-scoped init fixtures to session paths

Per-file mechanical pass:

1. For each file in Waste #2 table, grep for fixture body patterns:
   - `init_workbook(...)` → if args match a session fixture, replace fixture body with `return <session_path>` or `return load_workbook(<session_path>, data_only=...)`.
   - `init_workbook(..., assignments=DEMO_CAPACITY_ASSIGNMENTS)` → use `assigned_wb_path`.
   - `init_workbook(...); seed_demo_data(...)` → use `demo_wb_path`.
   - `init_workbook(...)` (no args) → use `fresh_wb_path`.
2. Verify no test in the file writes/saves/rebuilds the workbook after the fixture. If any does, leave that file's fixture alone (or split into a mutating-path-aware fixture).
3. Run that file's tests in isolation (`pytest tests/<file>.py -q`) after migration to catch cross-test leakage.

**Savings: ~10–11s.**

### Change 4: `cap_formulas` needs to outlive rebuild tests on demo_wb (verification)

Before migration, grep `demo_wb_path` usage across the suite for any writer that could mutate it session-wide. If any test saves over this path, Change 1 needs a separate "read-only" session path. (Current conftest docstring already warns: `# Read-only — tests that consume this fixture MUST NOT save or rebuild.`)

## Expected outcome

| Metric | Before | After | Δ |
|---|---|---|---|
| Runtime | 143s | ~120–122s | **−20–23s (~15%)** |
| Test count | 522 | 522 | 0 |
| Coverage | — | unchanged | — |
| Fixture cost concentration | 22 files × module init | 22 files × session path load | — |

No tests deleted. All savings from fixture scope + dedup.

## Execution order

1. **Change 1** (conftest `demo_recalc_path`) — 1 file, ~5 lines. Run `pytest tests/conftest.py --collect-only` to sanity-check.
2. **Change 2** (test_capacity_todate.py migration) — isolated, easy to verify. Run `pytest tests/test_capacity_todate.py -q --durations=5`. Expected: setup drops from 11.5s to <0.1s.
3. **Change 3** (mass migration) — one commit per file OR one commit batched by cluster (actuals files, capacity files, trends files). After each file or cluster, run `pytest tests/<file>.py -q` in isolation. After the full pass, run `pytest tests/ -q --durations=25`.
4. Final validation: full suite, confirm green + record new duration baseline in this doc under "Outcome".

## Not doing (explicitly out of scope for this plan)

- **Deleting low-value tests.** User narrowed scope to fixture efficiency. If the fixture pass yields only ~15% and CI still feels slow, a follow-up applying the 5-check gate (testing.md) to the 522-test count is warranted — top candidates for audit: the 7 `test_rebuild*.py` files (possible cross-file duplication), `test_reader_new.py` (824 lines) vs `test_reader_v2.py` (559 lines) versioned-name pair.
- **Promoting `test_rebuild_hours_forecast_e2e.py` recalc to a fixture.** Single consumer; premature until a second test needs it.
- **Parallelization (`pytest-xdist`).** Would mask fixture waste instead of fixing it; session fixtures don't survive worker boundaries cleanly for this workbook pattern.

## Validation checklist

- [ ] Full suite green after each change
- [ ] `pytest --durations=25` shows `test_capacity_todate` setup <0.1s after Change 2
- [ ] No module-scoped fixture in migrated files calls `init_workbook` or `seed_demo_data`
- [ ] `grep -rn "init_workbook" tests/` count drops by ~20
- [ ] Final runtime recorded below

## Outcome

Shipped across two stages (2026-04-19):

- **Stage 1** (Changes 1+2 — conftest `demo_recalc_path` + `test_capacity_todate.py` migration): 143s → 130.80s (−12.26s). Commits `9dc3686`, `39decca`.
- **Stage 2** (Change 3 — mass migration, 4 clusters, 19 files): 130.80s → **123.67s** (−7.13s).
  - Cluster A `fresh_wb_path` (4 files): `8717701`
  - Cluster B `assigned_wb_path` (11 files): `383e394`
  - Cluster C `demo_wb_path` (3 files): `cd9a206`
  - Cluster D `test_workbook` class fixtures: `e4d15b7`

**Totals**: 143s → 123.67s (**−19.33s / ~13.5%**). Expected range was 120–122s (−20–23s); landed just above.

### Durations top-5 after Stage 2

| Rank | Time | Test |
|---|---|---|
| 1 | 12.77s call | `test_hours_propagate_through_rebuild_and_recalc` (unshareable, documented) |
| 2 | 11.90s setup | `TestRebuildRecalc::test_recalc_zero_errors_after_fte_change` (class-shared, OK) |
| 3 | 11.45s setup | `test_capacity_todate::test_total_actual_sums_wp_actuals` — session `demo_recalc` attribution (alphabetical first-consumer artifact; real savings materialized in wall-clock per testing.md) |
| 4 | 1.99s call | `test_hours_actuals_through_month_round_trip` |
| 5 | 1.64s call | `test_override_orphaned_by_removed_wp_is_dropped` |

Three recalc subprocesses (one session demo recalc, one class recalc, one inline e2e) remain the irreducible floor at ~36s. Everything else is sub-2s.

### Validation checklist — all green

- [x] Full suite green after each change (522 passed at each commit)
- [x] `test_capacity_todate` setup dropped to demo_recalc_path session dep (shared artifact, not module-duplicated)
- [x] No module-scoped fixture in migrated files calls `init_workbook` or `seed_demo_data`
- [x] 18 test files no longer call `init_workbook` at all (all Waste #2 targets except test_gapped_wp/test_team/test_workbook module-root which were already delegated)
- [x] Final runtime 123.67s recorded above
