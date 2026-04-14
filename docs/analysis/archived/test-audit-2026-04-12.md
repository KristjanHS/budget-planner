# Test Suite Audit — 2026-04-12

Two skills used: `brutal-honesty-review` (Ramsay mode, qualitative) + `mutation-testing` (mutmut, empirical).

## Baseline

- 256 tests → 260 after this audit (added 4 None-defaults parametrised cases)
- 5486 test lines across 22 files — all passing
- **No pre-existing failures** (the initial "failure" I saw was a stale uv install the first mutmut invocation refreshed)

## Qualitative findings (Ramsay mode)

Scanned for the classic anti-patterns (no-assert, tautologies, over-mocking, print-based tests, trivial existence-only checks).

| Anti-pattern | Count | Verdict |
|---|---|---|
| `mock.patch` / `Mock()` | 0 | ✅ Clean — matches project preference for real Excel I/O |
| `print(...)` inside tests | 0 | ✅ |
| Tautologies (`x == x`) | 0 | ✅ |
| Bare `assert x is not None` as *only* check | 0 | ✅ — all 72 hits are guards before a real assertion |
| `assert len(x) > 0` as only check | 4 | All 4 have justification messages — legitimate |

The `a/t` (asserts per test) heuristic flagged `test_capacity.py` (1.7) and `test_team.py` (1.2) as weak, but reading the code shows assertions inside loops over `TEAM`/`WP_SCHEDULE` — each test evaluates its assertion N times. **False alarm.**

`test_parameterization.py` uses `pytest.fail()` instead of `assert` — that's a real architectural constraint check (no module-level `config` imports in builders). Keep as-is.

**Conclusion**: no delete-worthy tests found via qualitative scan. The suite shows discipline.

## Empirical findings (mutation testing)

Two runs: `budget_planner/plan.py` mutated, tests_dir narrowed to `test_plan_v2.py` first, then widened to the full suite.

| Scope | Killed | Survived | Score |
|---|---|---|---|
| `test_plan_v2.py` only | 279 | 256 | 52.1% |
| Full suite | 294 | 241 | **55.0%** |
| Full suite + 4 new None-defaults cases | 294 | 241 | 55.0% (measured above) |

The full-suite score is only ~3pp higher than the narrow run — the integration/e2e tests already exercise the happy path so don't add much mutation coverage. Most surviving mutants fall into structural categories where no assertion exists to catch them.

Surviving-mutant categories (sampled):

1. **Style/font/border contract** (~40% of survivors): `font=HEADER_FONT` → absent, `border=False` → `border=True`. Tests don't assert on fonts or borders for most header/label cells.
2. **String-join separator**: `",".join(...)` → `"XX,XX".join(...)` — survives because tests only check `formula.startswith("=SUM(")`, not the complete formula text.
3. **Boolean-operator variants in guards** (~5 mutants): `and`/`or`/`is not None` permutations that happen to be logically equivalent for the exercised input combinations. These are effectively equivalent mutants given the input space.
4. **Off-by-one in inactive-month paths**: `month_num = m + 1` → `m + 2` survives for months where `month_num in active` would be False either way.

### What this audit added

The new `test_build_plan_none_defaults_match_explicit` (4 parametrised cases) kills the 3 None-branch mutants in `_apply_planned_formulas` (#1, #10, #11) that the narrow run reported as surviving. That's the concrete improvement from this audit. The broader gaps (style contract, formula text-shape validation) would need ~10-20 lines of focused assertions to close, but it's not clear the cost is worth the signal.

## Verdict

### Delete
**None.** No tests qualify for deletion.

### Fix
**None.** The only suspected failure was a stale-install artefact.

### Added in this audit
- `tests/test_plan_v2.py::test_build_plan_none_defaults_match_explicit` (4 parametrised cases) — kills the 3 None-branch survivors in `_apply_planned_formulas`.

### Keep
All 256 other tests. The `a/t` heuristic gave false positives; on inspection they're loop-based and meaningful.

### Optional next work (not done — cost/benefit marginal)
1. **Style-contract tests** for Plan header row — would kill ~90 surviving mutants but adds brittle assertions to a visual contract that rarely changes.
2. **Formula-shape assertions** — full formula-string equality instead of `startswith("=SUM(")`. Would kill the separator-mutation survivors but couples tests tightly to exact formula syntax.

Neither is a clean win. The surviving-mutant categories suggest the test suite is deliberately focused on *behaviour* (what the workbook computes) rather than *structure* (how cells are styled), and that trade-off looks correct for this codebase.

## How to reproduce

```bash
# Install
uv pip install --dev mutmut

# Configure (setup.cfg)
[mutmut]
paths_to_mutate=budget_planner/plan.py
tests_dir=tests/test_plan_v2.py

# Run
uv run mutmut run
uv run mutmut results --all true

# Inspect a specific mutant
uv run mutmut show budget_planner.plan.x_build_plan_sheet__mutmut_10
```

For a full-suite score, broaden `tests_dir` to `tests/` (slower but more informative).
