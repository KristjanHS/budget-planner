# Timelog Solver — EUR Accounting Refactor

Date: 2026-04-18
Status: Approved — ready for implementation
Scope: Targeted refactor of `budget_planner/timelog_solver.py` + its callsite in `budget_planner/timelog_plan.py`. No new features; no new sheets.

## 1. Problem

Running the generator and opening `Timelog Plan → Per-WP Summary` shows roughly €5 000 of undershoot on WP7/8/9 and symmetric overshoot on WP1–6, even though ample people capacity exists. Root cause is a unit mismatch between the solver and the sheet:

- Solver stop target (`timelog_plan.py:671`, `_build_wps`): `ceiling_months = round(wp["fte"])`. Greedy fills PM until the **count** hits this target. It never reads €.
- Sheet `Assigned €` formula: months × *actual* per-person rate (Alice Senior €10 000, Bob/Dave Mid €7 500, Carol Junior €5 500).
- Sheet `Ceiling €` formula: `FTE Budget × 6 500` (nominal/avg rate, via `Config!E{wp_row}`).

Result: Carol-staffed WPs land ~3 % short in € at full PM target; Alice/Bob-staffed WPs overshoot. Per-WP PM totals are correct — the € numbers are wrong because the solver's currency of decision (PM) is not the currency the sheet grades it on (€).

User has chosen **option 3 — switch the solver to €-accounting** with **check-after-place semantics**: every WP lands ≥ its € ceiling, with the last placement possibly overshooting by up to `person_rate − 1` €. Undershoot-preserving semantics were considered and rejected: "money left on the table" is the concern, so we want ≥ ceiling € on every WP whose capacity allows.

## 2. Data model

### `Person` gains `rate_eur_per_pm: int`

```python
@dataclass(frozen=True)
class Person:
    name: str
    priority: int
    pool: int
    rate_eur_per_pm: int        # NEW — monthly EUR cost when this person logs 1.0 FTE
```

- **`int`, not `float`:** solver arithmetic is integer accumulation against an integer ceiling. Floats introduce rounding drift (e.g. 9 999.99 + 9 999.99 + 10 000.01 comparisons) and make the check-after-place "overshoot ≤ rate − 1" bound fuzzy. Grade rates in `config.py` are already integer euros; rounding to `int` at the data-model boundary is the single well-defined projection.
- **Field ordering — append at end:** Person currently has 3 positional fields and is constructed positionally in tests (`Person("Alice", priority=1, pool=6)`) and kwarg-style in the builder. Appending `rate_eur_per_pm` at the end keeps all existing positional call sites valid; the frozen dataclass ordering is preserved so no drift risk on legacy callers.

### `WP` renames `ceiling_months → ceiling_eur`

```python
@dataclass(frozen=True)
class WP:
    code: str
    active_start: int
    active_end: int
    ceiling_eur: int            # was ceiling_months; same role, new unit
```

Rename — not "new field + keep old" — because the old field has no role under €-accounting and a dual-field transition would invite "which target does greedy use?" confusion. Integer euros; same justification as above.

## 3. `greedy_pass` loop changes

One counter swap. Today:

```python
placed = 0
...
if placed >= wp.ceiling_months: break
...
assigned[p.name][m] = wp.code
pool[p.name] -= 1
placed += 1
if placed >= wp.ceiling_months: break
```

Under €-accounting:

```python
placed_eur = 0
...
if placed_eur >= wp.ceiling_eur: break
...
assigned[p.name][m] = wp.code
pool[p.name] -= 1
placed_eur += p.rate_eur_per_pm
if placed_eur >= wp.ceiling_eur: break
```

`pool` remains in PM (it tracks person capacity in months, independent of € cost). Only the **WP-level stop target** changes unit.

### Why check-after-place, not check-before

Two candidate semantics:

- **Check-before** (undershoot-preserving): `if placed_eur + p.rate_eur_per_pm > ceiling_eur: break`. Guarantees the WP never exceeds ceiling; can leave up to `max(rate) − 1 ≈ 9 999 €` short per WP.
- **Check-after** (user's choice): place then test. Overshoots by at most `p.rate_eur_per_pm − 1 €` (last placement's own cost minus one euro, since we break iff `placed_eur ≥ ceiling_eur`).

Check-after is the right choice because:

1. The "auditor haircut + partner spillover" buffer rationale from the original timelog design applies identically in the € frame — slight overshoot is a feature, not a bug.
2. Overshoot is neutral on the sheet; undershoot is red. Minimizing undershoot is the point of this refactor.
3. The bound `< person_rate €` is tighter than one PM — a Carol placement overshoots by ≤ €5 499 vs. a full PM of €5 500 under the PM-accounting bound.

## 4. `switch_reduction_pass` — rate-equality guard

Under PM-accounting, Pass 2's adjacent-pair swap `(p1, m, m+1) ↔ (p2, m, m+1)` preserved per-WP totals because each swap moves cells wholesale between persons: WP1 gains a cell on `p2` iff it loses a cell on `p1`, so the PM count is unchanged.

Under €-accounting, this invariant breaks: swapping an Alice (€10 000/PM) cell with a Carol (€5 500/PM) cell on the same WP shifts that WP's `Assigned €` by ±€4 400. Greedy's "placed_eur ≥ ceiling_eur" guarantee is destroyed.

**Fix — rate-equality guard:** Before considering a swap, require the two persons' rates to match.

```python
if persons_by_name[p1_name].rate_eur_per_pm != persons_by_name[p2_name].rate_eur_per_pm:
    continue
```

This preserves per-WP € totals exactly (equal-rate swaps are €-neutral), so the greedy-attained ceiling remains attained after Pass 2.

**Acknowledged loss:** Swaps between persons at different rate tiers (e.g. Alice Senior ↔ Carol Junior) no longer happen. In practice these swaps are unlikely to be useful — rate asymmetry means such a swap also changes each WP's budget usage, which Pass 2 has no visibility into. The equal-rate cohort (e.g. Bob Mid ↔ Dave Mid) is where most switch-reduction value lives, and that case is unaffected.

## 5. `timelog_plan.py` builder integration

Two surgical changes.

### Person construction (~line 182)

```python
persons = [
    Person(
        name=m["name"],
        priority=int(m.get("priority", 2)),
        pool=round(float(m["max_fte"]) * project_months),
        rate_eur_per_pm=int(round(grade_rates[m["grade"]])),   # NEW
    )
    for m in sorted_team
]
```

`grade_rates` is already a builder parameter (currently marked `# noqa: ARG001 — reserved for future per-grade tinting`). This refactor puts it to actual use. Remove the `noqa` and update the docstring.

### `_build_wps` (line 653+)

```python
ceiling_eur = int(round(float(wp["fte"]) * float(wp["rate"])))
wps.append(WP(
    code=wp["id"],
    active_start=active_start,
    active_end=active_end,
    ceiling_eur=ceiling_eur,
))
```

The `rate` key already exists on `work_packages` dicts (used elsewhere for the same `fte × rate` multiplication). The sheet's `Per-WP Summary.Ceiling €` column cross-refs `Config!E{wp_row}` which also equals `fte × rate` — so the solver's ceiling and the sheet's displayed ceiling agree by construction. (Minor arithmetic note: `Config!E` may display `=B×D` with un-rounded inputs; `int(round(...))` in Python will agree to the cent for integer `fte` and integer `rate`, which is the data we have. If either is fractional, a ±1 € disagreement is possible and tolerable — it stays inside the allowed overshoot band.)

No reader change. No new cross-sheet trust boundary — the sheet still displays what it always did; only the solver's internal target changes.

## 6. Test impact

`tests/test_timelog_solver.py` constructs `Person` and `WP` directly. Naive migration would require rewriting every fixture with realistic € amounts.

**Strategy — neutralise the € axis with unit rates:** Pass `rate_eur_per_pm=1` for all test persons and `ceiling_eur=<old ceiling_months value>` for all test WPs. Under this mapping, `placed_eur += 1` per cell and `placed_eur ≥ ceiling_eur` stops at the same count — **€-accounting becomes numerically identical to PM-accounting**. Every existing assertion (`_placed_count == ceiling`, pool bounds, window respect, Pass 2 preservation) still holds with this one-line-per-fixture edit.

**One new test** for the differential-rate behavior the refactor actually enables:

```python
def test_greedy_stops_at_eur_ceiling_with_mixed_rates():
    # Alice rate=2, Bob rate=1; WP wants 5 €. Greedy prefers Alice (prio 1,
    # lex order). Alice placements: 2 € each; Bob placements: 1 € each.
    # Expected stop: check-after-place, overshoot ≤ rate-1 = 1 €.
    persons = [
        Person("Alice", priority=1, pool=5, rate_eur_per_pm=2),
        Person("Bob",   priority=2, pool=5, rate_eur_per_pm=1),
    ]
    wps = [WP("WP1", active_start=1, active_end=10, ceiling_eur=5)]
    windows = [
        PersonWPWindow("Alice", "WP1", 1, 10),
        PersonWPWindow("Bob",   "WP1", 1, 10),
    ]
    assigned = greedy_pass(persons, wps, windows, project_months=10)
    alice_pm = sum(1 for v in assigned["Alice"].values() if v == "WP1")
    bob_pm   = sum(1 for v in assigned["Bob"].values()   if v == "WP1")
    placed_eur = alice_pm * 2 + bob_pm * 1
    assert placed_eur >= 5
    assert placed_eur <= 5 + (2 - 1)  # overshoot bound = max(rate) - 1
```

Do not pad tests beyond this. The existing suite, ported via the unit-rate trick, already covers determinism / priority / pool / windows / Pass 2. The single new test locks in the one behaviour the refactor is actually about.

## 7. Docstring + invariants updates

- **`timelog_solver.py` module docstring:** the bullet "WP totals approach their ceiling_months target" becomes "WP totals approach their `ceiling_eur` target (slight € overshoot tolerated, bounded by the placing person's rate)". The Pass 2 paragraph gains one sentence: "Under €-accounting, Pass 2 requires rate-equality between the two persons; swaps across rate tiers are skipped to preserve per-WP € totals exactly."
- **`Person.rate_eur_per_pm` attribute docstring:** "Integer EUR cost per 1.0-FTE month for this person. Used by the greedy pass for €-accounting and by Pass 2 as a swap-eligibility key."
- **`WP.ceiling_eur` attribute docstring:** "Target fill in integer EUR = `round(wp.fte × wp.rate)`."
- **`docs/design-invariants.md`:** grep for `Timelog|timelog|ceiling_months|solver` at implementation time; if any bullet claims the solver "fills PM targets" or similar, update to "fills € targets". (Current grep shows only file-path mentions — no stale PM-target claim to correct — but re-check on implementation.)

## 8. Sequenced implementation steps

1. **Data model edit** — `timelog_solver.py`: add `Person.rate_eur_per_pm: int`; rename `WP.ceiling_months → WP.ceiling_eur`. Update attribute docstrings.
2. **`greedy_pass` loop edit** — swap `placed: int` counter to `placed_eur: int`; `placed += 1` → `placed_eur += p.rate_eur_per_pm`; stop-condition identifiers. Update inline comment block about check-after-place to mention the new `< person_rate €` overshoot bound.
3. **`switch_reduction_pass` guard** — add `persons_by_name` dict lookup at pass entry; add rate-equality `continue` guard inside the `p2_name` loop, before the legality checks. Update the preservation comment in the pass docstring.
4. **Module docstring** — update the two bullets (ceiling unit + Pass 2 rate-equality). *Checkpoint: run `uv run pytest tests/test_timelog_solver.py -q` and `uv run pyright budget_planner/timelog_solver.py`. Suite fails until step 5 rewrites fixtures — expected.*
5. **Test fixtures** — `tests/test_timelog_solver.py`: mechanical edit — add `rate_eur_per_pm=1` to every `Person(...)`; rename `ceiling_months=` → `ceiling_eur=` on every `WP(...)`. Add the one new `test_greedy_stops_at_eur_ceiling_with_mixed_rates`. *Checkpoint: `pytest tests/test_timelog_solver.py -v` — all green, new test included.*
6. **Builder callsite** — `timelog_plan.py`: add `rate_eur_per_pm=int(round(grade_rates[m["grade"]]))` in the `persons = [...]` comprehension; in `_build_wps`, compute `ceiling_eur = int(round(float(wp["fte"]) * float(wp["rate"])))` and pass it to the `WP(...)` call. Remove the `# noqa: ARG001` on `grade_rates` and update its docstring line. *Checkpoint: `pytest tests/ -q` full suite + `uv run pyright`.*
7. **End-to-end verification** — `python scripts/init_budget.py && python scripts/recalc.py output/budget_plan.xlsx`. Open in LibreOffice (or Excel per dev-workflow rule) → `Per-WP Summary` → confirm `Undershoot €` column is near-zero across all WPs (bounded by rate-1 on any WP that was previously undershot). Confirm `Overshoot €` is bounded by `max(rate) = €10 000` on all WPs.
8. **Docstring + invariants doc commit** — adjust any prose in `docs/design-invariants.md` that still implies PM accounting; update `.claude/rules/` only if a trigger references `ceiling_months` (grep to check).

Steps 1–6 land as one commit (intermediate states fail the test suite). Step 8 can be a follow-up commit.

## 9. Out of scope / deferred

- **Not** fixing `Per-WP Summary.Undershoot €` display — it naturally reads near-zero post-change; no formula edit needed.
- **Not** changing `reader.py` — no new cross-sheet trust boundary; no new reader helper.
- **Not** changing the sheet layout, CF, or named ranges on Timelog Plan.
- **Not** touching Pass 2's adjacent-pair primitive shape (the Q16 limitation from the original design doc still stands).
- **Not** backfilling historical workbooks — this is a code-only refactor; any workbook rebuilt after the change gets the new behaviour automatically.

## Decisions log

| # | Decision | Rationale |
|---|---|---|
| D1 | Check-after-place (overshoot ≤ rate − 1 €), not check-before (undershoot ≤ max(rate) − 1 €) | Undershoot is red on the sheet; overshoot is neutral and matches the existing "auditor haircut buffer" rationale. User explicitly chose this. |
| D2 | Rate-equality guard on Pass 2 swaps | Under €-accounting, only equal-rate swaps preserve per-WP € totals. Equal-rate-cohort swaps capture most of the useful optimization (e.g. Bob ↔ Dave at Mid rate); cross-tier swaps are skipped as acceptable loss. |
| D3 | Integer € arithmetic (Person.rate_eur_per_pm: int, WP.ceiling_eur: int) | Eliminates float drift in the overshoot-bound comparison; grade rates are integer euros in `config.py` already. |
| D4 | Rename `ceiling_months → ceiling_eur` (not add a second field) | Old field has no role under €-accounting; keeping it invites "which target does greedy use?" confusion. |
| D5 | Test migration via `rate_eur_per_pm=1`, `ceiling_eur=<old months>` — one new test for differential rates | Preserves existing coverage with minimal fixture churn; locks in the one genuinely new behaviour. |
| D6 | `int(round(grade_rates[grade]))` at the builder boundary | Single well-defined projection from the `Mapping[str, float]` config type to the solver's `int` field. |

## Correction (applied during /impag, 2026-04-18)

Two assumptions in §5 and §6 turned out to be wrong at implementation time; both were resolved inline without changing the design intent:

1. **`wp["rate"]` is not present on init.** §5 claimed "The `rate` key already exists on `work_packages` dicts". True post-`reader.read_workbook_config`, but the seed `WORK_PACKAGES` in `config.py` has only `id`, `name`, `fte`. `_build_wps` now takes `project_params` and falls back to `project_params["rate"]` when the WP dict lacks `rate` — mirroring the existing pattern in `config_sheet.py:102-125`. No change to the sheet-vs-solver ceiling agreement: both paths project through the same `FTE × Rate` formula.
2. **Not every team-dict fixture carries `grade`.** §6 migration strategy ported `test_timelog_solver.py` fixtures with unit rates, but `tests/test_trends.py::_make_team` (a fixture consumed transitively by the full-workbook build path) omitted `grade`, tripping a `KeyError` inside `build_timelog_plan_sheet`. Fixed by adding `"grade": "Mid"` to the fixture — honest, since `grade` has always been documented as a required field on team dicts.

End-to-end verification confirmed the design intent: on `init --demo`, `Per-WP Summary.Undershoot €` is 0 across WP1–8 (was ~€3 000–€5 000 under PM-accounting on WP7–9) and `€1 200` on WP9 (pool-limited, not ceiling-miss). `Overshoot €` bounded by the placing person's rate on every WP that lands short of one more placement.
