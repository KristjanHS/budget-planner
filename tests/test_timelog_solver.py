"""Unit tests for the Timelog Plan solver.

Hand-crafted fixtures only — no dependency on ``budget_planner.config`` or the
workbook builders. These tests assert algorithmic properties of ``solve``,
``greedy_pass``, and ``switch_reduction_pass``.
"""

from __future__ import annotations

from budget_planner.timelog_solver import (
    Assignment,
    Person,
    PersonWPWindow,
    WP,
    _count_switches,
    greedy_pass,
    solve,
    switch_reduction_pass,
)


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------


def _placed_count(assigned: Assignment, wp_code: str) -> int:
    return sum(
        1
        for row in assigned.values()
        for v in row.values()
        if v == wp_code
    )


def _assignments(
    assigned: Assignment,
) -> list[tuple[str, int, str]]:
    """Flatten to a sorted list of (person, month, wp_code) for non-None cells."""
    out: list[tuple[str, int, str]] = []
    for name, row in assigned.items():
        for m, code in row.items():
            if code is not None:
                out.append((name, m, code))
    out.sort()
    return out


# ---------------------------------------------------------------------------
# 1. Determinism
# ---------------------------------------------------------------------------


def test_determinism_identical_inputs():
    persons = [
        Person("Alice", priority=1, pool=6, rate_eur_per_pm=1),
        Person("Bob", priority=2, pool=6, rate_eur_per_pm=1),
    ]
    wps = [
        WP("WP1", active_start=1, active_end=6, ceiling_eur=4),
        WP("WP2", active_start=3, active_end=8, ceiling_eur=4),
    ]
    windows = [
        PersonWPWindow("Alice", "WP1", 1, 6),
        PersonWPWindow("Alice", "WP2", 3, 8),
        PersonWPWindow("Bob", "WP1", 1, 6),
        PersonWPWindow("Bob", "WP2", 3, 8),
    ]
    a = solve(persons, wps, windows, project_months=8)
    b = solve(persons, wps, windows, project_months=8)
    assert a == b


# ---------------------------------------------------------------------------
# 2. Earliest-close WP filled first
# ---------------------------------------------------------------------------


def test_earliest_close_wp_filled_first():
    # WP1 closes at month 6, WP2 at month 10. Both want 4 months; one person
    # with pool=6. Greedy should fill WP1 fully before touching WP2.
    persons = [Person("Alice", priority=1, pool=6, rate_eur_per_pm=1)]
    wps = [
        WP("WP2", active_start=1, active_end=10, ceiling_eur=4),
        WP("WP1", active_start=1, active_end=6, ceiling_eur=4),
    ]
    windows = [
        PersonWPWindow("Alice", "WP1", 1, 6),
        PersonWPWindow("Alice", "WP2", 1, 10),
    ]
    assigned = greedy_pass(persons, wps, windows, project_months=10)
    # WP1 is filled (4 months) before pool runs out; WP2 gets the remaining 2.
    assert _placed_count(assigned, "WP1") == 4
    assert _placed_count(assigned, "WP2") == 2


# ---------------------------------------------------------------------------
# 3. Priority tiebreaker within a WP
# ---------------------------------------------------------------------------


def test_priority_tiebreaker_within_wp():
    # Both persons can fill WP1, but WP1 wants 4 months and each person has
    # pool=3 → neither can fill alone. Priority-1 should be fully exhausted
    # before priority-2 is touched.
    persons = [
        Person("Bob", priority=2, pool=3, rate_eur_per_pm=1),  # lex-first but lower priority
        Person("Alice", priority=1, pool=3, rate_eur_per_pm=1),
    ]
    wps = [WP("WP1", active_start=1, active_end=6, ceiling_eur=4)]
    windows = [
        PersonWPWindow("Alice", "WP1", 1, 6),
        PersonWPWindow("Bob", "WP1", 1, 6),
    ]
    assigned = greedy_pass(persons, wps, windows, project_months=6)
    alice_count = sum(1 for v in assigned["Alice"].values() if v == "WP1")
    bob_count = sum(1 for v in assigned["Bob"].values() if v == "WP1")
    # Alice (prio 1) fills all her 3 months first, Bob (prio 2) picks up 1.
    assert alice_count == 3
    assert bob_count == 1


# ---------------------------------------------------------------------------
# 4. Pool never exceeded
# ---------------------------------------------------------------------------


def test_pool_never_exceeded():
    persons = [
        Person("Alice", priority=1, pool=3, rate_eur_per_pm=1),
        Person("Bob", priority=1, pool=5, rate_eur_per_pm=1),
    ]
    wps = [
        WP("WP1", active_start=1, active_end=8, ceiling_eur=6),
        WP("WP2", active_start=1, active_end=8, ceiling_eur=6),
    ]
    windows = [
        PersonWPWindow("Alice", "WP1", 1, 8),
        PersonWPWindow("Alice", "WP2", 1, 8),
        PersonWPWindow("Bob", "WP1", 1, 8),
        PersonWPWindow("Bob", "WP2", 1, 8),
    ]
    assigned = solve(persons, wps, windows, project_months=8)
    for p in persons:
        placed = sum(1 for v in assigned[p.name].values() if v is not None)
        assert placed <= p.pool, (
            f"{p.name}: placed {placed} > pool {p.pool}"
        )


# ---------------------------------------------------------------------------
# 5. Person-WP window is honoured
# ---------------------------------------------------------------------------


def test_person_wp_window_honoured():
    persons = [Person("Alice", priority=1, pool=12, rate_eur_per_pm=1)]
    wps = [
        WP("WP1", active_start=1, active_end=12, ceiling_eur=6),
        WP("WP2", active_start=1, active_end=12, ceiling_eur=6),
    ]
    windows = [
        # Alice can only do WP1 in months 1-4, WP2 in months 7-10.
        PersonWPWindow("Alice", "WP1", 1, 4),
        PersonWPWindow("Alice", "WP2", 7, 10),
    ]
    assigned = solve(persons, wps, windows, project_months=12)

    windows_idx = {(w.person_name, w.wp_code): (w.start_month, w.end_month)
                   for w in windows}
    for name, m, code in _assignments(assigned):
        lo, hi = windows_idx[(name, code)]
        assert lo <= m <= hi, (
            f"({name}, m={m}) -> {code} outside person-WP window {lo}-{hi}"
        )


# ---------------------------------------------------------------------------
# 6. WP active window is honoured
# ---------------------------------------------------------------------------


def test_wp_window_honoured():
    persons = [Person("Alice", priority=1, pool=12, rate_eur_per_pm=1)]
    wps = [
        WP("WP1", active_start=3, active_end=8, ceiling_eur=10),
    ]
    windows = [PersonWPWindow("Alice", "WP1", 1, 12)]
    assigned = solve(persons, wps, windows, project_months=12)

    wps_by_code = {w.code: w for w in wps}
    for name, m, code in _assignments(assigned):
        wp = wps_by_code[code]
        assert wp.active_start <= m <= wp.active_end


# ---------------------------------------------------------------------------
# 7. Overshoot tolerated (at most 1)
# ---------------------------------------------------------------------------


def test_overshoot_tolerated():
    # One person, pool=8, WP wants 3 with window big enough — check-after-
    # placing means at worst we'd place one past the ceiling. Since we place
    # in a single inner loop that breaks *after* incrementing, the actual
    # count lands exactly at ceiling when possible (no overshoot needed).
    # To force an overshoot, set ceiling=3 and ensure the greedy loop can
    # legitimately land on ceiling+1. The spec says "allow overshoot by one
    # month" — we assert placed ∈ {ceiling, ceiling+1}.
    persons = [Person("Alice", priority=1, pool=10, rate_eur_per_pm=1)]
    wps = [WP("WP1", active_start=1, active_end=10, ceiling_eur=3)]
    windows = [PersonWPWindow("Alice", "WP1", 1, 10)]
    assigned = greedy_pass(persons, wps, windows, project_months=10)
    placed = _placed_count(assigned, "WP1")
    assert placed in (3, 4), (
        f"expected 3 or 4 (ceiling or ceiling+1), got {placed}"
    )
    # Crucially: never placed = ceiling + 2 or more.
    assert placed <= 4


# ---------------------------------------------------------------------------
# 8. Undershoot when capacity short
# ---------------------------------------------------------------------------


def test_undershoot_when_capacity_short():
    # WP wants 6 months but only one person with pool=3 can fill it.
    persons = [Person("Alice", priority=1, pool=3, rate_eur_per_pm=1)]
    wps = [WP("WP1", active_start=1, active_end=8, ceiling_eur=6)]
    windows = [PersonWPWindow("Alice", "WP1", 1, 8)]
    assigned = greedy_pass(persons, wps, windows, project_months=8)
    placed = _placed_count(assigned, "WP1")
    assert placed == 3
    assert placed < 6


# ---------------------------------------------------------------------------
# 9. Pass 2 preserves switch count (never worse)
# ---------------------------------------------------------------------------


def test_switch_reduction_reduces_or_preserves_switches():
    # Scenario designed so greedy leaves a swappable pattern. Two persons,
    # both eligible for WP1 and WP2, with interleaved assignments that a
    # swap can consolidate.
    persons = [
        Person("Alice", priority=1, pool=4, rate_eur_per_pm=1),
        Person("Bob", priority=1, pool=4, rate_eur_per_pm=1),
    ]
    wps = [
        WP("WP1", active_start=1, active_end=8, ceiling_eur=4),
        WP("WP2", active_start=1, active_end=8, ceiling_eur=4),
    ]
    windows = [
        PersonWPWindow("Alice", "WP1", 1, 8),
        PersonWPWindow("Alice", "WP2", 1, 8),
        PersonWPWindow("Bob", "WP1", 1, 8),
        PersonWPWindow("Bob", "WP2", 1, 8),
    ]
    greedy = greedy_pass(persons, wps, windows, project_months=8)
    # Precondition: the fixture must actually give the swap primitive something
    # to work with. If greedy already leaves zero switches, this test would
    # pass trivially (0 ≤ 0) and silently stop exercising switch_reduction_pass
    # — breaking its load-bearing property assertion. Lock in that the scenario
    # produces a multi-switch starting state.
    greedy_switches = _count_switches(greedy)
    assert greedy_switches >= 2, (
        f"precondition: greedy fixture should leave ≥2 switches for "
        f"switch_reduction_pass to act on, got {greedy_switches}"
    )

    after = switch_reduction_pass(greedy, persons, wps, windows, project_months=8)
    assert _count_switches(after) <= greedy_switches


def test_switch_reduction_actually_reduces_on_interleaved_input():
    # Hand-built 2-block swap scenario. Solver's swap shape is
    # (p1, m, m+1) <-> (p2, m, m+1); the outer loop also requires
    # p1 to have an internal switch at (m, m+1) — an end-block swap where
    # p1.(m, m+1) is a flat run is skipped by the `a1 == a2` guard.
    # The fixture below has internal switches on BOTH persons so the swap is
    # considered and the result is a strict reduction, not just a no-op.
    persons = [
        Person("Alice", priority=1, pool=3, rate_eur_per_pm=1),
        Person("Bob", priority=1, pool=3, rate_eur_per_pm=1),
    ]
    wps = [
        WP("WP1", active_start=1, active_end=3, ceiling_eur=3),
        WP("WP2", active_start=1, active_end=3, ceiling_eur=3),
    ]
    windows = [
        PersonWPWindow("Alice", "WP1", 1, 3),
        PersonWPWindow("Alice", "WP2", 1, 3),
        PersonWPWindow("Bob", "WP1", 1, 3),
        PersonWPWindow("Bob", "WP2", 1, 3),
    ]
    # Starting state — fully interleaved, 4 total switches:
    #   Alice = WP1, WP2, WP1   → 2 switches
    #   Bob   = WP2, WP1, WP2   → 2 switches
    # Swap (m=1, m+1=2) between Alice and Bob yields:
    #   Alice = WP2, WP1, WP1   → 1 switch
    #   Bob   = WP1, WP2, WP2   → 1 switch
    # → total 2; per-WP sums (WP1×3, WP2×3) preserved; pools unchanged.
    start: Assignment = {
        "Alice": {1: "WP1", 2: "WP2", 3: "WP1"},
        "Bob":   {1: "WP2", 2: "WP1", 3: "WP2"},
    }
    before = _count_switches(start)
    assert before == 4, f"precondition: fixture should start with 4 switches, got {before}"

    after = switch_reduction_pass(start, persons, wps, windows, project_months=3)
    after_switches = _count_switches(after)
    assert after_switches < before, (
        f"switch_reduction_pass should consolidate runs: "
        f"before={before}, after={after_switches}"
    )

    # Per-WP totals must be preserved (design §4 invariant).
    def _wp_sum(assigned: Assignment, wp_code: str) -> int:
        return sum(
            1 for p_row in assigned.values() for v in p_row.values() if v == wp_code
        )
    assert _wp_sum(after, "WP1") == _wp_sum(start, "WP1")
    assert _wp_sum(after, "WP2") == _wp_sum(start, "WP2")


# ---------------------------------------------------------------------------
# 10. Pass 2 preserves per-WP totals
# ---------------------------------------------------------------------------


def test_switch_reduction_preserves_wp_sums():
    persons = [
        Person("Alice", priority=1, pool=6, rate_eur_per_pm=1),
        Person("Bob", priority=1, pool=6, rate_eur_per_pm=1),
        Person("Carol", priority=2, pool=4, rate_eur_per_pm=1),
    ]
    wps = [
        WP("WP1", active_start=1, active_end=10, ceiling_eur=5),
        WP("WP2", active_start=1, active_end=10, ceiling_eur=5),
        WP("WP3", active_start=1, active_end=10, ceiling_eur=4),
    ]
    windows = [
        PersonWPWindow("Alice", "WP1", 1, 10),
        PersonWPWindow("Alice", "WP2", 1, 10),
        PersonWPWindow("Alice", "WP3", 1, 10),
        PersonWPWindow("Bob", "WP1", 1, 10),
        PersonWPWindow("Bob", "WP2", 1, 10),
        PersonWPWindow("Bob", "WP3", 1, 10),
        PersonWPWindow("Carol", "WP1", 1, 10),
        PersonWPWindow("Carol", "WP2", 1, 10),
        PersonWPWindow("Carol", "WP3", 1, 10),
    ]
    greedy = greedy_pass(persons, wps, windows, project_months=10)
    after = switch_reduction_pass(greedy, persons, wps, windows, project_months=10)
    for wp in wps:
        assert _placed_count(greedy, wp.code) == _placed_count(after, wp.code)


# ---------------------------------------------------------------------------
# 11. No PersonWPWindow row → excluded
# ---------------------------------------------------------------------------


def test_no_person_wp_row_means_excluded():
    # Alice has no row for WP1; even with pool and overlapping active window,
    # she must never be assigned to WP1. Bob is eligible and will fill it.
    persons = [
        Person("Alice", priority=1, pool=10, rate_eur_per_pm=1),
        Person("Bob", priority=2, pool=10, rate_eur_per_pm=1),
    ]
    wps = [WP("WP1", active_start=1, active_end=8, ceiling_eur=4)]
    windows = [
        # Alice only has a (dangling) row for a non-existent WP — irrelevant.
        PersonWPWindow("Bob", "WP1", 1, 8),
    ]
    assigned = solve(persons, wps, windows, project_months=8)
    # Alice never touches WP1.
    for m, code in assigned["Alice"].items():
        assert code != "WP1", f"Alice got WP1 at month {m} despite no window row"
    # Bob actually filled it.
    assert _placed_count(assigned, "WP1") == 4


# ---------------------------------------------------------------------------
# 12. €-accounting: greedy stops at ceiling_eur with mixed person rates
# ---------------------------------------------------------------------------


def test_greedy_stops_at_eur_ceiling_with_mixed_rates():
    # Alice rate=2, Bob rate=1; WP wants 5 €. Greedy prefers Alice (prio 1,
    # lex order). Alice placements: 2 € each; Bob placements: 1 € each.
    # Expected stop: check-after-place, overshoot ≤ rate-1 = 1 €.
    persons = [
        Person("Alice", priority=1, pool=5, rate_eur_per_pm=2),
        Person("Bob", priority=2, pool=5, rate_eur_per_pm=1),
    ]
    wps = [WP("WP1", active_start=1, active_end=10, ceiling_eur=5)]
    windows = [
        PersonWPWindow("Alice", "WP1", 1, 10),
        PersonWPWindow("Bob", "WP1", 1, 10),
    ]
    assigned = greedy_pass(persons, wps, windows, project_months=10)
    alice_pm = sum(1 for v in assigned["Alice"].values() if v == "WP1")
    bob_pm = sum(1 for v in assigned["Bob"].values() if v == "WP1")
    placed_eur = alice_pm * 2 + bob_pm * 1
    assert placed_eur >= 5
    # Overshoot bound = max(rate) - 1 across placing persons.
    max_rate = max(p.rate_eur_per_pm for p in persons)
    assert placed_eur <= 5 + (max_rate - 1)


# ---------------------------------------------------------------------------
# 13. Vacation months (2026-04-19 enhancements)
# ---------------------------------------------------------------------------


def test_solver_skips_vacation():
    """Greedy never places a WP in any month inside a Person.vacation_months set."""
    # WP wants every month filled; all months in window. Alice has months 3 and 7
    # as vacation — solver must skip both, placing at most 8 cells (10 - 2).
    persons = [
        Person(
            "Alice", priority=1, pool=10, rate_eur_per_pm=1,
            vacation_months=frozenset({3, 7}),
        ),
    ]
    wps = [WP("WP1", active_start=1, active_end=10, ceiling_eur=10)]
    windows = [PersonWPWindow("Alice", "WP1", 1, 10)]
    assigned = solve(persons, wps, windows, project_months=10)
    placed_months = {m for m, code in assigned["Alice"].items() if code == "WP1"}
    # Vacation months must never appear.
    assert 3 not in placed_months
    assert 7 not in placed_months
    # Bystander: non-vacation months are placed (solver didn't skip everything).
    assert len(placed_months) >= 1
    assert placed_months.issubset({1, 2, 4, 5, 6, 8, 9, 10})


def test_switch_reduction_respects_vacation():
    """Switch-reduction pass won't swap a cell into or out of a vacation month."""
    persons = [
        Person(
            "Alice", priority=1, pool=4, rate_eur_per_pm=1,
            vacation_months=frozenset({4}),
        ),
        Person(
            "Bob", priority=1, pool=4, rate_eur_per_pm=1,
            vacation_months=frozenset({4}),
        ),
    ]
    wps = [
        WP("WP1", active_start=1, active_end=5, ceiling_eur=4),
        WP("WP2", active_start=1, active_end=5, ceiling_eur=4),
    ]
    windows = [
        PersonWPWindow("Alice", "WP1", 1, 5),
        PersonWPWindow("Alice", "WP2", 1, 5),
        PersonWPWindow("Bob", "WP1", 1, 5),
        PersonWPWindow("Bob", "WP2", 1, 5),
    ]
    # Pre-seed an assignment that would tempt the pass to swap across month 4.
    pre_assigned: Assignment = {
        "Alice": {1: "WP1", 2: "WP2", 3: "WP1", 4: None, 5: "WP2"},
        "Bob":   {1: "WP2", 2: "WP1", 3: "WP2", 4: None, 5: "WP1"},
    }
    out = switch_reduction_pass(
        pre_assigned, persons, wps, windows, project_months=5,
    )
    assert out["Alice"][4] is None, "vacation month must stay empty for Alice"
    assert out["Bob"][4] is None, "vacation month must stay empty for Bob"


def test_vacation_months_default_empty():
    """Person instantiated without vacation_months has an empty frozenset default
    so existing callers (and legacy tests) don't need to pass the field."""
    p = Person("Alice", priority=1, pool=1, rate_eur_per_pm=1)
    assert p.vacation_months == frozenset()
