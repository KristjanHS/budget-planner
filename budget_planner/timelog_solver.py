"""Timelog Plan solver.

Pure-Python, deterministic, no I/O and no openpyxl dependency. Assigns each
(person, month) cell to exactly one WP code (or None) such that:

- WP totals approach their ``ceiling_eur`` target. Overshoot is bounded by
  ``MAX_BUFFER_PCT`` (currently +12% of the WP ceiling); undershoot is
  bounded by ``MIN_BUFFER_PCT`` (currently -11%). When the only candidate's
  rate would push past +12% AND skipping would leave the WP below -11%, the
  placement happens anyway — overshoot wins over undershoot. WPs with no
  remaining candidates may still end up below -11%; that is a capacity
  shortfall, not a policy choice.
- Each cell stays inside the WP's active window AND inside the person's
  per-(person, WP) window (the PERSON-WP ASSIGNMENTS row).
- Each person's total assigned months never exceeds their pool.
- Within a WP, priority-1 persons fill before priority-2 before priority-3.
- Contiguous runs of the same WP are preferred (reduces switch count).
- Earliest-closing WPs fill first.

Two passes:
1. Greedy fill (WPs in earliest-close order, priority-ordered candidates,
   continue-a-run preference).
2. Switch-reduction local improvement — swap adjacent-month blocks between
   two persons to reduce total switches without breaking any invariant.
   Under €-accounting, Pass 2 requires rate-equality between the two
   persons; swaps across rate tiers are skipped to preserve per-WP €
   totals exactly.

Known shape limitation (design §9 Q16): ``switch_reduction_pass`` considers
only adjacent-pair swaps — it exchanges ``(p1, m, m+1)`` with
``(p2, m, m+1)``. Rearrangements that require a 3+ month block swap, or a
same-month-window rotation among three persons, are not found. In practice
greedy leaves few such patterns; the algorithm trades optimality for
determinism + bounded runtime. A future extension could add a second
adjacent-pair pass with wider windows or a 3-way rotation primitive.

Cross-sheet invariants in ``docs/spec.md`` § Timelog Plan merged grid.
"""

from __future__ import annotations

from dataclasses import dataclass

# Per-WP buffer = (placed_eur − ceiling_eur) / ceiling_eur.
# Greedy skips any placement that would push placed_eur past
# ceiling_eur × (1 + MAX_BUFFER_PCT) — UNLESS skipping would leave the WP
# below ceiling_eur × (1 − MIN_BUFFER_PCT). In that tie-break case the
# placement happens (overshoot wins over deeper undershoot). WPs with no
# remaining candidates can still finish below the lower cap; the policy
# only fires when a placement is available.
MAX_BUFFER_PCT = 0.12
MIN_BUFFER_PCT = 0.11


@dataclass(frozen=True)
class Person:
    """A team member eligible for WP allocation.

    Attributes:
        name: Unique identifier (also used as a deterministic sort key).
        priority: 1, 2, or 3. Lower = filled first within each WP pass.
        pool: Total assignable months = round(max_fte * project_months).
        rate_eur_per_pm: Integer EUR cost per 1.0-FTE month for this
            person. Used by the greedy pass for €-accounting and by Pass 2
            as a swap-eligibility key.
        vacation_months: Project-1-indexed month numbers this person must not
            be assigned to (vacation). Default empty. For the current uniform-
            July model this is identical across every Person; the per-person
            field is what keeps the solver contract stable when a future
            design ships per-(person, year) overrides.
        frozen_months: Project-1-indexed month numbers whose assignee-grid
            cells are preserved verbatim by the builder (from the rebuilding
            workbook's preserved grid). The solver never places a WP code
            in a frozen month and never swaps across one. Frozen consumption
            is accounted for in the builder via ``ceiling_eur`` / ``pool``
            narrowing before ``solve`` is called — mirrors the vacation
            guard shape.
    """

    name: str
    priority: int
    pool: int
    rate_eur_per_pm: int
    vacation_months: frozenset[int] = frozenset()
    frozen_months: frozenset[int] = frozenset()


@dataclass(frozen=True)
class WP:
    """A work package with an active window and a fill target.

    Attributes:
        code: "WP1", "WP2", ... — also the value written into assigned cells.
        active_start: 1-indexed month in project frame, inclusive.
        active_end: 1-indexed month in project frame, inclusive.
        ceiling_eur: Target fill in integer EUR = ``round(wp.fte × wp.rate)``.
    """

    code: str
    active_start: int
    active_end: int
    ceiling_eur: int


@dataclass(frozen=True)
class PersonWPWindow:
    """A PERSON-WP ASSIGNMENTS row. Hard constraint: (person, wp) can only
    be placed inside this window (further intersected with WP active window).
    """

    person_name: str
    wp_code: str
    start_month: int
    end_month: int


# Return type alias: assignment[person_name][month_1based] -> wp_code or None.
Assignment = dict[str, dict[int, str | None]]


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------


def _build_windows_index(
    person_wp_windows: list[PersonWPWindow],
) -> dict[tuple[str, str], tuple[int, int]]:
    """Index PersonWPWindow rows by (person_name, wp_code) -> (start, end).

    If multiple rows exist for the same pair (not expected in practice), the
    widest enclosing window is used.
    """
    idx: dict[tuple[str, str], tuple[int, int]] = {}
    for w in person_wp_windows:
        key = (w.person_name, w.wp_code)
        existing = idx.get(key)
        if existing is None:
            idx[key] = (w.start_month, w.end_month)
        else:
            idx[key] = (
                min(existing[0], w.start_month),
                max(existing[1], w.end_month),
            )
    return idx


def _empty_assignment(persons: list[Person], project_months: int) -> Assignment:
    return {
        p.name: {m: None for m in range(1, project_months + 1)}
        for p in persons
    }


def _most_recent_wp_before(
    assigned: Assignment, person_name: str, month: int
) -> str | None:
    """Return the WP code assigned to person at the largest month < `month`
    that has a non-None assignment, or None if no prior placement.
    """
    row = assigned[person_name]
    for prev in range(month - 1, 0, -1):
        code = row[prev]
        if code is not None:
            return code
    return None


def _count_switches(assigned: Assignment) -> int:
    """Total adjacent-month switches across all persons.

    Counts empty <-> WP transitions too (design §9 Q15a). Within-run cells do
    not count.
    """
    total = 0
    for row in assigned.values():
        months = sorted(row.keys())
        for i in range(len(months) - 1):
            if row[months[i]] != row[months[i + 1]]:
                total += 1
    return total


def _cell_legal_for_person(
    person_name: str,
    wp_code: str | None,
    month: int,
    wps_by_code: dict[str, WP],
    windows_idx: dict[tuple[str, str], tuple[int, int]],
) -> bool:
    """True iff placing `wp_code` at (person, month) respects both the WP
    active window and the person-WP window. `None` is always legal.
    """
    if wp_code is None:
        return True
    wp = wps_by_code.get(wp_code)
    if wp is None:
        return False
    if not (wp.active_start <= month <= wp.active_end):
        return False
    pw = windows_idx.get((person_name, wp_code))
    if pw is None:
        return False
    return pw[0] <= month <= pw[1]


# ---------------------------------------------------------------------------
# Pass 1: greedy fill
# ---------------------------------------------------------------------------


def greedy_pass(
    persons: list[Person],
    wps: list[WP],
    person_wp_windows: list[PersonWPWindow],
    project_months: int,
) -> Assignment:
    """Run the greedy fill pass. Pure function; does not mutate inputs."""
    windows_idx = _build_windows_index(person_wp_windows)
    pool = {p.name: p.pool for p in persons}
    vacations = {p.name: p.vacation_months for p in persons}
    frozen = {p.name: p.frozen_months for p in persons}
    assigned = _empty_assignment(persons, project_months)

    # WPs in earliest-close order; code as deterministic tiebreaker.
    wp_order = sorted(wps, key=lambda w: (w.active_end, w.code))

    for wp in wp_order:
        placed_eur = 0
        for prio in (1, 2, 3):
            if placed_eur >= wp.ceiling_eur:
                break
            # Eligible persons at this priority who have a row for this WP.
            candidates_raw = [
                p
                for p in persons
                if p.priority == prio
                and (p.name, wp.code) in windows_idx
                and pool[p.name] > 0
            ]
            if not candidates_raw:
                continue

            # Sort: continue-a-run first, then name ascending. We recompute
            # "running" once at the top of the WP pass — the candidate set is
            # stable enough for this to be a good first-order signal, and it
            # keeps iteration deterministic.
            def _run_key(p: Person, wp_code: str = wp.code) -> tuple[int, str]:
                last = _most_recent_wp_before(
                    assigned, p.name, project_months + 1
                )
                running = 1 if last == wp_code else 0
                # Python sorts ascending, and we want "running first", so
                # invert with negation (equivalent to `not running`).
                return (0 if running else 1, p.name)

            candidates = sorted(candidates_raw, key=_run_key)

            for p in candidates:
                if placed_eur >= wp.ceiling_eur:
                    break
                pw = windows_idx[(p.name, wp.code)]
                lo = max(wp.active_start, pw[0])
                hi = min(wp.active_end, pw[1])
                for m in range(lo, hi + 1):
                    if pool[p.name] <= 0:
                        break
                    if m in vacations[p.name]:
                        # Vacation month for this person — never placed.
                        continue
                    if m in frozen[p.name]:
                        # Frozen month — builder preserves the cell verbatim
                        # from the rebuilding workbook. Solver never places
                        # here. Frozen consumption is pre-subtracted from
                        # ceiling_eur / pool by the caller.
                        continue
                    if assigned[p.name][m] is None:
                        # Two-sided buffer policy:
                        #   upper cap = ceiling × (1 + MAX_BUFFER_PCT)   (+12%)
                        #   lower cap = ceiling × (1 − MIN_BUFFER_PCT)   (-11%)
                        # Default: skip placements that would push placed_eur
                        # past the upper cap. ``placed_eur`` only grows, so
                        # once the guard fires for this person at this rate
                        # it fires for every remaining month — break.
                        # Tie-break: if skipping would leave the WP below the
                        # lower cap, place anyway. Overshoot beyond +12% is
                        # preferred over undershoot worse than -11%.
                        upper_limit = (
                            wp.ceiling_eur * (1 + MAX_BUFFER_PCT)
                        )
                        lower_limit = (
                            wp.ceiling_eur * (1 - MIN_BUFFER_PCT)
                        )
                        would_overshoot = (
                            placed_eur + p.rate_eur_per_pm > upper_limit
                        )
                        already_above_lower = placed_eur >= lower_limit
                        if would_overshoot and already_above_lower:
                            break
                        assigned[p.name][m] = wp.code
                        pool[p.name] -= 1
                        placed_eur += p.rate_eur_per_pm
                        # Check-after-placing: last placement may overshoot
                        # the ceiling, bounded above by MAX_BUFFER_PCT
                        # except when the lower-cap tie-break fired (in
                        # which case overshoot can exceed +12%).
                        # Break stops further placements for this person;
                        # outer break then exits the WP pass.
                        if placed_eur >= wp.ceiling_eur:
                            break

    return assigned


# ---------------------------------------------------------------------------
# Pass 2: switch-reduction local improvement
# ---------------------------------------------------------------------------


def switch_reduction_pass(
    assigned: Assignment,
    persons: list[Person],
    wps: list[WP],
    person_wp_windows: list[PersonWPWindow],
    project_months: int,
) -> Assignment:
    """Iteratively swap (p1, m, m+1) <-> (p2, m, m+1) pairs to reduce the
    total switch count. Preserves per-WP totals in € (only equal-rate
    persons swap, so moving a cell between them is €-neutral),
    per-person totals (each person keeps their count), and all windows.

    Bounded: stops when one full outer pass finds no improving swap. Returns
    a new Assignment dict; does not mutate the caller's object.
    """
    wps_by_code = {w.code: w for w in wps}
    windows_idx = _build_windows_index(person_wp_windows)
    persons_by_name = {p.name: p for p in persons}

    # Work on a deep-ish copy. Inner dicts are month->str|None and need
    # their own copy; the outer dict is rebuilt by comprehension.
    state: Assignment = {
        name: dict(row) for name, row in assigned.items()
    }
    person_names = sorted(state.keys())

    improved = True
    # Outer loop bound: at most (#switches) improving swaps are possible,
    # and each strictly decreases the count. So an upper bound equal to the
    # initial switch count is plenty. In practice the inner search exits
    # quickly once no improving swap remains.
    max_outer = _count_switches(state) + 1
    outer_iterations = 0
    while improved and outer_iterations < max_outer:
        improved = False
        outer_iterations += 1
        current_switches = _count_switches(state)

        found_swap = False
        for p1_name in person_names:
            if found_swap:
                break
            for m in range(1, project_months):  # m and m+1 both in [1, N]
                if found_swap:
                    break
                a1 = state[p1_name][m]
                a2 = state[p1_name][m + 1]
                if a1 == a2:
                    # No switch here → swapping cannot reduce at this column.
                    continue
                for p2_name in person_names:
                    if p2_name == p1_name:
                        continue
                    # Rate-equality guard: under €-accounting, only swaps
                    # between persons at the same rate are €-neutral on
                    # each WP. Cross-tier swaps would shift per-WP € totals
                    # and break greedy's "placed_eur ≥ ceiling_eur" result.
                    if (
                        persons_by_name[p1_name].rate_eur_per_pm
                        != persons_by_name[p2_name].rate_eur_per_pm
                    ):
                        continue
                    # Vacation guard: swap eligibility requires neither m nor
                    # m+1 to fall in either person's vacation set. For the
                    # uniform-July model this is symmetric; the per-person
                    # check is what keeps the contract stable once a future
                    # design ships per-(person, year) overrides.
                    vac1 = persons_by_name[p1_name].vacation_months
                    vac2 = persons_by_name[p2_name].vacation_months
                    if (
                        m in vac1 or m + 1 in vac1
                        or m in vac2 or m + 1 in vac2
                    ):
                        continue
                    # Frozen guard: no swap may touch a frozen cell on
                    # either side. Mirrors the vacation guard shape.
                    fr1 = persons_by_name[p1_name].frozen_months
                    fr2 = persons_by_name[p2_name].frozen_months
                    if (
                        m in fr1 or m + 1 in fr1
                        or m in fr2 or m + 1 in fr2
                    ):
                        continue
                    b1 = state[p2_name][m]
                    b2 = state[p2_name][m + 1]
                    # Identical pair → trivial swap, no change.
                    if (a1, a2) == (b1, b2):
                        continue
                    # Legality: each cell must fit its new owner's windows.
                    if not _cell_legal_for_person(
                        p1_name, b1, m, wps_by_code, windows_idx
                    ):
                        continue
                    if not _cell_legal_for_person(
                        p1_name, b2, m + 1, wps_by_code, windows_idx
                    ):
                        continue
                    if not _cell_legal_for_person(
                        p2_name, a1, m, wps_by_code, windows_idx
                    ):
                        continue
                    if not _cell_legal_for_person(
                        p2_name, a2, m + 1, wps_by_code, windows_idx
                    ):
                        continue

                    # Apply tentatively, measure switches, revert if not
                    # strictly better.
                    state[p1_name][m], state[p2_name][m] = b1, a1
                    state[p1_name][m + 1], state[p2_name][m + 1] = b2, a2
                    new_switches = _count_switches(state)
                    if new_switches < current_switches:
                        improved = True
                        found_swap = True
                        break
                    # Revert.
                    state[p1_name][m], state[p2_name][m] = a1, b1
                    state[p1_name][m + 1], state[p2_name][m + 1] = a2, b2

    return state


# ---------------------------------------------------------------------------
# Composed entry point
# ---------------------------------------------------------------------------


def solve(
    persons: list[Person],
    wps: list[WP],
    person_wp_windows: list[PersonWPWindow],
    project_months: int,
) -> Assignment:
    """Return ``assignment[person_name][month_1based] -> wp_code | None``.

    Deterministic. Pure. See module docstring for the full contract.
    """
    assigned = greedy_pass(persons, wps, person_wp_windows, project_months)
    assigned = switch_reduction_pass(
        assigned, persons, wps, person_wp_windows, project_months
    )
    return assigned


# ---------------------------------------------------------------------------
# Stop-reason classifier (post-pass UI metadata)
# ---------------------------------------------------------------------------


def _placed_eur_for_wp(
    assigned: Assignment,
    persons: list[Person],
    wp_code: str,
) -> int:
    rate_by_name = {p.name: p.rate_eur_per_pm for p in persons}
    total = 0
    for name, row in assigned.items():
        rate = rate_by_name.get(name, 0)
        for code in row.values():
            if code == wp_code:
                total += rate
    return total


def _remaining_pool(
    assigned: Assignment, person_name: str, person_pool: int
) -> int:
    row = assigned.get(person_name, {})
    used = sum(1 for v in row.values() if v is not None)
    return person_pool - used


def _legal_unplaced_months(
    person: Person,
    wp: WP,
    pw_window: tuple[int, int],
    assigned: Assignment,
) -> list[int]:
    lo = max(wp.active_start, pw_window[0])
    hi = min(wp.active_end, pw_window[1])
    row = assigned.get(person.name, {})
    return [
        m
        for m in range(lo, hi + 1)
        if m not in person.vacation_months
        and m not in person.frozen_months
        and row.get(m) is None
    ]


def classify_wp_stop_reasons(
    assigned: Assignment,
    persons: list[Person],
    wps: list[WP],
    person_wp_windows: list[PersonWPWindow],
    project_months: int,
) -> dict[str, str]:
    """Return ``{wp_code: reason_string}`` for every WP in ``wps``.

    Pure; classifies each WP's solver termination reason from the final
    ``Assignment`` plus the same inputs ``solve`` consumes. Reason strings
    are 1–2 words from a closed enum:

    - ``"No candidates"`` — no PersonWPWindow row references this WP
    - ``"No window"`` — rows exist but none overlaps the WP's active range
    - ``"-11% override"`` — placed beyond +12% cap because skipping would
      have left the WP below -11% (lower-cap tie-break)
    - ``"Filled"`` — ``placed_eur >= ceiling_eur``
    - ``"+12% cap"`` — under ceiling; an eligible candidate has remaining
      pool, a legal unplaced month, and a rate that would push the next
      placement past the upper cap
    - ``"Pool empty"`` — under ceiling; no eligible candidate has remaining
      pool
    - ``"Window full"`` — under ceiling; eligible candidates have pool but
      no legal unplaced months remain
    - ``"Underfilled"`` — catch-all (should be unreachable; surfaces as a
      signal of policy drift, not silently masked)

    Priority order: first match wins.
    """
    del project_months  # signature parity with ``solve``; not needed here.
    windows_idx = _build_windows_index(person_wp_windows)
    reasons: dict[str, str] = {}
    for wp in wps:
        rows_for_wp = [w for w in person_wp_windows if w.wp_code == wp.code]
        if not rows_for_wp:
            reasons[wp.code] = "No candidates"
            continue
        any_overlap = any(
            max(w.start_month, wp.active_start)
            <= min(w.end_month, wp.active_end)
            for w in rows_for_wp
        )
        if not any_overlap:
            reasons[wp.code] = "No window"
            continue
        placed_eur = _placed_eur_for_wp(assigned, persons, wp.code)
        upper_limit = wp.ceiling_eur * (1 + MAX_BUFFER_PCT)
        if placed_eur > upper_limit:
            reasons[wp.code] = "-11% override"
            continue
        if placed_eur >= wp.ceiling_eur:
            reasons[wp.code] = "Filled"
            continue
        eligible = [p for p in persons if (p.name, wp.code) in windows_idx]
        cap_blocked = False
        any_pool = False
        any_pool_with_legal = False
        for p in eligible:
            if _remaining_pool(assigned, p.name, p.pool) <= 0:
                continue
            any_pool = True
            legal = _legal_unplaced_months(
                p, wp, windows_idx[(p.name, wp.code)], assigned
            )
            if legal:
                any_pool_with_legal = True
                if placed_eur + p.rate_eur_per_pm > upper_limit:
                    cap_blocked = True
        if cap_blocked:
            reasons[wp.code] = "+12% cap"
        elif not any_pool:
            reasons[wp.code] = "Pool empty"
        elif not any_pool_with_legal:
            reasons[wp.code] = "Window full"
        else:
            reasons[wp.code] = "Underfilled"
    return reasons
