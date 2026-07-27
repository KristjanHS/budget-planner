"""Static guard: no cell formula may pass a multi-cell RANGE to ``ISNUMBER``.

Why this is static and not a recalc assertion
---------------------------------------------
Inside a non-CSE formula Excel gives ``ISNUMBER(range)`` **no array
context** — it collapses to a scalar, so ``MAX(... * ISNUMBER(range))``
returns 0 and every cell guarded by ``IF(idx=0,"",…)`` renders blank.
LibreOffice array-evaluates the same expression correctly, and
``scripts/recalc.py`` drives LibreOffice. **No recalc-based test in this
suite can see this defect class at all** — `test_actuals_posted.py`'s
sparse-posting recalc test was green while the shipped workbook was broken
in Excel. The only honest guard is therefore static: assert the banned
construct is absent from the emitted formulas.

Detection is by ARGUMENT SHAPE, not by wrapper. An earlier design keyed on
``SUMPRODUCT(`` and ``ISNUMBER(`` co-occurring; that check false-NEGATIVES
on a future ``MAX(…*ISNUMBER(range))`` written without a SUMPRODUCT
wrapper, which is reason enough to key on the argument instead. (Whether
it would also false-positive on the scalar ``IF(ISNUMBER(cell),…)`` idiom
`.claude/rules/openpyxl.md` recommends depends on whether such terms ever
appear under a SUMPRODUCT — as of this commit they do not: ``trends.py``
emits its ``IF(ISNUMBER(E12),…)`` terms ``+``-joined and divided by
``partner_ceiling``, with no SUMPRODUCT wrapper. So that risk is
hypothetical, not observed.) Scalar uses — ``ISNUMBER(B1)``,
``IF(ISNUMBER(B1),B1,C1)`` — are permitted by construction.

Conditional-formatting rules are stored separately from cell values in the
OOXML and are therefore skipped for free by ``iter_rows()``; the scalar
``ISNUMBER`` uses in ``trends.py``/``capacity.py``/``timelog_plan.py`` CF
rules are out of scope either way.

Diagnosis + Excel-computed evidence:
``2026-07-27-isnumber-array-context-excel-fix.md`` (docs/plans/).
"""

import re

import openpyxl
import pytest

from budget_planner.layout import SHEET_ACTUALS_HOURS

# ``ISNUMBER(`` whose first argument is a range reference (``C7:AT7``,
# ``$C$7:$AT$13``, optionally sheet-qualified). Deliberately does NOT match
# a single-cell argument.
ISNUMBER_RANGE_RE = re.compile(
    r"ISNUMBER\(\s*(?:'[^']+'|[A-Za-z0-9_\-]+)?!?\$?[A-Z]{1,3}\$?\d+:\$?[A-Z]{1,3}\$?\d+",
)

# ── Stage-B allowlist ────────────────────────────────────────────────────
# Actuals-Hours scans, deliberately NOT fixed yet (the user is not filling
# hours data at this stage). Held here so the gap is NAMED rather than
# silently passing: these cells evaluate to 0 in Excel today, which makes
# Capacity "Actual PM" read 0 — correct-by-accident while no hours are
# posted, wrong the moment any are.
#
# The allowlist pins the two exact formula SHAPES Stage B still emits, not
# merely the hosting sheet: a hosting-sheet allowlist would wave through
# any NEW ``ISNUMBER(range)`` added to Actuals-Hours, including one
# scanning an ``Actuals!`` range (which would silently re-open the EUR path
# fixed in f47297d). Both shapes below admit only UNQUALIFIED (same-sheet)
# range refs, so a cross-sheet argument can never match. Nothing here is
# keyed on a column letter or a cell count — both move with
# ``PROJECT_MONTHS`` and with the fixture.
#
# Stage B is literally "delete entries from KNOWN_CSE_GAPS until it is
# empty." See the plan doc's § Stage B.
_LOCAL_RANGE = r"\$[A-Z]{1,3}\$\d+:\$[A-Z]{1,3}\$\d+"
_ANCHOR = r"\$[A-Z]{1,3}\$1"
# One ``SUMPRODUCT(MAX(...))`` latest-filled term of the C1 auto-derive.
_C1_SCAN_TERM = rf"SUMPRODUCT\(MAX\(\(COLUMN\({_LOCAL_RANGE}\)-COLUMN\({_ANCHOR}\)\+1\)\*ISNUMBER\({_LOCAL_RANGE}\)\)\)"

STAGE_B_SHAPES = (
    # budget_planner/actuals_hours.py `_cum_hours_formula` — per-person
    # cumulative-hours latest-filled row scan (one per hour-grid row).
    re.compile(
        rf"^=IFERROR\(INDEX\({_LOCAL_RANGE},1,"
        rf"SUMPRODUCT\(MAX\("
        rf"\(COLUMN\({_LOCAL_RANGE}\)-COLUMN\({_ANCHOR}\)\+1<=hours_actuals_through_month\)"
        rf"\*\(COLUMN\({_LOCAL_RANGE}\)-COLUMN\({_ANCHOR}\)\+1\)"
        rf"\*ISNUMBER\({_LOCAL_RANGE}\)"
        rf"\)\)\),0\)$"
    ),
    # budget_planner/actuals_hours.py — ``Actuals-Hours!C1`` auto-derived
    # through-month: MAX over one scan term per person's programme-WP row block.
    re.compile(rf"^=MAX\({_C1_SCAN_TERM}(?:,{_C1_SCAN_TERM})*\)$"),
)

KNOWN_CSE_GAPS: dict[str, tuple[re.Pattern[str], ...]] = {
    SHEET_ACTUALS_HOURS: STAGE_B_SHAPES,
}


def _multi_cell_name_re(wb):
    """Regex matching ``ISNUMBER(<name>)`` for RANGE-valued defined names.

    ``ISNUMBER_RANGE_RE`` only sees literal ``A1:B2`` refs. The workbook
    also carries range-valued defined names (``wp_budget_with_indirect``,
    ``roster_lookup``, ``assign_persons``, ``tbl_GradeRates``) that the
    codebase already dereferences (``INDEX(wp_budget_with_indirect,…)``),
    so ``ISNUMBER(wp_budget_with_indirect)`` is the same defect wearing a
    different name. Scalar names (``current_month``, ``fte_rate``) stay
    legal — only names whose destination spans multiple cells are listed,
    which is why this is not a generic bare-identifier match.

    Returns ``None`` when the workbook defines no range-valued names.
    """
    names = set()
    scopes = [wb.defined_names] + [ws.defined_names for ws in wb.worksheets]
    for scope in scopes:
        for name, defn in scope.items():
            try:
                dests = list(defn.destinations)
            except Exception:  # pragma: no cover - malformed/formula-valued name
                continue
            if any(":" in coord for _sheet, coord in dests):
                names.add(name)
    if not names:
        return None
    return re.compile(r"ISNUMBER\(\s*(?:" + "|".join(re.escape(n) for n in sorted(names)) + r")\s*\)")


def _offending_cells(wb):
    """All (sheet, coord, formula) triples passing a range to ISNUMBER."""
    name_re = _multi_cell_name_re(wb)
    out = []
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                value = cell.value
                if not (isinstance(value, str) and value.startswith("=")):
                    continue
                if ISNUMBER_RANGE_RE.search(value) or (name_re is not None and name_re.search(value)):
                    out.append((ws.title, cell.coordinate, value))
    return out


def _is_allowlisted(sheet: str, formula: str) -> bool:
    """True only for a known Stage-B shape emitted on its own sheet."""
    return any(shape.match(formula) for shape in KNOWN_CSE_GAPS.get(sheet, ()))


@pytest.mark.parametrize("wb_fixture", ["demo_wb_path", "fresh_wb_path"])
def test_no_isnumber_over_range_in_built_workbook(request, wb_fixture):
    """No emitted formula may pass a multi-cell range to ISNUMBER, except
    the exact Stage-B shapes listed in KNOWN_CSE_GAPS (deferred).

    Parametrized over both builder paths: ``demo_wb_path`` (assignments +
    seeded data) and ``fresh_wb_path`` (no assignments — the branchy
    empty-Capacity path emits a different formula population).
    """
    wb = openpyxl.load_workbook(request.getfixturevalue(wb_fixture), data_only=False)
    offenders = [o for o in _offending_cells(wb) if not _is_allowlisted(o[0], o[2])]
    assert not offenders, 'ISNUMBER(range) has no array context in Excel; use (range<>""):\n' + "\n".join(
        f"  {sheet}!{coord}: {formula}" for sheet, coord, formula in offenders
    )


def test_known_cse_gaps_are_not_stale(demo_wb_path):
    """Every allowlisted SHAPE must still match at least one emitted
    formula on its sheet. Without this, a Stage-B fix would leave a stale
    allowlist entry that silently re-opens the hole for a future regression
    on that sheet. ``demo_wb_path`` is the fixture where both shapes appear
    (the no-assignments ``fresh_wb_path`` emits only the per-person scan).
    """
    wb = openpyxl.load_workbook(demo_wb_path, data_only=False)
    offenders = _offending_cells(wb)
    stale = [
        f"{sheet} shape #{i}"
        for sheet, shapes in KNOWN_CSE_GAPS.items()
        for i, shape in enumerate(shapes)
        if not any(shape.match(f) for s, _, f in offenders if s == sheet)
    ]
    assert not stale, f"KNOWN_CSE_GAPS entries no longer needed — delete them: {stale}"
