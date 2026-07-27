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
``SUMPRODUCT(`` and ``ISNUMBER(`` co-occurring, which both false-positives
(``trends.py`` legitimately sums ``IF(ISNUMBER(E12),…)`` scalar terms inside
a SUMPRODUCT — the form ``.claude/rules/openpyxl.md`` recommends) and
false-negatives (a future ``MAX(…*ISNUMBER(range))`` written without a
SUMPRODUCT wrapper). Scalar uses — ``ISNUMBER(B1)``,
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

from budget_planner.layout import SHEET_ACTUALS_HOURS

# ``ISNUMBER(`` whose first argument is a range reference (``C7:AT7``,
# ``$C$7:$AT$13``, optionally sheet-qualified). Deliberately does NOT match
# a single-cell argument.
ISNUMBER_RANGE_RE = re.compile(
    r"ISNUMBER\(\s*(?:'[^']+'|[A-Za-z0-9_\-]+)?!?\$?[A-Z]{1,3}\$?\d+:\$?[A-Z]{1,3}\$?\d+",
)

# Stage B of the fix — Actuals-Hours scans, deliberately NOT fixed yet
# (the user is not filling hours data at this stage). Held here so the gap
# is NAMED rather than silently passing: these cells evaluate to 0 in Excel
# today, which makes Capacity "Actual PM" read 0 — correct-by-accident while
# no hours are posted, wrong the moment any are.
#
# Stage B is literally "delete entries from KNOWN_CSE_GAPS until it is
# empty." See the plan doc's § Stage B.
KNOWN_CSE_GAPS = {
    SHEET_ACTUALS_HOURS,
}


def _offending_cells(wb):
    """All (sheet, coord, formula) triples passing a range to ISNUMBER."""
    out = []
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                value = cell.value
                if isinstance(value, str) and value.startswith("=") and ISNUMBER_RANGE_RE.search(value):
                    out.append((ws.title, cell.coordinate, value))
    return out


def test_no_isnumber_over_range_in_built_workbook(demo_wb_path):
    """No emitted formula may pass a multi-cell range to ISNUMBER, except on
    the sheets named in KNOWN_CSE_GAPS (Stage B, deferred)."""
    wb = openpyxl.load_workbook(demo_wb_path, data_only=False)
    offenders = [o for o in _offending_cells(wb) if o[0] not in KNOWN_CSE_GAPS]
    assert not offenders, 'ISNUMBER(range) has no array context in Excel; use (range<>""):\n' + "\n".join(
        f"  {sheet}!{coord}: {formula}" for sheet, coord, formula in offenders
    )


def test_known_cse_gaps_are_not_stale(demo_wb_path):
    """Every sheet allowlisted in KNOWN_CSE_GAPS must still contain at least
    one offending formula. Without this, a Stage-B fix would leave a stale
    allowlist entry that silently re-opens the hole for a future regression
    on that sheet.
    """
    wb = openpyxl.load_workbook(demo_wb_path, data_only=False)
    offending_sheets = {sheet for sheet, _, _ in _offending_cells(wb)}
    stale = KNOWN_CSE_GAPS - offending_sheets
    assert not stale, f"KNOWN_CSE_GAPS entries no longer needed — delete them: {sorted(stale)}"
