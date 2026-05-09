"""Actuals Breach In posted-cumulative-slope + Posted-through header tests.

Post 2026-04-18 Config-through-month cleanup:
- Breach In formula inlines ``MAX(1, eur_actuals_through_month)``
  (derived ``months_elapsed_*`` named ranges were dropped).
- Row 1 holds the through-month override strip (label / override / auto /
  effective). The "Posted through:" header row shifted from row 1 to
  row 2.

Posted-cumulative redesign (2026-05-09):
- Per-WP Breach -i = `(Budget - Posted) / (Posted / MAX(1, eur_actuals_through_month))`
  guarded by `IF(posted_idx=0, "", IF(Posted=0, "—", ...))`. The forecast
  monitoring block was deleted; hours-side breach lives on Actuals-Hours
  now (Phase 3 — Stage D). The Actuals-Hours side header still reads
  "As of:" until Stage D ships.
"""
import openpyxl
import pytest

from budget_planner.config import (
    VAL_COL_BREACH,
)


@pytest.fixture(scope="module")
def actuals_ws(fresh_wb_path):
    wb = openpyxl.load_workbook(fresh_wb_path, data_only=False)
    return wb["Actuals"]


def _find_first_wp_row(ws):
    for r in range(1, ws.max_row + 1):
        val = ws.cell(r, 1).value
        if isinstance(val, str) and val.startswith("=Config!A"):
            return r
    raise AssertionError("No WP row found in Actuals")


def _find_total_row(ws):
    for r in range(1, ws.max_row + 1):
        if ws.cell(r, 2).value == "TOTAL":
            return r
    raise AssertionError("No TOTAL row found in Actuals")


def _find_month_header_row(ws):
    """Row whose col A cell is the literal 'WP ID' header."""
    for r in range(1, ws.max_row + 1):
        if ws.cell(r, 1).value == "WP ID":
            return r
    raise AssertionError("No 'WP ID' header row found in Actuals")


# ── Breach In formulas: cumulative-slope denominators ──


def test_actual_wp_breach_uses_eur_actuals_through_month(actuals_ws):
    """Per-WP Actual Breach In inlines MAX(1,eur_actuals_through_month)."""
    row = _find_first_wp_row(actuals_ws)
    formula = str(actuals_ws.cell(row, VAL_COL_BREACH).value)
    assert "MAX(1,eur_actuals_through_month)" in formula, formula
    assert "current_month" not in formula, formula


def test_total_actual_breach_uses_eur_actuals_through_month(actuals_ws):
    """TOTAL row Actual Breach In inlines MAX(1,eur_actuals_through_month)."""
    row = _find_total_row(actuals_ws)
    formula = str(actuals_ws.cell(row, VAL_COL_BREACH).value)
    assert "MAX(1,eur_actuals_through_month)" in formula, formula
    assert "current_month" not in formula, formula


# ── "As of" header row ──


def test_posted_through_header_actual_block(actuals_ws):
    """Row 2 col B = 'Posted through:', col C = green
    EDATE(project_start, eur_actuals_through_month - 1).

    Row 1 holds the through-month override strip; the "Posted through:"
    header sits on row 2. Posted-cumulative redesign (2026-05-09)
    renamed the legacy "As of:" label to match the per-WP "Posted -i"
    column convention. Through-month cells are integer month indices;
    EDATE turns them into a calendar date so the header reads
    "Posted through: Apr 2026" rather than the Excel serial-date
    rendering of a small integer.
    """
    assert actuals_ws.cell(2, 2).value == "Posted through:"
    date_cell = actuals_ws.cell(2, 3)
    formula = str(date_cell.value)
    assert formula.startswith("=EDATE(project_start"), formula
    assert "eur_actuals_through_month" in formula, formula
    # Green cross-sheet-ref font (color 008000)
    rgb = getattr(getattr(date_cell.font, "color", None), "rgb", None)
    assert rgb is not None and str(rgb).upper().endswith("008000"), rgb


# ── Layout invariants after row-shift ──


def test_freeze_panes_matches_first_wp_row(actuals_ws):
    """Freeze panes column C; row equals first WP row (shifted down by 1 from pre-T7)."""
    import re
    fp = actuals_ws.freeze_panes
    assert fp is not None
    m = re.match(r"([A-Z]+)(\d+)", fp)
    assert m is not None, fp
    assert m.group(1) == "C"
    freeze_row = int(m.group(2))
    assert freeze_row == _find_first_wp_row(actuals_ws)


def test_row_ordering_is_consistent(actuals_ws):
    """Strip(1) < As-of(2) < section header(3) < year/month rows < first_wp < total."""
    month_header = _find_month_header_row(actuals_ws)
    first_wp = _find_first_wp_row(actuals_ws)
    total = _find_total_row(actuals_ws)
    # Row 1 is the through-month override strip (A1 label).
    assert actuals_ws.cell(1, 1).value == "EUR Actuals Through Month:"
    # "Posted through:" header row (renamed from "As of:" in the
    # posted-cumulative redesign 2026-05-09).
    assert actuals_ws.cell(2, 2).value == "Posted through:"
    # Section header split by T3 (indirect-view restructure 2026-04-22)
    # into "ACTUAL EUR -i" (cols 1..breach) + "ACTUAL EUR +i" (cols
    # sum_pi..breach_pi) merged regions on row 3; col 1 anchors the -i
    # region.
    assert actuals_ws.cell(3, 1).value == "ACTUAL EUR -i"
    # Year row sits between section header and month header
    assert 3 < month_header - 1 < month_header
    assert month_header + 1 == first_wp
    assert first_wp < total


# ── Through-month control strip (Row 1 on Actuals + Actuals-Hours) ──


@pytest.fixture(scope="module")
def actuals_hours_ws_demo(demo_wb_path):
    """Loaded from demo workbook so per-person programme-WP rows exist — without
    assignments the auto-derive falls back to `=0`, which masks the
    SUMPRODUCT scan this test is meant to pin down."""
    wb = openpyxl.load_workbook(demo_wb_path, data_only=False)
    return wb["Actuals-Hours"]


def test_through_month_strip_formulas_on_actuals(actuals_ws):
    """C1 auto-derive is a SUMPRODUCT(MAX(...ISNUMBER...)) scan over the WP
    data block; D1 effective is IF(ISNUMBER(B1),B1,C1). Bug surface: a
    refactor that narrows the scan range (e.g. includes the TOTAL row) or
    inverts the IF precedence silently flips `eur_actuals_through_month`
    to the wrong integer without breaking recalc.
    """
    auto = str(actuals_ws.cell(1, 3).value)
    assert auto.startswith("=SUMPRODUCT("), auto
    assert "MAX(" in auto and "ISNUMBER(" in auto, auto
    assert actuals_ws.cell(1, 4).value == "=IF(ISNUMBER(B1),B1,C1)"


def test_through_month_strip_formulas_on_actuals_hours(actuals_hours_ws_demo):
    """Same strip contract as Actuals (label / override / auto / effective)
    registered against `hours_actuals_through_month`. The auto scan here
    is `=MAX(SUMPRODUCT(...), ...)` over per-person programme-WP row ranges —
    externals and per-person TOTAL rows are excluded.

    The demo fixture has multiple team members with programme-WP assignments,
    so the formula must contain ≥2 `SUMPRODUCT(` terms. A regression that
    collapses the scan to a single sheet-wide range (over-counting over
    subtotal + TOTAL rows) would leave exactly one term and fail this
    check. Covers IF-precedence inversion too (D1 formula asserted
    verbatim).
    """
    assert actuals_hours_ws_demo.cell(1, 1).value == "Hours Actuals Through Month:"
    auto = str(actuals_hours_ws_demo.cell(1, 3).value)
    assert auto.startswith("=MAX(SUMPRODUCT("), auto
    assert "ISNUMBER(" in auto, auto
    assert auto.count("SUMPRODUCT(") >= 2, (
        f"expected per-person SUMPRODUCT terms, got: {auto}"
    )
    assert actuals_hours_ws_demo.cell(1, 4).value == "=IF(ISNUMBER(B1),B1,C1)"
