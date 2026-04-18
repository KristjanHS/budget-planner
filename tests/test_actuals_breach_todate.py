"""T7: Actuals Breach In cumulative-slope + As-of header tests.

Covers the Stage 2 / T7 changes in ``budget_planner/actuals.py``:
- Actual + Forecast Breach In formulas divide cum by
  ``months_elapsed_eur`` / ``months_elapsed_hours`` rather than the
  raw project-month ``current_month``.
- Two green ``=..._actuals_through_month`` "As of" header cells sit above
  the Actual-EUR and Forecast Monitoring blocks on row 1, with the
  section headers + year/month/WP rows all shifted down by one.
"""
import openpyxl
import pytest

from budget_planner.config import (
    VAL_COL_BREACH,
)
from budget_planner.layout import forecast_val_columns


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


def test_actual_wp_breach_uses_months_elapsed_eur(actuals_ws):
    """Per-WP Actual Breach In divides by months_elapsed_eur, not current_month."""
    row = _find_first_wp_row(actuals_ws)
    formula = str(actuals_ws.cell(row, VAL_COL_BREACH).value)
    assert "months_elapsed_eur" in formula, formula
    assert "current_month" not in formula, formula


def test_forecast_wp_breach_uses_months_elapsed_hours(actuals_ws):
    """Per-WP Forecast Breach In divides by months_elapsed_hours, not current_month."""
    from budget_planner.config import PROJECT_MONTHS
    fcols = forecast_val_columns(PROJECT_MONTHS)
    row = _find_first_wp_row(actuals_ws)
    formula = str(actuals_ws.cell(row, fcols["breach_forecast"]).value)
    assert "months_elapsed_hours" in formula, formula
    assert "current_month" not in formula, formula


def test_total_actual_breach_uses_months_elapsed_eur(actuals_ws):
    """TOTAL row Actual Breach In divides by months_elapsed_eur."""
    row = _find_total_row(actuals_ws)
    formula = str(actuals_ws.cell(row, VAL_COL_BREACH).value)
    assert "months_elapsed_eur" in formula, formula
    assert "current_month" not in formula, formula


def test_total_forecast_breach_uses_months_elapsed_hours(actuals_ws):
    """TOTAL row Forecast Breach In divides by months_elapsed_hours."""
    from budget_planner.config import PROJECT_MONTHS
    fcols = forecast_val_columns(PROJECT_MONTHS)
    row = _find_total_row(actuals_ws)
    formula = str(actuals_ws.cell(row, fcols["breach_forecast"]).value)
    assert "months_elapsed_hours" in formula, formula
    assert "current_month" not in formula, formula


# ── "As of" header row ──


def test_asof_header_actual_block(actuals_ws):
    """Row 1 col B = 'As of:', col C = green EDATE(project_start, eur_actuals_through_month - 1).

    Through-month cells are integer month indices; EDATE turns them into a
    calendar date so the header reads "As of: Apr 2026" rather than the
    Excel serial-date rendering of a small integer.
    """
    assert actuals_ws.cell(1, 2).value == "As of:"
    date_cell = actuals_ws.cell(1, 3)
    formula = str(date_cell.value)
    assert formula.startswith("=EDATE(project_start"), formula
    assert "eur_actuals_through_month" in formula, formula
    # Green cross-sheet-ref font (color 008000)
    rgb = getattr(getattr(date_cell.font, "color", None), "rgb", None)
    assert rgb is not None and str(rgb).upper().endswith("008000"), rgb


def test_asof_header_forecast_block(actuals_ws):
    """Row 1 at col_cum_forecast = 'As of:', next col = green EDATE hours formula."""
    from budget_planner.config import PROJECT_MONTHS
    fcols = forecast_val_columns(PROJECT_MONTHS)
    col_cum_forecast = fcols["cum_forecast"]
    assert actuals_ws.cell(1, col_cum_forecast).value == "As of:"
    date_cell = actuals_ws.cell(1, col_cum_forecast + 1)
    formula = str(date_cell.value)
    assert formula.startswith("=EDATE(project_start"), formula
    assert "hours_actuals_through_month" in formula, formula
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
    """As-of row(1) < section header(2) < year/month rows < first_wp < total."""
    month_header = _find_month_header_row(actuals_ws)
    first_wp = _find_first_wp_row(actuals_ws)
    total = _find_total_row(actuals_ws)
    # Section header "ACTUAL EUR" now on row 2 (shifted down from row 1)
    assert actuals_ws.cell(2, 1).value == "ACTUAL EUR"
    # Year row sits between section header and month header
    assert 2 < month_header - 1 < month_header
    assert month_header + 1 == first_wp
    assert first_wp < total
