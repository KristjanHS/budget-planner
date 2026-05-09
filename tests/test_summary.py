"""Summary sheet — project-end Forecast EUR column tests.

Phase 2.3 (2026-05-09) retired the per-year Forecast EUR column. The Summary
sheet now has 4 columns per calendar-year group (Plan FTE, Plan EUR, Act EUR,
Var EUR) PLUS a single project-end Forecast EUR column sitting between the
year band and PROJECT TOTAL. The Forecast EUR cell projects from the Actuals
posted-cumulative cells (posted_idx + posted_value + slope) out to
project_months — no Plan-Forecast tab dependency anymore.
"""

import pytest
import openpyxl

from budget_planner.layout import (
    _calendar_year_groups, val_columns,
)
from budget_planner.config import (
    NUM_WPS, PROJECT_MONTHS, PROJECT_START,
)
from budget_planner.styles import col_letter


_YEAR_GROUP_START = 3
_COLS_PER_YEAR = 4  # Plan FTE, Plan EUR, Act EUR, Var EUR
_CAL_GROUPS = list(_calendar_year_groups(PROJECT_START, PROJECT_MONTHS))
_NUM_YEARS = len(_CAL_GROUPS)
# Project-end Forecast EUR column sits between the year band and PROJECT TOTAL.
_FORECAST_COL = _YEAR_GROUP_START + _NUM_YEARS * _COLS_PER_YEAR
_TOTAL_GROUP_START = _FORECAST_COL + 1


@pytest.fixture(scope="module")
def summary_wb(assigned_wb_path):
    return openpyxl.load_workbook(assigned_wb_path, data_only=False)


def _find_summary_first_wp_row(ws):
    for r in range(1, ws.max_row + 1):
        val = ws.cell(r, 1).value
        if isinstance(val, str) and val.startswith("=Config!A"):
            return r
    raise AssertionError("Cannot find Summary first WP row")


def _find_actuals_first_wp_row(ws):
    for r in range(1, ws.max_row + 1):
        val = ws.cell(r, 1).value
        if isinstance(val, str) and val.startswith("=Config!A"):
            return r
    raise AssertionError("Cannot find Actuals first WP row")


# Structure

def test_summary_has_four_cols_per_year(summary_wb):
    """Sub-headers row: each year group has 4 cols ending in 'Var €'."""
    ws = summary_wb["Summary"]
    # Find sub-header row by scanning for "Plan FTE" label in year col
    sub_row = None
    for r in range(1, 10):
        if ws.cell(r, _YEAR_GROUP_START).value == "Plan FTE":
            sub_row = r
            break
    assert sub_row is not None

    for y in range(_NUM_YEARS):
        base = _YEAR_GROUP_START + y * _COLS_PER_YEAR
        labels = [ws.cell(sub_row, base + i).value for i in range(4)]
        # Phase 2.3: per-year Fcst € slot retired.
        assert labels == ["Plan FTE", "Plan €", "Act €",
                          "Var €"], (
            f"Year {y}: unexpected sub-header labels {labels}"
        )


def test_summary_forecast_eur_references_actuals_posted_cells(summary_wb):
    """Project-end Forecast EUR formula references Actuals posted_idx /
    posted_value / slope cells and the project_months named range."""
    ws = summary_wb["Summary"]
    actuals_ws = summary_wb["Actuals"]
    first_wp = _find_summary_first_wp_row(ws)
    actuals_first_wp = _find_actuals_first_wp_row(actuals_ws)

    val_cols = val_columns(PROJECT_MONTHS)
    posted_idx_l = col_letter(val_cols["posted_idx"])
    posted_value_l = col_letter(val_cols["posted_value"])
    slope_l = col_letter(val_cols["slope"])

    # Verify the Forecast EUR sub-header is present where we expect.
    sub_row = None
    for r in range(1, 10):
        if ws.cell(r, _YEAR_GROUP_START).value == "Plan FTE":
            sub_row = r
            break
    assert sub_row is not None
    assert ws.cell(sub_row, _FORECAST_COL).value == "Forecast €", (
        f"Expected 'Forecast €' header at col {_FORECAST_COL}, got "
        f"{ws.cell(sub_row, _FORECAST_COL).value!r}"
    )

    for i in range(NUM_WPS):
        row = first_wp + i
        actuals_row = actuals_first_wp + i
        cell = ws.cell(row, _FORECAST_COL)
        formula = cell.value
        assert isinstance(formula, str), (
            f"WP{i}: Forecast EUR cell is not a formula: {formula!r}"
        )
        # Must reference all three Actuals cells for this WP.
        assert f"Actuals!{posted_idx_l}{actuals_row}" in formula, (
            f"WP{i} forecast formula missing posted_idx ref "
            f"Actuals!{posted_idx_l}{actuals_row}: {formula}"
        )
        assert f"Actuals!{posted_value_l}{actuals_row}" in formula, (
            f"WP{i} forecast formula missing posted_value ref: {formula}"
        )
        assert f"Actuals!{slope_l}{actuals_row}" in formula, (
            f"WP{i} forecast formula missing slope ref: {formula}"
        )
        assert "project_months" in formula, (
            f"WP{i} forecast formula missing project_months named range: "
            f"{formula}"
        )
        # Empty-sentinel guard: =0 then "" branch.
        assert "=0," in formula and '""' in formula, (
            f"WP{i} forecast formula missing IF posted_idx=0 empty sentinel: "
            f"{formula}"
        )


def test_summary_forecast_eur_is_green_cross_sheet(summary_wb):
    """Project-end Forecast EUR cells use GREEN_XREF font color (008000)."""
    ws = summary_wb["Summary"]
    first_wp = _find_summary_first_wp_row(ws)
    cell = ws.cell(first_wp, _FORECAST_COL)
    color = cell.font.color
    assert color is not None
    rgb = color.rgb
    assert rgb is not None and str(rgb).endswith("008000"), (
        f"Expected green font (008000) on project-end Forecast EUR, "
        f"got {rgb!r}"
    )
