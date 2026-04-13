"""Summary sheet — Task 10 Forecast EUR column tests.

The Summary sheet has 5 columns per calendar-year group (and per project
total): Plan FTE, Plan EUR, Act EUR, Fcst EUR, Var EUR. The Fcst EUR column
was added in Task 10 and sources values from the Plan-Forecast sheet.
"""

import pytest
import openpyxl

from budget_planner.workbook import init_workbook
from budget_planner.layout import (
    MONTH_COL_START, _calendar_year_groups,
)
from budget_planner.config import (
    NUM_WPS, PROJECT_MONTHS, PROJECT_START,
    DEMO_CAPACITY_ASSIGNMENTS,
)
from budget_planner.styles import col_letter


_YEAR_GROUP_START = 3
_COLS_PER_YEAR = 5  # Plan FTE, Plan EUR, Act EUR, Fcst EUR, Var EUR
_CAL_GROUPS = list(_calendar_year_groups(PROJECT_START, PROJECT_MONTHS))
_NUM_YEARS = len(_CAL_GROUPS)
_TOTAL_GROUP_START = _YEAR_GROUP_START + _NUM_YEARS * _COLS_PER_YEAR


@pytest.fixture(scope="module")
def summary_wb(tmp_path_factory):
    out = tmp_path_factory.mktemp("summary") / "test.xlsx"
    init_workbook(str(out), assignments=DEMO_CAPACITY_ASSIGNMENTS)
    return openpyxl.load_workbook(str(out), data_only=False)


def _find_summary_first_wp_row(ws):
    for r in range(1, ws.max_row + 1):
        val = ws.cell(r, 1).value
        if isinstance(val, str) and val.startswith("=Config!A"):
            return r
    raise AssertionError("Cannot find Summary first WP row")


def _find_plan_forecast_first_wp_row(ws):
    for r in range(1, ws.max_row + 1):
        val = ws.cell(r, 1).value
        if isinstance(val, str) and val.startswith("=Config!A"):
            return r
    raise AssertionError("Cannot find Plan-Forecast first WP row")


# ── Structure ──

def test_summary_has_five_cols_per_year(summary_wb):
    """Sub-headers row: each year group has 5 cols ending in 'Var €'."""
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
        labels = [ws.cell(sub_row, base + i).value for i in range(5)]
        # Exact match, including the new Fcst column (Task 10)
        assert labels == ["Plan FTE", "Plan \u20ac", "Act \u20ac",
                          "Fcst \u20ac", "Var \u20ac"], (
            f"Year {y}: unexpected sub-header labels {labels}"
        )


def test_summary_forecast_eur_formulas_reference_plan_forecast(summary_wb):
    """Each WP × year Fcst EUR cell = SUM over Plan-Forecast monthly range."""
    ws = summary_wb["Summary"]
    pf = summary_wb["Plan-Forecast"]
    first_wp = _find_summary_first_wp_row(ws)
    pf_first_wp = _find_plan_forecast_first_wp_row(pf)

    for i in range(NUM_WPS):
        row = first_wp + i
        pf_row = pf_first_wp + i
        for y, (group_start, span) in enumerate(_CAL_GROUPS):
            sc = _YEAR_GROUP_START + y * _COLS_PER_YEAR
            # Fcst EUR is at sc + 3
            cell = ws.cell(row, sc + 3)
            first_col_l = col_letter(MONTH_COL_START + group_start)
            last_col_l = col_letter(MONTH_COL_START + group_start + span - 1)
            expected = (
                f"=SUM('Plan-Forecast'!{first_col_l}{pf_row}:"
                f"{last_col_l}{pf_row})"
            )
            # Excel quotes sheet names with hyphens; accept both quoted/unquoted
            expected_unquoted = (
                f"=SUM(Plan-Forecast!{first_col_l}{pf_row}:"
                f"{last_col_l}{pf_row})"
            )
            assert cell.value in (expected, expected_unquoted), (
                f"WP{i} year {y} Fcst EUR: got {cell.value!r}"
            )


def test_summary_forecast_eur_is_green_cross_sheet(summary_wb):
    """Fcst EUR cells use GREEN_XREF font color (008000)."""
    ws = summary_wb["Summary"]
    first_wp = _find_summary_first_wp_row(ws)
    # Sample first WP × first year
    sc = _YEAR_GROUP_START + 0 * _COLS_PER_YEAR
    cell = ws.cell(first_wp, sc + 3)
    color = cell.font.color
    assert color is not None
    rgb = color.rgb
    # Green convention: 008000 (style may add alpha prefix FF)
    assert rgb is not None and str(rgb).endswith("008000"), (
        f"Expected green font (008000), got {rgb!r}"
    )


def test_summary_project_total_forecast_eur_sums_years(summary_wb):
    """Project-total Fcst EUR = sum of per-year Fcst EUR columns."""
    ws = summary_wb["Summary"]
    first_wp = _find_summary_first_wp_row(ws)
    for i in range(NUM_WPS):
        row = first_wp + i
        fcst_year_cols = [
            col_letter(_YEAR_GROUP_START + y * _COLS_PER_YEAR + 3)
            for y in range(_NUM_YEARS)
        ]
        expected = "=" + "+".join(f"{cl}{row}" for cl in fcst_year_cols)
        # tc + 3 is the project-total Fcst EUR column
        cell_val = ws.cell(row, _TOTAL_GROUP_START + 3).value
        assert cell_val == expected, (
            f"Row {row} project-total Fcst EUR: expected {expected}, "
            f"got {cell_val}"
        )
