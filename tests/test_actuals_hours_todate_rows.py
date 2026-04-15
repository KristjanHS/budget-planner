"""Tests for the per-person hours-to-date summary rows on Actuals-Hours.

Two new rows sit between externals and TOTAL for each person:

* "Cumulative hours to date" — SUM of each data row's cum-hours helper cell.
* "Actual PM to date" — ``cum_hours / hours_per_fte_month``.

These are cosmetic aggregates of the existing per-row cum-hours helper —
they let auditors see the project-to-date burn without scanning every
row. TOTAL still sums only the data rows (not the summaries).
"""

import re

import pytest
from openpyxl import load_workbook

from budget_planner.workbook import init_workbook
from budget_planner.config import DEMO_CAPACITY_ASSIGNMENTS, TEAM, PROJECT_MONTHS
from budget_planner.layout import val_columns
from budget_planner.styles import col_letter


FORECAST_COL_START = val_columns(PROJECT_MONTHS)["sum"] + 2
FORECAST_COL_END = FORECAST_COL_START + PROJECT_MONTHS - 1
CUM_HOURS_COL = FORECAST_COL_END + 2


@pytest.fixture(scope="module")
def wb_path(tmp_path_factory):
    out = tmp_path_factory.mktemp("ah_todate") / "wb.xlsx"
    init_workbook(str(out), assignments=DEMO_CAPACITY_ASSIGNMENTS)
    return str(out)


@pytest.fixture(scope="module")
def hrs(wb_path):
    wb = load_workbook(wb_path, data_only=False)
    return wb["Actuals-Hours"]


def _find_person_header_row(ws, person_name):
    merged_ranges = [(mr.min_row, mr.min_col, mr.max_col)
                     for mr in ws.merged_cells.ranges]
    for r in range(1, ws.max_row + 1):
        if ws.cell(r, 1).value != person_name:
            continue
        if any(mr_r == r and mr_c1 == 1 and mr_c2 >= 2
               for mr_r, mr_c1, mr_c2 in merged_ranges):
            return r
    raise AssertionError(f"Person header for {person_name!r} not found")


def _find_summary_rows(ws, header_row):
    """Return (cum_hours_row, actual_pm_row) under a person header by
    scanning for the label in col B."""
    cum_row = None
    pm_row = None
    for r in range(header_row + 1, header_row + 30):
        label = ws.cell(r, 2).value
        if label == "Cumulative hours to date":
            cum_row = r
        elif label == "Actual PM to date":
            pm_row = r
            break
    assert cum_row is not None
    assert pm_row is not None
    return cum_row, pm_row


def _find_data_rows(ws, header_row, cum_row):
    """Data rows (WP + external) sit strictly between header and cum row."""
    return list(range(header_row + 1, cum_row))


def test_summary_rows_present_for_every_person(hrs):
    for member in TEAM:
        header = _find_person_header_row(hrs, member["name"])
        cum_row, pm_row = _find_summary_rows(hrs, header)
        assert pm_row == cum_row + 1


def test_summary_row_labels_exact(hrs):
    """Col B labels must be the exact design-doc strings — the reader
    relies on them (non-TOTAL, non-WP) and auditors read them on-screen."""
    header = _find_person_header_row(hrs, "Alice")
    cum_row, pm_row = _find_summary_rows(hrs, header)
    assert hrs.cell(cum_row, 2).value == "Cumulative hours to date"
    assert hrs.cell(pm_row, 2).value == "Actual PM to date"
    # Col A stays blank so the rebuild reader skips these rows.
    assert hrs.cell(cum_row, 1).value in (None, "")
    assert hrs.cell(pm_row, 1).value in (None, "")


def test_cum_hours_todate_formula_spans_data_rows(hrs):
    """Cum-hours-to-date cell = SUM over the per-row cum-hours helper
    column for this person's data rows only (no other persons, no TOTAL)."""
    header = _find_person_header_row(hrs, "Alice")
    cum_row, _ = _find_summary_rows(hrs, header)
    data_rows = _find_data_rows(hrs, header, cum_row)
    assert data_rows, "Alice must have at least one data row"

    formula = hrs.cell(cum_row, CUM_HOURS_COL).value
    assert isinstance(formula, str)
    cl = col_letter(CUM_HOURS_COL)
    expected = f"=SUM({cl}{data_rows[0]}:{cl}{data_rows[-1]})"
    assert formula == expected, (
        f"Expected {expected!r}, got {formula!r}"
    )


def test_actual_pm_todate_divides_by_hours_per_fte_month(hrs):
    header = _find_person_header_row(hrs, "Alice")
    cum_row, pm_row = _find_summary_rows(hrs, header)
    formula = hrs.cell(pm_row, CUM_HOURS_COL).value
    assert isinstance(formula, str)
    cl = col_letter(CUM_HOURS_COL)
    assert formula == f"={cl}{cum_row}/hours_per_fte_month", formula


def test_summary_rows_outline_level_one(hrs):
    """Summary rows collapse with the rest of the person block."""
    for member in TEAM:
        header = _find_person_header_row(hrs, member["name"])
        cum_row, pm_row = _find_summary_rows(hrs, header)
        for r in (cum_row, pm_row):
            level = hrs.row_dimensions[r].outline_level or 0
            assert level == 1, (
                f"{member['name']}: summary row {r} outline_level={level} "
                f"(expected 1)"
            )


def test_total_row_sums_data_rows_not_summary_rows(hrs):
    """TOTAL row per-month SUM spans the data rows above the summary pair,
    not the summary rows themselves — otherwise the summary's aggregate
    would double-count into TOTAL."""
    header = _find_person_header_row(hrs, "Alice")
    cum_row, pm_row = _find_summary_rows(hrs, header)
    # TOTAL row is directly after the Actual PM row.
    total_row = pm_row + 1
    assert hrs.cell(total_row, 2).value == "TOTAL"
    data_rows = _find_data_rows(hrs, header, cum_row)
    data_start, data_end = data_rows[0], data_rows[-1]

    # M1 col is col C (MONTH_COL_START = 3).
    from budget_planner.layout import MONTH_COL_START
    m1_formula = hrs.cell(total_row, MONTH_COL_START).value
    assert isinstance(m1_formula, str)
    m = re.match(r"=SUM\(([A-Z]+)(\d+):([A-Z]+)(\d+)\)", m1_formula)
    assert m, f"TOTAL M1 formula not a simple SUM range: {m1_formula!r}"
    assert int(m.group(2)) == data_start
    assert int(m.group(4)) == data_end
    # Sanity: summary rows sit strictly between data_end and total_row.
    assert data_end < cum_row < pm_row < total_row
