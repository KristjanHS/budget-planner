"""Tests for the Trends sheet builder."""

import pytest
import openpyxl

from budget_planner.workbook import init_workbook, seed_demo_data
from budget_planner.config import (
    NUM_WPS, TEAM,
    VAL_COL_BUDGET,
)
from budget_planner.styles import col_letter


# ── Fixtures ──

@pytest.fixture(scope="module")
def trends_workbook(tmp_path_factory):
    out = tmp_path_factory.mktemp("trends") / "test.xlsx"
    init_workbook(str(out))
    seed_demo_data(str(out))
    return openpyxl.load_workbook(str(out), data_only=False)


@pytest.fixture(scope="module")
def trends_ws(trends_workbook):
    return trends_workbook["Trends"]


def _find_row(ws, col, value):
    """Find the first row where ws.cell(row, col).value == value."""
    for r in range(1, ws.max_row + 1):
        v = ws.cell(r, col).value
        if v == value:
            return r
    return None


def _find_row_containing(ws, col, substring):
    """Find the first row where substring appears in str(ws.cell(row, col).value)."""
    for r in range(1, ws.max_row + 1):
        v = ws.cell(r, col).value
        if v is not None and substring in str(v):
            return r
    return None


# ================================================================
# Sheet basics
# ================================================================

def test_trends_sheet_exists(trends_workbook):
    assert "Trends" in trends_workbook.sheetnames


def test_trends_tab_color(trends_ws):
    tc = trends_ws.sheet_properties.tabColor
    assert tc is not None
    assert "FFD700" in str(tc.rgb)


def test_current_month_named_range(trends_workbook):
    defined = {}
    for name in trends_workbook.defined_names:
        dn = trends_workbook.defined_names[name]
        defined[dn.name] = dn.attr_text
    assert "current_month" in defined
    assert defined["current_month"] == "Trends!$B$1"


def test_current_month_default_value(trends_ws):
    """B1 has auto-calculating formula referencing project_start and TODAY()."""
    val = trends_ws["B1"].value
    assert isinstance(val, str) and val.startswith("="), (
        f"Expected formula in Trends!B1, got {val}"
    )
    assert "TODAY()" in val
    assert "project_start" in val


# ================================================================
# WP Burn Rate table
# ================================================================

def test_burn_rate_section_header(trends_ws):
    """Row 3 has the WP BURN RATE & EXHAUSTION section header."""
    val = trends_ws.cell(3, 1).value
    assert val is not None
    assert "WP BURN RATE" in str(val)


def test_burn_rate_wp_count(trends_ws):
    """8 WP rows between header row and TOTAL row."""
    total_row = _find_row(trends_ws, 2, "TOTAL")
    assert total_row is not None
    # Header row is 4, first WP is 5
    header_row = _find_row_containing(trends_ws, 1, "WP ID")
    assert header_row is not None
    first_wp_row = header_row + 1
    wp_count = total_row - first_wp_row
    assert wp_count == NUM_WPS


def test_burn_rate_wp_id_refs_plan(trends_ws):
    """Col A formulas reference Config WP rows."""
    header_row = _find_row_containing(trends_ws, 1, "WP ID")
    assert header_row is not None
    first_wp_row = header_row + 1
    for i in range(NUM_WPS):
        cell = trends_ws.cell(first_wp_row + i, 1)
        formula = str(cell.value)
        assert formula.startswith("=Config!A"), (
            f"WP row {i}: expected Config!A ref, got {formula}"
        )


def test_burn_rate_time_elapsed_formula(trends_ws):
    """Col C has formula referencing current_month for % Time Elapsed."""
    header_row = _find_row_containing(trends_ws, 1, "WP ID")
    assert header_row is not None
    first_wp_row = header_row + 1
    cell = trends_ws.cell(first_wp_row, 3)
    formula = str(cell.value)
    assert "current_month" in formula
    assert "Config!D" in formula


def test_burn_rate_budget_eur_formula(trends_ws):
    """Col D references Actuals cumulative columns for % Budget EUR."""
    header_row = _find_row_containing(trends_ws, 1, "WP ID")
    assert header_row is not None
    first_wp_row = header_row + 1
    cell = trends_ws.cell(first_wp_row, 4)
    formula = str(cell.value)
    assert "Actuals!" in formula
    # Should reference both the cumulative column and the ceiling column
    ceil_letter = col_letter(VAL_COL_BUDGET)
    assert f"Actuals!{ceil_letter}" in formula


def test_burn_rate_delta_formula(trends_ws):
    """Col E = D - C (Burn Delta)."""
    header_row = _find_row_containing(trends_ws, 1, "WP ID")
    assert header_row is not None
    first_wp_row = header_row + 1
    cell = trends_ws.cell(first_wp_row, 5)
    formula = str(cell.value)
    row = first_wp_row
    assert f"D{row}-C{row}" in formula or f"D{row} - C{row}" in formula


def test_burn_rate_status_formula(trends_ws):
    """Col F has IF/ABS formula for Status."""
    header_row = _find_row_containing(trends_ws, 1, "WP ID")
    assert header_row is not None
    first_wp_row = header_row + 1
    cell = trends_ws.cell(first_wp_row, 6)
    formula = str(cell.value)
    assert "IF" in formula
    assert "ABS" in formula


def test_burn_rate_exhaustion_formula(trends_ws):
    """Col G has COUNTIF + ROUND formula for Exhaustion Month."""
    header_row = _find_row_containing(trends_ws, 1, "WP ID")
    assert header_row is not None
    first_wp_row = header_row + 1
    cell = trends_ws.cell(first_wp_row, 7)
    formula = str(cell.value)
    assert "COUNTIF" in formula
    assert "ROUND" in formula


def test_burn_rate_total_row(trends_ws):
    """TOTAL row has 'TOTAL' in col B and weighted avg formula in col E."""
    total_row = _find_row(trends_ws, 2, "TOTAL")
    assert total_row is not None
    cell = trends_ws.cell(total_row, 5)
    formula = str(cell.value)
    assert "ISNUMBER" in formula
    assert "partner_ceiling" in formula


# ================================================================
# Person Load table
# ================================================================

def test_person_load_section_header(trends_ws):
    """Trends has a PERSON LOAD section header."""
    row = _find_row_containing(trends_ws, 1, "PERSON LOAD")
    assert row is not None, "PERSON LOAD section header not found in column A"
    assert "PERSON LOAD" in str(trends_ws.cell(row, 1).value)


def test_person_load_person_count(trends_ws):
    """4 person rows (from TEAM config)."""
    # Person header row is 33, data starts at 34
    header_row = _find_row_containing(trends_ws, 1, "Person")
    assert header_row is not None
    first_person_row = header_row + 1
    count = 0
    for i in range(20):  # generous upper bound
        cell = trends_ws.cell(first_person_row + i, 1)
        if cell.value is None or str(cell.value).strip() == "":
            break
        count += 1
    assert count == len(TEAM)


def test_person_load_avg_to_date_formula(trends_ws):
    """Col B has SUMPRODUCT formula for Avg FTE (to date)."""
    header_row = _find_row_containing(trends_ws, 1, "Person")
    assert header_row is not None
    first_person_row = header_row + 1
    cell = trends_ws.cell(first_person_row, 2)
    formula = str(cell.value)
    assert "SUMPRODUCT" in formula
    assert "current_month" in formula


def test_person_load_avg_remaining_formula(trends_ws):
    """Col C has SUMPRODUCT formula for Avg FTE (remaining)."""
    header_row = _find_row_containing(trends_ws, 1, "Person")
    assert header_row is not None
    first_person_row = header_row + 1
    cell = trends_ws.cell(first_person_row, 3)
    formula = str(cell.value)
    assert "SUMPRODUCT" in formula
    assert "current_month" in formula


def test_person_load_peak_formula(trends_ws):
    """Col D has MAX formula for Peak Month FTE."""
    header_row = _find_row_containing(trends_ws, 1, "Person")
    assert header_row is not None
    first_person_row = header_row + 1
    cell = trends_ws.cell(first_person_row, 4)
    formula = str(cell.value)
    assert "MAX" in formula
    assert "Capacity!" in formula


def test_person_load_status_formula(trends_ws):
    """Col E has IF formula with Config sheet reference."""
    header_row = _find_row_containing(trends_ws, 1, "Person")
    assert header_row is not None
    first_person_row = header_row + 1
    cell = trends_ws.cell(first_person_row, 5)
    formula = str(cell.value)
    assert "IF" in formula
    assert "Config!" in formula


# ================================================================
# Chart
# ================================================================

def test_cumulative_chart_exists(trends_ws):
    """Trends has at least one chart."""
    assert len(trends_ws._charts) >= 1


def test_chart_has_three_source_rows(trends_ws):
    """Chart data spans 3 helper rows (planned, actual, ceiling).

    Note: openpyxl expands multi-column References into individual series
    on save/load roundtrip, so we check chart title + that series count
    is a multiple of 3 (one per month column per source row).
    """
    chart = trends_ws._charts[0]
    # After roundtrip the title is a rich Title object; extract text runs
    title_text = "".join(
        run.t for p in chart.title.tx.rich.paragraphs for run in p.r
    )
    assert "Planned vs Actual" in title_text
    # 60 months * 3 rows = 180 series after roundtrip expansion
    assert len(chart.series) >= 3
    assert len(chart.series) % 3 == 0


def test_helper_rows_hidden(trends_ws):
    """Chart helper rows (between person load section and chart) are hidden."""
    # Discover helper rows dynamically: they are the hidden rows in the sheet
    hidden_rows = [
        r for r in range(1, trends_ws.max_row + 1)
        if trends_ws.row_dimensions[r].hidden
    ]
    assert len(hidden_rows) == 4, (
        f"Expected 4 hidden helper rows (planned, actual, ceiling, labels), "
        f"found {len(hidden_rows)}: {hidden_rows}"
    )
    # Verify they are contiguous
    for i in range(1, len(hidden_rows)):
        assert hidden_rows[i] == hidden_rows[i - 1] + 1, (
            f"Hidden helper rows should be contiguous, got {hidden_rows}"
        )


# ================================================================
# Conditional formatting
# ================================================================

def test_trends_has_conditional_formatting(trends_ws):
    """Trends sheet has conditional formatting rules."""
    assert len(trends_ws.conditional_formatting._cf_rules) > 0


# ================================================================
# Dynamic row layout — parametric tests for row collision safety
# ================================================================

def _make_wps(n):
    """Generate n synthetic work packages."""
    return [
        {"id": f"WP{i+1}", "name": f"Work Package {i+1}",
         "fte": 10.0, "start": 1, "end": 60}
        for i in range(n)
    ]


def _make_team(n, work_packages):
    """Generate n synthetic team members, each assigned to WP1."""
    wp_id = work_packages[0]["id"]
    return [
        {"name": f"Person{i+1}", "max_fte": 1.0, "wps": {wp_id: 1.0}}
        for i in range(n)
    ]


def _build_trends_only(num_wps, num_team):
    """Build a workbook with given WP/team counts and return the Trends ws + layout."""
    from datetime import datetime
    from openpyxl import Workbook as WB
    from budget_planner.workbook import _build_all_sheets
    from budget_planner.layout import SHEET_META, SHEET_TRENDS

    work_packages = _make_wps(num_wps)
    team = _make_team(num_team, work_packages)
    external_commitments: list[dict[str, object]] = []
    project_params = {
        "title": "Test", "start": datetime(2026, 1, 1),
        "months": 24, "rate": 7500,
    }

    wb = WB()
    meta_ws = wb.active
    assert meta_ws is not None
    meta_ws.title = SHEET_META

    _build_all_sheets(wb, project_params, work_packages, team,
                      external_commitments)

    trends_ws = wb[SHEET_TRENDS]
    return trends_ws


@pytest.mark.parametrize("num_wps,num_team", [
    (3, 2),    # small — fewer WPs than default
    (9, 4),    # current default config size
    (15, 8),   # large — more WPs and team members
])
def test_trends_no_row_collisions(num_wps, num_team):
    """Sections must not overlap regardless of WP / team count.

    Verifies that:
    - WP data rows don't overlap with chart anchor
    - Chart anchor + 15 (chart height) doesn't overlap person section
    - Person rows don't overlap with hidden helper rows
    - All hidden rows are contiguous and below person data
    """
    ws = _build_trends_only(num_wps, num_team)

    # Discover sections dynamically
    total_row = _find_row(ws, 2, "TOTAL")
    assert total_row is not None, "TOTAL row not found"

    wp_header = _find_row_containing(ws, 1, "WP ID")
    assert wp_header is not None
    first_wp = wp_header + 1
    last_wp = total_row - 1
    assert last_wp - first_wp + 1 == num_wps, (
        f"Expected {num_wps} WP rows, got {last_wp - first_wp + 1}"
    )

    person_section = _find_row_containing(ws, 1, "PERSON LOAD")
    assert person_section is not None, "PERSON LOAD section not found"

    person_header = _find_row_containing(ws, 1, "Person")
    assert person_header is not None
    first_person = person_header + 1

    # Count person data rows
    person_count = 0
    for i in range(50):
        cell = ws.cell(first_person + i, 1)
        if cell.value is None or str(cell.value).strip() == "":
            break
        person_count += 1
    assert person_count == num_team
    last_person = first_person + person_count - 1

    # Find hidden (helper) rows
    hidden_rows = sorted(
        r for r in range(1, ws.max_row + 1)
        if ws.row_dimensions[r].hidden
    )
    assert len(hidden_rows) == 4, (
        f"Expected 4 hidden helper rows, found {len(hidden_rows)}: {hidden_rows}"
    )

    # Verify no overlaps
    # Chart anchor should be after total row
    chart_anchor = total_row + 2  # matches the code
    assert chart_anchor > total_row, "Chart anchor must be after TOTAL row"

    # Person section must start after chart (chart height ~15 rows)
    assert person_section >= chart_anchor + 15, (
        f"Person section row {person_section} too close to chart at {chart_anchor}. "
        f"Chart needs ~15 rows of height."
    )

    # Helper rows must be after person data
    assert hidden_rows[0] > last_person, (
        f"First helper row {hidden_rows[0]} overlaps person data ending at {last_person}"
    )

    # Helper rows must be contiguous
    for i in range(1, len(hidden_rows)):
        assert hidden_rows[i] == hidden_rows[i - 1] + 1, (
            f"Hidden helper rows not contiguous: {hidden_rows}"
        )
