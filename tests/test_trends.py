"""Tests for the Trends sheet builder."""

import pytest
import openpyxl

from budget_planner.workbook import generate_workbook
from budget_planner.config import (
    NUM_WPS, TEAM,
    VAL_COL_BUDGET,
)
from budget_planner.styles import col_letter


# ── Fixtures ──

@pytest.fixture(scope="module")
def trends_workbook(tmp_path_factory):
    out = tmp_path_factory.mktemp("trends") / "test.xlsx"
    generate_workbook(str(out))
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
    assert trends_ws["B1"].value == 1


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
    """Col A formulas reference Plan config rows."""
    header_row = _find_row_containing(trends_ws, 1, "WP ID")
    assert header_row is not None
    first_wp_row = header_row + 1
    for i in range(NUM_WPS):
        cell = trends_ws.cell(first_wp_row + i, 1)
        formula = str(cell.value)
        assert formula.startswith("=Plan!A"), (
            f"WP row {i}: expected Plan!A ref, got {formula}"
        )


def test_burn_rate_time_elapsed_formula(trends_ws):
    """Col C has formula referencing current_month for % Time Elapsed."""
    header_row = _find_row_containing(trends_ws, 1, "WP ID")
    assert header_row is not None
    first_wp_row = header_row + 1
    cell = trends_ws.cell(first_wp_row, 3)
    formula = str(cell.value)
    assert "current_month" in formula
    assert "Plan!D" in formula


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
    assert "SUMPRODUCT" in formula
    assert "partner_ceiling" in formula


# ================================================================
# Person Load table
# ================================================================

def test_person_load_section_header(trends_ws):
    """Row 32 has PERSON LOAD section header."""
    val = trends_ws.cell(32, 1).value
    assert val is not None
    assert "PERSON LOAD" in str(val)


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
    """Col E has IF formula with Team sheet reference."""
    header_row = _find_row_containing(trends_ws, 1, "Person")
    assert header_row is not None
    first_person_row = header_row + 1
    cell = trends_ws.cell(first_person_row, 5)
    formula = str(cell.value)
    assert "IF" in formula
    assert "Team!" in formula


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
    """Helper rows 40, 41, 42, 43 are hidden."""
    for r in [40, 41, 42, 43]:
        assert trends_ws.row_dimensions[r].hidden, f"Row {r} should be hidden"


# ================================================================
# Conditional formatting
# ================================================================

def test_trends_has_conditional_formatting(trends_ws):
    """Trends sheet has conditional formatting rules."""
    assert len(trends_ws.conditional_formatting._cf_rules) > 0
