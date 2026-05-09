"""Tests for the Actuals-Hours sheet builder (posted-cumulative hours)."""

import pytest
from openpyxl import load_workbook

from budget_planner.config import (
    DEMO_CAPACITY_ASSIGNMENTS, EXTERNAL_COMMITMENTS, TEAM,
    PROJECT_MONTHS, WORK_PACKAGES,
)
from budget_planner.layout import MONTH_COL_START, val_columns
from budget_planner.styles import GREY_FILL, col_letter


# Cum-hours helper col sits one gutter after the Sum validation col
# (forecast-EUR band was dropped 2026-05-09 with the posted-cumulative
# rewrite).
CUM_HOURS_COL = val_columns(PROJECT_MONTHS)["sum"] + 2


@pytest.fixture(scope="module")
def wb(assigned_wb_path):
    return load_workbook(assigned_wb_path, data_only=False)


@pytest.fixture(scope="module")
def hrs(wb):
    return wb["Actuals-Hours"]


def _person_assigned_wps(person_name):
    return {a["wp"] for a in DEMO_CAPACITY_ASSIGNMENTS
            if a["person"] == person_name}


def _find_person_header_row(ws, person_name):
    """Find row of the person header (merged A:B) for ``person_name``.

    Person headers carry the literal person name in col A (display under
    data_only reads must be stable). We locate the row by matching col-A
    value == person_name AND verifying the A:B merge.
    """
    merged_ranges = [(mr.min_row, mr.min_col, mr.max_col)
                     for mr in ws.merged_cells.ranges]
    for r in range(1, ws.max_row + 1):
        if ws.cell(r, 1).value != person_name:
            continue
        is_person = any(
            mr_min_r == r and mr_min_c == 1 and mr_max_c >= 2
            for mr_min_r, mr_min_c, mr_max_c in merged_ranges
        )
        if is_person:
            return r
    raise AssertionError(
        f"Person header for {person_name!r} not found on {ws.title}"
    )


# ── Existence ──
# Tab position covered by tests/test_tab_order.py (Actuals-Hours = position 3).


def test_actuals_hours_sheet_exists(wb):
    assert "Actuals-Hours" in wb.sheetnames


# ── Structure: person groups match team size ──

def test_actuals_hours_has_person_group_per_team_member(hrs):
    """One header row per team member."""
    team_names = [m["name"] for m in TEAM]
    rows = [_find_person_header_row(hrs, n) for n in team_names]
    assert len(set(rows)) == len(team_names)
    # Rows monotonically increase (sections stack top-down)
    assert rows == sorted(rows)


def test_actuals_hours_person_has_assigned_wp_rows(hrs):
    """Each person has a WP-ID row for each assigned WP, then externals,
    then TOTAL."""
    for member in TEAM:
        name = member["name"]
        header_row = _find_person_header_row(hrs, name)
        expected_wps = _person_assigned_wps(name)
        expected_externals = [e for e in EXTERNAL_COMMITMENTS
                              if e["person"] == name]

        # WP rows immediately follow header
        wp_row_count = 0
        r = header_row + 1
        while True:
            v = hrs.cell(r, 1).value
            if isinstance(v, str) and v.startswith("=Config!A"):
                wp_row_count += 1
                r += 1
                continue
            # External project row: col A = literal project name (str)
            if isinstance(v, str) and v in {e["project"]
                                             for e in expected_externals}:
                r += 1
                continue
            break
        assert wp_row_count == len(expected_wps), (
            f"{name}: expected {len(expected_wps)} WP rows, got {wp_row_count}"
        )

        # Next row should be TOTAL label
        total_label = hrs.cell(r, 2).value
        assert total_label == "TOTAL", (
            f"{name}: expected TOTAL at row {r}, got {total_label!r}"
        )


# ── Inactive months: grey + empty ──

def test_actuals_hours_inactive_month_is_empty_and_grey(hrs):
    """The first WP row for Bob, at any month before that WP's active start,
    must be empty + grey-filled. Per project testing rule, derive Bob's
    first WP and its active_start from DEMO_CAPACITY_ASSIGNMENTS +
    WORK_PACKAGES — don't hardcode WP IDs."""
    bob_wps = [a for a in DEMO_CAPACITY_ASSIGNMENTS if a["person"] == "Bob"]
    if not bob_wps:
        pytest.skip("Bob has no WP assignments in current demo")
    first_wp_code = bob_wps[0]["wp"]  # row order matches DEMO list order
    wp_def = next(w for w in WORK_PACKAGES if w["id"] == first_wp_code)
    if wp_def["start"] <= 1:
        pytest.skip(
            f"Bob's first WP {first_wp_code} is active at M1 — "
            f"no inactive cell available for this assertion"
        )
    bob_header = _find_person_header_row(hrs, "Bob")
    first_wp_row = bob_header + 1
    m1_col = MONTH_COL_START  # M1
    cell = hrs.cell(first_wp_row, m1_col)
    assert cell.value in (None, ""), (
        f"Bob {first_wp_code} @ M1 should be empty (inactive); "
        f"got {cell.value!r}"
    )
    fill = cell.fill
    assert fill is not None
    # PatternFill equality is stricter — compare by fgColor rgb
    assert str(getattr(fill.fgColor, "rgb", "")).upper().endswith(
        str(GREY_FILL.fgColor.rgb).upper()
    ), (
        f"Bob {first_wp_code} @ M1 fill should be GREY_FILL, "
        f"got {fill.fgColor.rgb!r}"
    )


# ── Active months: blue input + empty (no demo hours pre-filled) ──

def test_actuals_hours_active_month_is_blue_and_empty(hrs):
    """Alice WP1 is active M1-M14. Her WP1 row at M1 must be blue + empty."""
    alice_header = _find_person_header_row(hrs, "Alice")
    wp1_row = alice_header + 1
    m1_col = MONTH_COL_START
    cell = hrs.cell(wp1_row, m1_col)
    # Empty (no demo pre-fill on hours)
    assert cell.value in (None, ""), (
        f"Alice WP1 @ M1 should be empty; got {cell.value!r}"
    )
    # Blue font (0000FF)
    color = getattr(cell.font.color, "rgb", None)
    assert color is not None
    assert str(color).upper().endswith("0000FF"), (
        f"Alice WP1 @ M1 font should be blue, got {color!r}"
    )
    # Integer number format — no decimals
    assert "0" in cell.number_format
    assert "." not in cell.number_format


# ── TOTAL row formula ──

def test_actuals_hours_total_row_sums_data_rows(hrs):
    """TOTAL row M1 formula = SUM of the data rows above for that person."""
    alice_header = _find_person_header_row(hrs, "Alice")
    # Derive Alice's expected data-row count from DEMO_CAPACITY_ASSIGNMENTS +
    # EXTERNAL_COMMITMENTS so this test tracks demo edits per the project
    # testing rule (no hardcoded seed values).
    alice_wps = sum(1 for a in DEMO_CAPACITY_ASSIGNMENTS if a["person"] == "Alice")
    alice_externals = sum(1 for e in EXTERNAL_COMMITMENTS if e["person"] == "Alice")
    expected_data_rows = alice_wps + alice_externals
    data_row_start = alice_header + 1
    data_row_end = alice_header + expected_data_rows
    total_row = data_row_end + 1
    assert hrs.cell(total_row, 2).value == "TOTAL"

    m1_col = MONTH_COL_START
    cl = col_letter(m1_col)
    formula = hrs.cell(total_row, m1_col).value
    expected = f"=SUM({cl}{data_row_start}:{cl}{data_row_end})"
    assert formula == expected, (
        f"TOTAL @ M1 expected {expected!r}, got {formula!r}"
    )


def test_person_header_is_literal_name(hrs):
    """Person header cell (col A of the merged A:B banner) must carry the
    literal person name, not a formula."""
    for member in TEAM:
        name = member["name"]
        header_row = _find_person_header_row(hrs, name)
        value = hrs.cell(header_row, 1).value
        assert isinstance(value, str), (
            f"Person header for {name} must be a string, got {value!r}"
        )
        assert not value.startswith("="), (
            f"Person header for {name} must be literal, got formula {value!r}"
        )
        assert value == name, (
            f"Person header at row {header_row} expected {name!r}, "
            f"got {value!r}"
        )


def test_actuals_hours_has_only_sum_validation_column(hrs):
    """Year header row has 'Sum' at val_col_sum and no Budget/Variance
    label at val_col_budget — the hours log has no budget-vs-actual view.
    """
    vcols = val_columns(PROJECT_MONTHS)
    # Year header row = 2 (row 1 = through-month strip, row 2 = year,
    # row 3 = month) under posted-cumulative layout.
    year_row = 2
    sum_label = hrs.cell(year_row, vcols["sum"]).value
    budget_label = hrs.cell(year_row, vcols["budget"]).value
    assert sum_label == "Sum"
    assert budget_label in (None, "")


# ── Cum-hours helper column (feeds Capacity to-date Actual PM) ──

def test_cum_hours_header_present(hrs):
    """Header row carries the 'Cum hours (to date)' label."""
    # Month header row = 3 (strip=1, year=2, month=3).
    assert hrs.cell(3, CUM_HOURS_COL).value == "Cum hours (to date)"


def test_cum_hours_wp_row_formula_shape(hrs):
    """WP rows carry a SUMPRODUCT-MAX-INDEX latest-filled scan keyed on
    ``hours_actuals_through_month`` (posted-cumulative semantics — the
    helper returns the latest filled cell, not a SUM that would
    over-count cumulative snapshots)."""
    alice_row = _find_person_header_row(hrs, "Alice") + 1  # first WP row
    formula = hrs.cell(alice_row, CUM_HOURS_COL).value
    assert isinstance(formula, str)
    assert formula.startswith("=IFERROR(INDEX(")
    assert "SUMPRODUCT(MAX(" in formula
    assert "ISNUMBER(" in formula
    assert "hours_actuals_through_month" in formula
    assert "SUM(" not in formula  # the broken-shape regression guard
    start_cl = col_letter(MONTH_COL_START)
    end_cl = col_letter(MONTH_COL_START + PROJECT_MONTHS - 1)
    assert f"${start_cl}${alice_row}:${end_cl}${alice_row}" in formula


def test_cum_hours_total_formula(hrs):
    """TOTAL row also carries the latest-filled scan — Capacity's
    aggregate path doesn't need it today, but the helper col is uniform
    per row so cross-section reads remain consistent."""
    alice_header = _find_person_header_row(hrs, "Alice")
    total_row = None
    for r in range(alice_header + 1, alice_header + 30):
        if hrs.cell(r, 2).value == "TOTAL":
            total_row = r
            break
    assert total_row is not None
    formula = hrs.cell(total_row, CUM_HOURS_COL).value
    assert isinstance(formula, str)
    assert "hours_actuals_through_month" in formula
    assert "INDEX" in formula
    assert "SUMPRODUCT(MAX(" in formula
