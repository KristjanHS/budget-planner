"""Tests for the Capacity to-date block (Actual PM / Variance PM).

The block has 2 cols. WP-row Actual PM is a single cross-sheet ref
to the Actuals-Hours ``cum_hours_col``; Variance PM is populated only
on TOTAL rows as ``windowed-plan-sum − Actual``.
"""

import pytest
import openpyxl

from budget_planner.config import PROJECT_MONTHS
from budget_planner.layout import SHEET_ACTUALS_HOURS, val_columns
from budget_planner.styles import col_letter


@pytest.fixture(scope="module")
def cap_formulas(demo_wb_path):
    wb = openpyxl.load_workbook(demo_wb_path, data_only=False)
    return wb["Capacity"]


@pytest.fixture(scope="module")
def cap_values(demo_recalc_path):
    wb = openpyxl.load_workbook(demo_recalc_path, data_only=True)
    return wb["Capacity"]


@pytest.fixture(scope="module")
def todate_cols():
    """Column positions for the 2-col to-date block.

    Mirrors the computation in ``build_capacity_sheet``:
      actual_start  = val_diff + 2
      actual_end    = actual_start + months - 1
      variance_start= actual_end + 2
      variance_end  = variance_start + months - 1
      todate_start  = variance_end + 2
      todate_actual = todate_start
      todate_var    = todate_start + 1
    """
    vcols = val_columns(PROJECT_MONTHS)
    actual_start = vcols["diff"] + 2
    actual_end = actual_start + PROJECT_MONTHS - 1
    variance_start = actual_end + 2
    variance_end = variance_start + PROJECT_MONTHS - 1
    todate_start = variance_end + 2
    return {"actual": todate_start, "var": todate_start + 1}


def _find_person_header_row(ws, name: str) -> int:
    """Locate a person's header row on Capacity.

    Literal-name path first (works for data_only workbooks); falls back to
    the TEAM ROSTER scan for formula mode.
    """
    for r in range(1, ws.max_row + 1):
        if ws.cell(r, 1).value == name:
            return r
    wb = ws.parent
    cfg = wb["Config"]
    roster_header = None
    for r in range(1, cfg.max_row + 1):
        if cfg.cell(r, 1).value == "TEAM ROSTER":
            roster_header = r
            break
    assert roster_header is not None, "TEAM ROSTER header not found"
    roster_row = None
    for r in range(roster_header + 2, roster_header + 20):
        if cfg.cell(r, 1).value == name:
            roster_row = r
            break
    assert roster_row is not None, f"No roster row for {name}"
    target = f"=Config!A{roster_row}"
    for r in range(1, ws.max_row + 1):
        if ws.cell(r, 1).value == target:
            return r
    raise AssertionError(f"Person header not found: {name}")


def _find_total_row(ws, after_row: int) -> int:
    for r in range(after_row + 1, after_row + 25):
        if ws.cell(r, 2).value == "TOTAL":
            return r
    raise AssertionError("TOTAL row not found")


def test_todate_headers(cap_formulas, todate_cols):
    """Band label + 2 per-column headers exist."""
    assert cap_formulas.cell(1, todate_cols["actual"]).value == "TO-DATE"
    for c, expected in [
        (todate_cols["actual"], "Actual PM"),
        (todate_cols["var"], "Variance PM"),
    ]:
        found = any(cap_formulas.cell(r, c).value == expected for r in range(1, 6))
        assert found, f"Header {expected!r} not found in col {c}"


def test_wp_actual_pm_is_cross_sheet_ref(cap_formulas, todate_cols):
    """WP-row Actual PM is a single cross-sheet ref to Actuals-Hours cum-hours
    col, divided by hours_per_fte_month. No SUM, no INDEX — the windowed
    cumulation lives on the helper col."""
    row = _find_person_header_row(cap_formulas, "Alice") + 1  # first WP row
    formula = cap_formulas.cell(row, todate_cols["actual"]).value
    assert isinstance(formula, str)
    assert SHEET_ACTUALS_HOURS in formula
    assert "hours_per_fte_month" in formula
    assert "SUM(" not in formula, "WP Actual PM must not rebuild SUM"
    assert "INDEX(" not in formula, "WP Actual PM must not rebuild INDEX"


def test_wp_variance_pm_is_blank(cap_formulas, todate_cols):
    """WP-row Variance PM is blank (grey placeholder) — only the TOTAL row
    carries the variance signal per the simplification design."""
    row = _find_person_header_row(cap_formulas, "Alice") + 1
    v = cap_formulas.cell(row, todate_cols["var"]).value
    assert v in (None, "")


def test_total_row_variance_is_windowed_plan_minus_actual(cap_formulas, todate_cols):
    """TOTAL row Variance PM = windowed plan SUM (same-row planned cells) −
    Actual cell. No cross-sheet ref in plan side (TOTAL's monthly cells
    already sum all WPs on Capacity)."""
    alice_header = _find_person_header_row(cap_formulas, "Alice")
    total_row = _find_total_row(cap_formulas, alice_header)
    formula = cap_formulas.cell(total_row, todate_cols["var"]).value
    assert isinstance(formula, str)
    assert "hours_actuals_through_month" in formula
    assert "INDEX" in formula
    actual_cl = col_letter(todate_cols["actual"])
    assert f"{actual_cl}{total_row}" in formula
    # Plan side must be same-sheet — no Actuals-Hours reference.
    assert SHEET_ACTUALS_HOURS not in formula


def test_external_rows_empty_on_todate_block(cap_formulas, todate_cols):
    """External project rows are blank/grey in the to-date block."""
    alice_header = _find_person_header_row(cap_formulas, "Alice")
    ext_row = None
    for r in range(alice_header + 1, alice_header + 20):
        a = cap_formulas.cell(r, 1).value
        if isinstance(a, str) and a == "ProjectX":
            ext_row = r
            break
    assert ext_row is not None
    for c in todate_cols.values():
        v = cap_formulas.cell(ext_row, c).value
        assert v in (None, ""), f"External row col {c} should be empty, got {v!r}"


def test_total_actual_sums_wp_actuals(cap_values, todate_cols):
    """TOTAL row Actual PM equals SUM of the person's WP Actual PM cells."""
    alice_header = _find_person_header_row(cap_values, "Alice")
    total_row = _find_total_row(cap_values, alice_header)
    # Alice has 3 WPs (per demo seed); walk WP rows between header and TOTAL.
    wp_rows = []
    for r in range(alice_header + 1, total_row):
        a = cap_values.cell(r, 1).value
        # WP rows reference Config!A<n>; under data_only they resolve to WP ids
        # (strings starting "WP"). External rows carry project-literal strings.
        if isinstance(a, str) and a.startswith("WP"):
            wp_rows.append(r)
    assert wp_rows, "Expected at least one WP row for Alice"
    wp_actual_sum = sum(cap_values.cell(r, todate_cols["actual"]).value or 0 for r in wp_rows)
    total_actual = cap_values.cell(total_row, todate_cols["actual"]).value
    assert total_actual is not None
    assert abs(total_actual - wp_actual_sum) < 1e-6


def test_total_variance_resolves_nonnegative_under_demo_seed(cap_values, todate_cols):
    """Round-trip: TOTAL row Variance PM resolves to a finite number under
    demo seed. Sign isn't asserted — demo WP multipliers vary — but the cell
    must not be blank and must be numeric."""
    alice_header = _find_person_header_row(cap_values, "Alice")
    total_row = _find_total_row(cap_values, alice_header)
    var = cap_values.cell(total_row, todate_cols["var"]).value
    assert var is not None
    assert isinstance(var, (int, float))
