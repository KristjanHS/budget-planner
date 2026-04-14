"""Tests for T5 — Capacity to-date column block (Plan PM / Actual PM /
Variance PM / Variance %)."""

import pytest
import openpyxl

from budget_planner.workbook import init_workbook, seed_demo_data
from budget_planner.config import (
    DEMO_CAPACITY_ASSIGNMENTS, PROJECT_MONTHS,
)
from budget_planner.layout import SHEET_ACTUALS_HOURS, val_columns
from budget_planner.styles import col_letter


@pytest.fixture(scope="module")
def demo_wb(tmp_path_factory):
    out = tmp_path_factory.mktemp("td") / "test.xlsx"
    init_workbook(str(out), assignments=DEMO_CAPACITY_ASSIGNMENTS)
    seed_demo_data(str(out))
    return str(out)


@pytest.fixture(scope="module")
def cap_formulas(demo_wb):
    wb = openpyxl.load_workbook(demo_wb, data_only=False)
    return wb["Capacity"]


@pytest.fixture(scope="module")
def cap_values(demo_wb):
    # Recalc so data_only reads return numbers, not None.
    import subprocess
    subprocess.run(
        ["python", "scripts/recalc.py", demo_wb],
        check=True, capture_output=True,
    )
    wb = openpyxl.load_workbook(demo_wb, data_only=True)
    return wb["Capacity"]


@pytest.fixture(scope="module")
def todate_cols():
    vcols = val_columns(PROJECT_MONTHS)
    variance_end = vcols["diff"] + 2 + PROJECT_MONTHS - 1 + 1 + PROJECT_MONTHS
    # Match the computation in build_capacity_sheet:
    # actual_start = val_diff + 2
    # actual_end   = actual_start + months - 1
    # variance_start = actual_end + 2
    # variance_end = variance_start + months - 1
    # todate_start = variance_end + 2
    actual_start = vcols["diff"] + 2
    actual_end = actual_start + PROJECT_MONTHS - 1
    variance_start = actual_end + 2
    variance_end = variance_start + PROJECT_MONTHS - 1
    todate_start = variance_end + 2
    return {
        "plan": todate_start,
        "actual": todate_start + 1,
        "var": todate_start + 2,
        "pct": todate_start + 3,
    }


def _find_person_header_row(ws, name: str) -> int:
    """Locate a person's header row on Capacity.

    Capacity uses ``=Config!A<roster_row>`` formulas for person names. The
    roster section lives under the "TEAM ROSTER" header on Config, so we
    scan that specific section for ``name`` to avoid collisions with
    assignment/external-commitment rows that also hold person names.
    """
    # Literal-name scan for data_only workbooks.
    for r in range(1, ws.max_row + 1):
        if ws.cell(r, 1).value == name:
            return r
    # Formula mode: find the roster row specifically.
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


def test_todate_headers(cap_formulas, todate_cols):
    """Band label and per-column headers exist at expected columns."""
    assert cap_formulas.cell(1, todate_cols["plan"]).value == "TO-DATE"
    # month_header_row is 3 (band=1, year=2, month=3) given the standard layout
    for c, expected in [
        (todate_cols["plan"], "Plan PM"),
        (todate_cols["actual"], "Actual PM"),
        (todate_cols["var"], "Variance PM"),
        (todate_cols["pct"], "Variance %"),
    ]:
        # Scan the first few rows for the header — robust to layout shifts.
        found = False
        for r in range(1, 6):
            if cap_formulas.cell(r, c).value == expected:
                found = True
                break
        assert found, f"Header {expected!r} not found in col {c}"


def test_plan_pm_formula_shape(cap_formulas, todate_cols):
    """Plan PM cell sums first hours_actuals_through_month columns on same row."""
    row = _find_person_header_row(cap_formulas, "Alice") + 1  # first WP row
    formula = cap_formulas.cell(row, todate_cols["plan"]).value
    assert isinstance(formula, str)
    assert formula.startswith("=IF(hours_actuals_through_month<=0,0,")
    assert "INDEX" in formula
    assert f"${col_letter(3)}${row}" in formula  # anchor at col C, this row


def test_actual_pm_formula_references_actuals_hours(cap_formulas, todate_cols):
    """Actual PM cell references Actuals-Hours sheet and divides by
    hours_per_fte_month."""
    row = _find_person_header_row(cap_formulas, "Alice") + 1
    formula = cap_formulas.cell(row, todate_cols["actual"]).value
    assert isinstance(formula, str)
    assert SHEET_ACTUALS_HOURS in formula
    assert "hours_per_fte_month" in formula


def test_variance_and_pct_derive_from_plan_and_actual(cap_formulas, todate_cols):
    """Variance = Plan − Actual; Variance % guards div-by-zero."""
    row = _find_person_header_row(cap_formulas, "Alice") + 1
    var = cap_formulas.cell(row, todate_cols["var"]).value
    pct = cap_formulas.cell(row, todate_cols["pct"]).value
    assert var == f"={col_letter(todate_cols['plan'])}{row}-{col_letter(todate_cols['actual'])}{row}"
    assert pct.startswith(f"=IF(ABS({col_letter(todate_cols['plan'])}{row})<0.01,")


def test_external_rows_empty_on_todate_block(cap_formulas, todate_cols):
    """External project rows are blank/grey in the to-date block (no programme plan)."""
    # Find Alice's first external row (ProjectX — after her WP rows).
    alice_header = _find_person_header_row(cap_formulas, "Alice")
    # Walk rows after header until we hit an external row (col A literal string,
    # not a =Config!A... formula).
    ext_row = None
    for r in range(alice_header + 1, alice_header + 20):
        a = cap_formulas.cell(r, 1).value
        if isinstance(a, str) and a == "ProjectX":
            ext_row = r
            break
    assert ext_row is not None
    # All four to-date cells are empty string / unset value.
    for c in todate_cols.values():
        v = cap_formulas.cell(ext_row, c).value
        assert v in (None, ""), f"External row col {c} should be empty, got {v!r}"


def test_todate_values_match_seeded_demo(cap_values, todate_cols):
    """Round-trip: with demo seeded hours, Plan/Actual/Variance resolve to
    plausible positive numbers on Alice's WP1 row."""
    alice_header = _find_person_header_row(cap_values, "Alice")
    wp_row = alice_header + 1  # WP1
    plan = cap_values.cell(wp_row, todate_cols["plan"]).value
    actual = cap_values.cell(wp_row, todate_cols["actual"]).value
    var = cap_values.cell(wp_row, todate_cols["var"]).value
    assert plan is not None and plan > 0
    assert actual is not None and actual > 0
    # Expected: plan ≈ 4.0 FTE / 14 months × 12 months ≈ 3.43
    assert abs(plan - 3.43) < 0.01
    # Variance ≈ plan - actual (signed), demo multiplier WP1 = 1.0 so ~0.
    assert abs(var - (plan - actual)) < 1e-6


def test_todate_skip_and_cumulate(cap_values, todate_cols):
    """A WP with a mid-WP gap (WP3) still cumulates only its active months."""
    alice_header = _find_person_header_row(cap_values, "Alice")
    # Alice WP3 is the third WP row.
    wp3_row = alice_header + 3
    plan = cap_values.cell(wp3_row, todate_cols["plan"]).value
    actual = cap_values.cell(wp3_row, todate_cols["actual"]).value
    assert plan is not None and plan > 0
    # WP3 has gap in month 4; with DEMO_CURRENT_MONTH=12, only
    # months 1–3 and 5–7 are active-and-planned. Actual reflects the same.
    assert actual is not None and actual >= 0
    # Actual PM should be less than plan (demo multiplier for WP3 = 0.30).
    assert actual < plan


def test_total_row_sums_wp_rows(cap_values, todate_cols):
    """Person TOTAL row in the to-date block equals SUM of the person's WP rows."""
    alice_header = _find_person_header_row(cap_values, "Alice")
    wp_rows = [alice_header + i for i in range(1, 4)]  # Alice has 3 WPs
    # TOTAL row is after external rows; locate by col B = 'TOTAL'.
    total_row = None
    for r in range(alice_header + 1, alice_header + 20):
        if cap_values.cell(r, 2).value == "TOTAL":
            total_row = r
            break
    assert total_row is not None
    wp_plan_sum = sum(
        cap_values.cell(r, todate_cols["plan"]).value for r in wp_rows
    )
    total_plan = cap_values.cell(total_row, todate_cols["plan"]).value
    assert abs(total_plan - wp_plan_sum) < 1e-6
