"""Tests for the Capacity sheet builder."""

import pytest
import openpyxl

from budget_planner.config import (
    TEAM,
    WORK_PACKAGES,
    EXTERNAL_COMMITMENTS,
    DEMO_CAPACITY_ASSIGNMENTS,
    MONTH_COL_START,
    MONTH_COL_END,
    PROJECT_MONTHS,
    PROJECT_START,
    VAL_COL_SUM,
    VAL_COL_BUDGET,
    VAL_COL_DIFF,
)
from budget_planner.styles import col_letter
from budget_planner.capacity import build_capacity_sheet, distribute


# ── Fixtures ──


@pytest.fixture(scope="module")
def full_wb(assigned_wb_path):
    return openpyxl.load_workbook(assigned_wb_path, data_only=False)


@pytest.fixture(scope="module")
def cap_ws(full_wb):
    return full_wb["Capacity"]


@pytest.fixture(scope="module")
def capacity_layout(tmp_path_factory):
    """Build a full workbook with assignments and capture the capacity_layout."""
    from openpyxl import Workbook as WB
    from budget_planner.plan import build_plan_sheet
    from budget_planner.config_sheet import build_config_sheet

    wb = WB()
    config_ws = wb.create_sheet("Config")
    config_layout = build_config_sheet(
        config_ws,
        wb,
        {"title": "Test", "start": PROJECT_START, "months": 14, "rate": 7500},
        WORK_PACKAGES,
        assignments=DEMO_CAPACITY_ASSIGNMENTS,
        external_commitments=EXTERNAL_COMMITMENTS,
    )
    plan_ws = wb.create_sheet("Plan")
    plan_layout = build_plan_sheet(plan_ws, wb, config_layout=config_layout)
    cap_ws = wb.create_sheet("Capacity")
    layout = build_capacity_sheet(
        cap_ws,
        config_layout,
        assignments=DEMO_CAPACITY_ASSIGNMENTS,
        config_first_wp_row=plan_layout["config_first_wp_row"],
    )
    layout["config_layout"] = config_layout
    layout["plan_layout"] = plan_layout
    return layout


# ── Build person-assignment lookup for assertions ──


def _person_assigned_wps(person_name):
    """Return set of WP IDs assigned to a person in demo data."""
    return {a["wp"] for a in DEMO_CAPACITY_ASSIGNMENTS if a["person"] == person_name}


# ── Person groups ──


def test_capacity_has_person_group_per_team_member(capacity_layout):
    """One section per TEAM member."""
    expected_names = {m["name"] for m in TEAM}
    assert set(capacity_layout["person_groups"].keys()) == expected_names


# ── Test: programme WP rows per person (only assigned WPs) ──


def test_capacity_person_has_assigned_wp_rows(capacity_layout):
    """Each person has rows only for their assigned WPs."""
    for member in TEAM:
        name = member["name"]
        group = capacity_layout["person_groups"][name]
        expected_wps = _person_assigned_wps(name)
        assert set(group["wp_rows"].keys()) == expected_wps, (
            f"{name}: expected WPs {expected_wps}, got {set(group['wp_rows'].keys())}"
        )


# ── Test: External project rows ──


def test_capacity_person_has_external_project_rows(capacity_layout):
    """External project rows present for persons with commitments."""
    for entry in EXTERNAL_COMMITMENTS:
        group = capacity_layout["person_groups"][entry["person"]]
        assert entry["project"] in group["external_rows"], (
            f"{entry['person']} missing external row for {entry['project']}"
        )
    # Persons without external commitments should have empty external_rows
    persons_with_ext = {e["person"] for e in EXTERNAL_COMMITMENTS}
    for member in TEAM:
        if member["name"] not in persons_with_ext:
            group = capacity_layout["person_groups"][member["name"]]
            assert len(group["external_rows"]) == 0


# ─�� Test: TOTAL row ──


def test_capacity_person_has_total_row(cap_ws, capacity_layout):
    """TOTAL row exists per person."""
    for member in TEAM:
        group = capacity_layout["person_groups"][member["name"]]
        total_row = group["total_row"]
        assert cap_ws.cell(row=total_row, column=2).value == "TOTAL"


# ── Test: WP rows are blue input ──


def test_capacity_wp_rows_are_blue_input(cap_ws, capacity_layout):
    """Monthly cells in programme WP rows have blue font."""
    for member in TEAM:
        name = member["name"]
        group = capacity_layout["person_groups"][name]
        for wp_id, row in group["wp_rows"].items():
            # Find the assignment to get the start month
            asgn = next(a for a in DEMO_CAPACITY_ASSIGNMENTS if a["person"] == name and a["wp"] == wp_id)
            c = MONTH_COL_START + asgn["start"] - 1  # first active month
            font = cap_ws.cell(row=row, column=c).font
            assert font.color and font.color.rgb in ("0000FF", "000000FF"), (
                f"{name}/{wp_id} M{asgn['start']} font: {font.color}"
            )


# ── Test: TOTAL row SUM formula ──


def test_capacity_total_row_is_sum_formula(cap_ws, capacity_layout):
    """TOTAL monthly cells have SUM formulas."""
    for member in TEAM:
        group = capacity_layout["person_groups"][member["name"]]
        total_row = group["total_row"]
        # Check first month column
        val = cap_ws.cell(row=total_row, column=MONTH_COL_START).value
        assert isinstance(val, str) and val.startswith("=SUM("), (
            f"{member['name']} TOTAL M1: expected SUM formula, got {val}"
        )


# ── Test: Inactive months empty ──


def _inactive_params():
    """Generate (person_name, wp_id) for assigned WPs with inactive months."""
    params = []
    for asgn in DEMO_CAPACITY_ASSIGNMENTS:
        if asgn["start"] > 1 or asgn["end"] < PROJECT_MONTHS:
            params.append((asgn["person"], asgn["wp"]))
    return params[:3]


@pytest.mark.parametrize("person_name,wp_id", _inactive_params())
def test_capacity_inactive_months_empty(cap_ws, capacity_layout, person_name, wp_id):
    """Inactive months have empty value and grey fill."""
    asgn = next(a for a in DEMO_CAPACITY_ASSIGNMENTS if a["person"] == person_name and a["wp"] == wp_id)
    group = capacity_layout["person_groups"][person_name]
    row = group["wp_rows"][wp_id]

    for mi in range(PROJECT_MONTHS):
        month_num = mi + 1
        c = MONTH_COL_START + mi
        if month_num < asgn["start"] or month_num > asgn["end"]:
            val = cap_ws.cell(row=row, column=c).value
            assert val is None or val == "" or val == 0, (
                f"{person_name}/{wp_id} M{month_num}: '{val}' in inactive month"
            )
            fill = cap_ws.cell(row=row, column=c).fill
            assert fill.fgColor and fill.fgColor.rgb in ("D9D9D9", "00D9D9D9"), (
                f"{person_name}/{wp_id} M{month_num}: expected grey fill, got {fill.fgColor}"
            )


# ── Test: WP row Sum formula ──


def test_capacity_wp_row_sum_formula(cap_ws, capacity_layout):
    """programme WP rows have Sum = SUM(C{row}:..{row})."""
    start_cl = col_letter(MONTH_COL_START)
    end_cl = col_letter(MONTH_COL_END)
    for member in TEAM:
        group = capacity_layout["person_groups"][member["name"]]
        for wp_id, row in group["wp_rows"].items():
            val = cap_ws.cell(row=row, column=VAL_COL_SUM).value
            expected = f"=SUM({start_cl}{row}:{end_cl}{row})"
            assert val == expected, f"{member['name']}/{wp_id} Sum: expected {expected}, got {val}"


# ── Test: WP row Budget formula refs assignment config block ──


def test_capacity_wp_row_budget_formula(cap_ws, capacity_layout):
    """programme WP rows have Budget = Config!C{assign_config_row} (assignment Total FTE)."""
    assign_config_rows = capacity_layout["assign_config_rows"]

    for member in TEAM:
        name = member["name"]
        group = capacity_layout["person_groups"][name]
        for wp_id, row in group["wp_rows"].items():
            asgn_row = assign_config_rows.get((name, wp_id))
            assert asgn_row is not None, f"{name}/{wp_id} not in assign_config_rows"
            val = cap_ws.cell(row=row, column=VAL_COL_BUDGET).value
            expected = f"=Config!C{asgn_row}"
            assert val == expected, f"{name}/{wp_id} Budget: expected {expected}, got {val}"


# ── Test: WP row Variance formula = Budget − Sum ──


def test_capacity_wp_row_variance_formula(cap_ws, capacity_layout):
    """programme WP rows have Variance = {budget_col}{row} - {sum_col}{row}."""
    budget_cl = col_letter(VAL_COL_BUDGET)
    sum_cl = col_letter(VAL_COL_SUM)
    for member in TEAM:
        group = capacity_layout["person_groups"][member["name"]]
        for wp_id, row in group["wp_rows"].items():
            val = cap_ws.cell(row=row, column=VAL_COL_DIFF).value
            expected = f"={budget_cl}{row}-{sum_cl}{row}"
            assert val == expected, f"{member['name']}/{wp_id} Variance: expected {expected}, got {val}"


# ── Test: External row Sum/Budget/Variance formulas ──


def test_capacity_external_row_validation_formulas(cap_ws, capacity_layout):
    """External rows: Sum, Budget refs config Total FTE, Variance."""
    start_cl = col_letter(MONTH_COL_START)
    end_cl = col_letter(MONTH_COL_END)
    budget_cl = col_letter(VAL_COL_BUDGET)
    sum_cl = col_letter(VAL_COL_SUM)

    ext_config_rows = capacity_layout["external_config_rows"]

    for member in TEAM:
        group = capacity_layout["person_groups"][member["name"]]
        for project, row in group["external_rows"].items():
            # Sum
            sum_val = cap_ws.cell(row=row, column=VAL_COL_SUM).value
            expected_sum = f"=SUM({start_cl}{row}:{end_cl}{row})"
            assert sum_val == expected_sum, f"{member['name']}/{project} Sum: expected {expected_sum}, got {sum_val}"

            # Budget: =Config!C{ext_config_row}
            ext_config_row = ext_config_rows[(member["name"], project)]
            budget_val = cap_ws.cell(row=row, column=VAL_COL_BUDGET).value
            expected_budget = f"=Config!C{ext_config_row}"
            assert budget_val == expected_budget, (
                f"{member['name']}/{project} Budget: expected {expected_budget}, got {budget_val}"
            )

            # Variance: Budget - Sum
            var_val = cap_ws.cell(row=row, column=VAL_COL_DIFF).value
            expected_var = f"={budget_cl}{row}-{sum_cl}{row}"
            assert var_val == expected_var, (
                f"{member['name']}/{project} Variance: expected {expected_var}, got {var_val}"
            )


# ── Distribution curve tests ──


def test_distribute_flat():
    result = distribute(12.0, 6, "flat")
    assert len(result) == 6
    assert all(v == pytest.approx(2.0) for v in result)
    assert sum(result) == pytest.approx(12.0)


def test_distribute_front_loaded():
    result = distribute(10.0, 5, "front-loaded")
    assert len(result) == 5
    assert result[0] > result[-1]
    assert sum(result) == pytest.approx(10.0)


def test_distribute_back_loaded():
    result = distribute(10.0, 5, "back-loaded")
    assert len(result) == 5
    assert result[0] < result[-1]
    assert sum(result) == pytest.approx(10.0)


def test_distribute_bell():
    result = distribute(10.0, 7, "bell")
    assert len(result) == 7
    mid = len(result) // 2
    assert result[mid] > result[0]
    assert result[mid] > result[-1]
    assert sum(result) == pytest.approx(10.0)


def test_distribute_single_month():
    result = distribute(5.0, 1, "bell")
    assert result == [5.0]


def test_distribute_zero_months():
    result = distribute(5.0, 0, "flat")
    assert result == []


# ── Test: TOTAL-row over-allocation CF (folded REMAINING-row CF) ──


def test_capacity_total_row_overalloc_red_cf(cap_ws, capacity_layout):
    """TOTAL row has a red CF rule per person whose formula compares the
    monthly TOTAL cell against Config Max FTE (folds the dropped REMAINING
    < 0 CF onto TOTAL > Max FTE)."""
    from openpyxl.styles import PatternFill

    expected_total_rows = {capacity_layout["person_groups"][m["name"]]["total_row"] for m in TEAM}
    matched: set[int] = set()
    for cf_range in cap_ws.conditional_formatting:
        rules = cap_ws.conditional_formatting[cf_range]
        for rule in rules:
            if not (rule.dxf and rule.dxf.fill):
                continue
            fill = rule.dxf.fill
            if not isinstance(fill, PatternFill):
                continue
            bg = fill.bgColor and str(fill.bgColor.rgb)
            if not bg or not bg.endswith("FFC7CE"):
                continue
            formulas = list(rule.formula or [])
            for f in formulas:
                # Red TOTAL over-alloc CF: "<col><row>>Config!$B$<row>"
                # (no 0.8 amber multiplier, no SUM, no <0 — distinguishes
                # from amber TOTAL, EU CF, and variance reds.)
                if ">Config!$B$" in f and "0.8" not in f and "SUM(" not in f:
                    for total_row in expected_total_rows:
                        if f"{total_row}>Config!$B$" in f:
                            matched.add(total_row)
    assert matched == expected_total_rows, (
        f"Expected red over-alloc CF on every TOTAL row {sorted(expected_total_rows)}, matched {sorted(matched)}"
    )
