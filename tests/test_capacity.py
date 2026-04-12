"""Tests for the Capacity sheet builder."""

import pytest
import openpyxl

from budget_planner.workbook import init_workbook
from budget_planner.config import (
    TEAM, WORK_PACKAGES, EXTERNAL_COMMITMENTS, DEMO_CAPACITY_ASSIGNMENTS,
    MONTH_COL_START, MONTH_COL_END, PROJECT_MONTHS,
    VAL_COL_SUM, VAL_COL_BUDGET, VAL_COL_DIFF,
)
from budget_planner.styles import col_letter
from budget_planner.capacity import build_capacity_sheet, distribute


# ── Fixtures ──

@pytest.fixture(scope="module")
def full_wb(tmp_path_factory):
    """Generate a workbook with demo assignments and load it for inspection."""
    out = tmp_path_factory.mktemp("cap") / "test.xlsx"
    init_workbook(str(out), assignments=DEMO_CAPACITY_ASSIGNMENTS)
    wb = openpyxl.load_workbook(str(out), data_only=False)
    return wb


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
    meta_ws = wb.active
    assert meta_ws is not None
    meta_ws.title = "_META"
    config_ws = wb.create_sheet("Config")
    config_layout = build_config_sheet(config_ws, wb,
                                        {"title": "Test", "start": __import__("datetime").datetime(2026, 11, 1),
                                         "months": 14, "rate": 7500},
                                        WORK_PACKAGES,
                                        assignments=DEMO_CAPACITY_ASSIGNMENTS,
                                        external_commitments=EXTERNAL_COMMITMENTS)
    plan_ws = wb.create_sheet("Plan")
    plan_layout = build_plan_sheet(plan_ws, wb, config_layout=config_layout)
    cap_ws = wb.create_sheet("Capacity")
    layout = build_capacity_sheet(
        cap_ws, config_layout,
        assignments=DEMO_CAPACITY_ASSIGNMENTS,
        config_first_wp_row=plan_layout["config_first_wp_row"],
    )
    layout["config_layout"] = config_layout
    layout["plan_layout"] = plan_layout
    return layout


# ── Build person-assignment lookup for assertions ──

def _person_assigned_wps(person_name):
    """Return set of WP IDs assigned to a person in demo data."""
    return {a["wp"] for a in DEMO_CAPACITY_ASSIGNMENTS
            if a["person"] == person_name}


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
            f"{name}: expected WPs {expected_wps}, "
            f"got {set(group['wp_rows'].keys())}"
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


# ── Test: REMAINING row ──

def test_capacity_person_has_remaining_row(cap_ws, capacity_layout):
    """REMAINING row exists per person."""
    for member in TEAM:
        group = capacity_layout["person_groups"][member["name"]]
        remaining_row = group["remaining_row"]
        assert cap_ws.cell(row=remaining_row, column=2).value == "REMAINING"


# ── Test: WP rows are blue input ──

def test_capacity_wp_rows_are_blue_input(cap_ws, capacity_layout):
    """Monthly cells in programme WP rows have blue font."""
    for member in TEAM:
        name = member["name"]
        group = capacity_layout["person_groups"][name]
        for wp_id, row in group["wp_rows"].items():
            # Find the assignment to get the start month
            asgn = next(a for a in DEMO_CAPACITY_ASSIGNMENTS
                        if a["person"] == name and a["wp"] == wp_id)
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


# ── Test: REMAINING row formula ──

def test_capacity_remaining_row_formula(cap_ws, capacity_layout):
    """REMAINING references Config max_fte minus TOTAL."""
    for member in TEAM:
        group = capacity_layout["person_groups"][member["name"]]
        remaining_row = group["remaining_row"]
        val = cap_ws.cell(row=remaining_row, column=MONTH_COL_START).value
        assert isinstance(val, str) and "Config!" in val, (
            f"{member['name']} REMAINING M1: expected Config! ref, got {val}"
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
    asgn = next(a for a in DEMO_CAPACITY_ASSIGNMENTS
                if a["person"] == person_name and a["wp"] == wp_id)
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
            assert val == expected, (
                f"{member['name']}/{wp_id} Sum: expected {expected}, got {val}"
            )


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
            assert val == expected, (
                f"{name}/{wp_id} Budget: expected {expected}, got {val}"
            )


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
            assert val == expected, (
                f"{member['name']}/{wp_id} Variance: expected {expected}, got {val}"
            )


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
            assert sum_val == expected_sum, (
                f"{member['name']}/{project} Sum: expected {expected_sum}, got {sum_val}"
            )

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


# ── Test: EU TOTAL row ──

def test_capacity_person_has_eu_total_row(cap_ws, capacity_layout):
    """EU TOTAL row exists per person."""
    for member in TEAM:
        group = capacity_layout["person_groups"][member["name"]]
        eu_total_row = group.get("eu_total_row")
        assert eu_total_row is not None, f"{member['name']} missing eu_total_row"
        assert cap_ws.cell(row=eu_total_row, column=2).value == "EU TOTAL"


def test_capacity_eu_total_row_formula(cap_ws, capacity_layout):
    """EU TOTAL monthly cells have SUM formulas referencing WP rows."""
    for member in TEAM:
        name = member["name"]
        group = capacity_layout["person_groups"][name]
        eu_total_row = group["eu_total_row"]
        val = cap_ws.cell(row=eu_total_row, column=MONTH_COL_START).value
        # Should be a SUM formula (or =0 if no WP rows)
        assert isinstance(val, str) and val.startswith("="), (
            f"{name} EU TOTAL M1: expected formula, got {val}"
        )


def test_capacity_eu_total_includes_eu_external(cap_ws, capacity_layout):
    """EU TOTAL formula includes EU-flagged external rows, excludes non-EU."""
    from budget_planner.config import EXTERNAL_COMMITMENTS
    for member in TEAM:
        name = member["name"]
        group = capacity_layout["person_groups"][name]
        eu_total_row = group["eu_total_row"]
        formula = str(cap_ws.cell(row=eu_total_row, column=MONTH_COL_START).value)
        # EU-flagged externals should be in the formula
        for ext in EXTERNAL_COMMITMENTS:
            if ext["person"] == name and ext.get("is_eu"):
                ext_row = group["external_rows"][ext["project"]]
                assert str(ext_row) in formula, (
                    f"{name} EU TOTAL should reference EU external row {ext_row}"
                )
        # Non-EU externals should NOT be in the formula
        for ext in EXTERNAL_COMMITMENTS:
            if ext["person"] == name and not ext.get("is_eu"):
                ext_row = group["external_rows"][ext["project"]]
                assert str(ext_row) not in formula, (
                    f"{name} EU TOTAL should NOT reference non-EU external row {ext_row}"
                )
