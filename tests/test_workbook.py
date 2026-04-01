"""Structural checks for the generated workbook — maps from verify.py."""
import pytest
import openpyxl

from budget_planner.workbook import generate_workbook
from budget_planner.reader import discover_layout
from budget_planner.config import WORK_PACKAGES, NUM_WPS, MONTH_COL_START, PROJECT_MONTHS


# ── Fixtures that don't need LibreOffice ──

@pytest.fixture(scope="module")
def fresh_path(tmp_path_factory):
    out = tmp_path_factory.mktemp("struct") / "test.xlsx"
    generate_workbook(str(out))
    return str(out)


@pytest.fixture(scope="module")
def workbook(fresh_path):
    return openpyxl.load_workbook(fresh_path, data_only=False)


@pytest.fixture(scope="module")
def layout(workbook):
    return discover_layout(workbook["Plan"])


# ── Check 2-3: Sheet existence + global params ──

def test_sheets_exist(workbook):
    for name in ["_META", "Plan", "Team", "Capacity", "Summary"]:
        assert name in workbook.sheetnames, f"Missing sheet: {name}"


def test_global_params(workbook):
    plan = workbook["Plan"]
    assert plan["B2"].value is not None, "Plan!B2 (project_start) is empty"
    val = plan["E2"].value
    assert val is not None and val > 0, f"Plan!E2 (fte_rate) is invalid: {val}"


# ── Check 4: Named ranges ──

def test_named_ranges(workbook):
    defined = {}
    for name in workbook.defined_names:
        dn = workbook.defined_names[name]
        defined[dn.name] = dn.attr_text
    assert defined.get("project_start") == "Plan!$B$2"
    assert defined.get("fte_rate") == "Plan!$E$2"


# ── Check 5: Config table ──

@pytest.mark.parametrize("wp_idx", [0, 2])
def test_wp_config_valid(workbook, layout, wp_idx):
    plan = workbook["Plan"]
    row = layout["config_first_wp_row"] + wp_idx
    wp_id = plan.cell(row, 1).value
    fte = plan.cell(row, 3).value
    start = plan.cell(row, 4).value
    end = plan.cell(row, 5).value
    assert wp_id is not None, f"WP ID is empty at row {row}"
    assert fte is not None and fte > 0, f"{wp_id}: FTE invalid ({fte})"
    assert start is not None and end is not None, f"{wp_id}: Start/End empty"
    assert start <= end, f"{wp_id}: Start ({start}) > End ({end})"


# ── Actual FTE — should be zeros in fresh workbook ──

@pytest.mark.parametrize("wp_idx", [0, 2])
def test_actual_fte_zeros(workbook, layout, wp_idx):
    """Fresh workbook actual FTE cells should be 0 in active months."""
    plan = workbook["Plan"]
    row = layout["actual_first_wp_row"] + wp_idx
    wp = WORK_PACKAGES[wp_idx]
    for m in range(PROJECT_MONTHS):
        month_num = m + 1
        col = MONTH_COL_START + m
        val = plan.cell(row, col).value
        if wp["start"] <= month_num <= wp["end"]:
            assert val == 0, f"WP{wp_idx+1} M{month_num}: expected 0, got {val}"


# ── Check 11: Inactive months empty ──

@pytest.mark.parametrize("wp_idx", [1, 2])
def test_inactive_months_empty(workbook, layout, wp_idx):
    plan = workbook["Plan"]
    wp = WORK_PACKAGES[wp_idx]
    for block_key in ["planned_first_wp_row", "actual_first_wp_row"]:
        row = layout[block_key] + wp_idx
        for m in range(PROJECT_MONTHS):
            month_num = m + 1
            col = MONTH_COL_START + m
            if month_num < wp["start"] or month_num > wp["end"]:
                val = plan.cell(row, col).value
                assert val is None or val == "" or val == 0, \
                    f"WP{wp_idx+1} M{month_num} row {row}: '{val}' in inactive month"


# ── Check 12: Summary ──

def test_summary_wp_count(workbook):
    summary = workbook["Summary"]
    count = 0
    for r in range(1, summary.max_row + 1):
        val = summary.cell(r, 1).value
        if val and isinstance(val, str) and val.startswith("="):
            count += 1
    assert count == NUM_WPS, f"Summary has {count} WP rows, expected {NUM_WPS}"


# ── _META ──

def test_meta_has_modification_guide(workbook):
    meta = workbook["_META"]
    content = str(meta["A1"].value or "")
    assert "MODIFICATION GUIDE" in content
