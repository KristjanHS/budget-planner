"""Structural checks for the generated workbook — maps from verify.py."""
import re

import pytest
import openpyxl

from budget_planner.reader import discover_layout
from budget_planner.config import WORK_PACKAGES, NUM_WPS, MONTH_COL_START, PROJECT_MONTHS


# ── Fixtures that don't need LibreOffice ──

@pytest.fixture(scope="module")
def workbook(fresh_wb_path):
    return openpyxl.load_workbook(fresh_wb_path, data_only=False)


@pytest.fixture(scope="module")
def layout(workbook):
    return discover_layout(workbook["Plan"])


# ── Check 2-3: Sheet existence + global params ──

def test_sheets_exist(workbook):
    for name in ["_META", "Config", "Plan", "Capacity", "Actuals", "Summary"]:
        assert name in workbook.sheetnames, f"Missing sheet: {name}"


def test_global_params(workbook):
    # Plan compact header references Config sheet via formulas
    plan = workbook["Plan"]
    # B1 = title ref, E1 = start ref, G1 = rate ref, I1 = ceiling ref
    assert plan["B1"].value is not None, "Plan!B1 (title ref) is empty"
    assert plan["E1"].value is not None, "Plan!E1 (start ref) is empty"
    assert plan["G1"].value is not None, "Plan!G1 (rate ref) is empty"
    assert plan["I1"].value is not None, "Plan!I1 (ceiling ref) is empty"
    # Check Config directly for valid rate (find by label, not hardcoded row)
    # FTE Rate is now a formula (weighted average of grade rates), not a static value
    config = workbook["Config"]
    rate_val = None
    for r in range(1, config.max_row + 1):
        if "fte rate" in str(config.cell(r, 1).value or "").lower():
            rate_val = config.cell(r, 2).value
            break
    assert rate_val is not None, "Config FTE Rate cell is empty"
    assert isinstance(rate_val, str) and rate_val.startswith("="), (
        f"Config FTE Rate should be a formula, got: {rate_val}"
    )


# ── Check 4: Named ranges ──

def test_named_ranges(workbook):
    defined = {}
    for name in workbook.defined_names:
        dn = workbook.defined_names[name]
        defined[dn.name] = dn.attr_text
    assert defined.get("project_start") == "Config!$B$2"
    assert defined.get("fte_rate") == "Config!$B$4"
    assert defined.get("partner_ceiling") == "Plan!$I$1"


# ── Check 5: Config table ──

@pytest.mark.parametrize("wp_idx", [0, 2])
def test_wp_config_valid(workbook, layout, wp_idx):
    config = workbook["Config"]
    row = layout["config_first_wp_row"] + wp_idx
    wp_id = config.cell(row, 1).value
    fte = config.cell(row, 3).value
    start = config.cell(row, 4).value
    end = config.cell(row, 5).value
    assert wp_id is not None, f"WP ID is empty at row {row}"
    assert fte is not None and fte > 0, f"{wp_id}: FTE invalid ({fte})"
    assert start is not None and end is not None, f"{wp_id}: Start/End empty"
    assert start <= end, f"{wp_id}: Start ({start}) > End ({end})"


# ── Check 11: Inactive months empty ──

def _wp_indices_with_inactive_months():
    """Find WP indices that have inactive months (start > 1 or end < PROJECT_MONTHS)."""
    return [
        i for i, wp in enumerate(WORK_PACKAGES)
        if wp["start"] > 1 or wp["end"] < PROJECT_MONTHS
    ][:3]


@pytest.mark.parametrize("wp_idx", _wp_indices_with_inactive_months())
def test_inactive_months_empty(workbook, layout, wp_idx):
    plan = workbook["Plan"]
    wp = WORK_PACKAGES[wp_idx]
    row = layout["planned_first_wp_row"] + wp_idx
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
    # Budget Overview + RP Summary each have one formula row per WP
    assert count == NUM_WPS * 2, f"Summary has {count} WP rows, expected {NUM_WPS * 2}"


# ── _META ──

def test_meta_has_modification_guide(workbook):
    meta = workbook["_META"]
    content = str(meta["A1"].value or "")
    assert "MODIFICATION GUIDE" in content


# ── Freeze panes ──


def _parse_freeze_panes(fp: str) -> tuple[str, int]:
    """Parse a freeze_panes string like 'C12' into (column, row)."""
    m = re.match(r"^([A-Z]+)(\d+)$", fp)
    assert m is not None, f"freeze_panes '{fp}' doesn't match expected pattern"
    return m.group(1), int(m.group(2))


@pytest.mark.parametrize("sheet_name", ["Plan", "Capacity", "Actuals", "Summary"])
def test_freeze_panes_set_at_column_c(workbook, sheet_name):
    """All data sheets freeze columns A-B (WP ID + Name) at column C."""
    ws = workbook[sheet_name]
    fp = ws.freeze_panes
    assert fp is not None, f"{sheet_name} has no freeze_panes set"
    col, row = _parse_freeze_panes(fp)
    assert col == "C", f"{sheet_name} freeze_panes column is '{col}', expected 'C'"
    assert row > 1, f"{sheet_name} freeze_panes row {row} should be > 1"
    assert row <= ws.max_row, (
        f"{sheet_name} freeze_panes row {row} exceeds max_row {ws.max_row}"
    )


def test_freeze_panes_plan_exact(workbook, layout):
    """Plan freeze_panes is exactly C{planned_first_wp_row}."""
    plan = workbook["Plan"]
    expected = f"C{layout['planned_first_wp_row']}"
    assert plan.freeze_panes == expected, (
        f"Plan freeze_panes is '{plan.freeze_panes}', expected '{expected}'"
    )


def test_freeze_panes_actuals_at_first_wp(workbook):
    """Actuals freeze_panes row matches the first WP data row."""
    ws = workbook["Actuals"]
    fp = ws.freeze_panes
    assert fp is not None
    _, freeze_row = _parse_freeze_panes(fp)
    # WP rows have cross-sheet formulas like =Config!A{row} in column A
    for r in range(1, ws.max_row + 1):
        val = ws.cell(r, 1).value
        if isinstance(val, str) and val.startswith("=Config!A"):
            assert freeze_row == r, (
                f"Actuals freeze row {freeze_row} != first WP row {r}"
            )
            return
    pytest.fail("Could not find first WP formula row in Actuals column A")


def test_freeze_panes_summary_at_first_wp(workbook):
    """Summary freeze_panes row matches the first WP data row."""
    ws = workbook["Summary"]
    fp = ws.freeze_panes
    assert fp is not None
    _, freeze_row = _parse_freeze_panes(fp)
    # Summary WP rows have formulas referencing Config WP IDs in column A
    for r in range(1, ws.max_row + 1):
        val = ws.cell(r, 1).value
        if val is not None and isinstance(val, str) and val.startswith("=Config!A"):
            assert freeze_row == r, (
                f"Summary freeze row {freeze_row} != first WP row {r}"
            )
            return
    pytest.fail("Could not find first WP formula row in Summary column A")
