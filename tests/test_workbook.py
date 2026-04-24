"""Structural checks for the generated workbook — maps from verify.py."""
import re

import pytest
import openpyxl

from budget_planner.reader import discover_layout
from budget_planner.config import WORK_PACKAGES, NUM_WPS, MONTH_COL_START, PROJECT_MONTHS
from budget_planner.layout import wp_month_bounds
from budget_planner.workbook import init_workbook


# ── Fixtures that don't need LibreOffice ──

@pytest.fixture(scope="module")
def workbook(fresh_wb_path):
    return openpyxl.load_workbook(fresh_wb_path, data_only=False)


@pytest.fixture(scope="module")
def layout(workbook):
    return discover_layout(workbook["Plan"])


# ── Check 2-3: Sheet existence + global params ──

def test_sheets_exist(workbook):
    for name in ["Config", "Plan", "Capacity", "Actuals",
                 "Actuals-Hours", "Plan-Forecast", "Summary"]:
        assert name in workbook.sheetnames, f"Missing sheet: {name}"


def test_global_params(workbook):
    # Plan compact header references Config sheet via formulas
    plan = workbook["Plan"]
    # B1 = title ref, E1 = start ref, G1 = ceiling ref
    assert plan["B1"].value is not None, "Plan!B1 (title ref) is empty"
    assert plan["E1"].value is not None, "Plan!E1 (start ref) is empty"
    assert plan["G1"].value is not None, "Plan!G1 (ceiling ref) is empty"
    # Check Config directly for valid rate (find by label, not hardcoded row)
    # FTE Rate is now a formula (weighted average of grade rates), not a static value
    config = workbook["Config"]
    rate_val = None
    for r in range(1, config.max_row + 1):
        if "fte rate" in str(config.cell(r, 1).value or "").lower():
            rate_val = config.cell(r, 2).value
            break
    assert rate_val is not None, "Config FTE Rate cell is empty"
    from openpyxl.worksheet.formula import ArrayFormula
    assert isinstance(rate_val, (str, ArrayFormula)), (
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
    assert defined.get("partner_ceiling") == "Plan!$G$1"


# ── Check 5: Config table ──

@pytest.mark.parametrize("wp_idx", [0, 2])
def test_wp_config_valid(workbook, layout, wp_idx):
    config = workbook["Config"]
    row = layout["config_first_wp_row"] + wp_idx
    wp_id = config.cell(row, 1).value
    fte = config.cell(row, 3).value
    rate = config.cell(row, 4).value
    assert wp_id is not None, f"WP ID is empty at row {row}"
    assert fte is not None and fte > 0, f"{wp_id}: FTE invalid ({fte})"
    assert rate is not None and rate > 0, f"{wp_id}: Rate invalid ({rate})"


# ── Check 11: Inactive months empty ──

def _wp_indices_with_inactive_months():
    """Find WP indices that have inactive months (start > 1 or end < PROJECT_MONTHS)."""
    result = []
    for i, wp in enumerate(WORK_PACKAGES):
        start, end = wp_month_bounds(wp["id"], WORK_PACKAGES)
        if start > 1 or end < PROJECT_MONTHS:
            result.append(i)
        if len(result) >= 3:
            break
    return result


@pytest.mark.parametrize("wp_idx", _wp_indices_with_inactive_months())
def test_inactive_months_empty(workbook, layout, wp_idx):
    plan = workbook["Plan"]
    wp = WORK_PACKAGES[wp_idx]
    start, end = wp_month_bounds(wp["id"], WORK_PACKAGES)
    row = layout["planned_first_wp_row"] + wp_idx
    for m in range(PROJECT_MONTHS):
        month_num = m + 1
        col = MONTH_COL_START + m
        if month_num < start or month_num > end:
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


# ── Config assignments round-trip ──


class TestConfigAssignments:
    """Verify assignments and external commitments are written to Config
    and survive a rebuild round-trip."""

    @pytest.fixture(scope="class")
    def demo_wb(self, assigned_wb_path):
        return openpyxl.load_workbook(assigned_wb_path, data_only=False)

    def test_init_writes_config_assignments(self, demo_wb):
        """Config sheet has a PERSON-WP ASSIGNMENTS section with data."""
        from budget_planner.reader import read_assignments_from_config
        from budget_planner.config import DEMO_CAPACITY_ASSIGNMENTS

        assignments = read_assignments_from_config(demo_wb["Config"])
        assert len(assignments) == len(DEMO_CAPACITY_ASSIGNMENTS)
        # Spot-check first entry
        first = DEMO_CAPACITY_ASSIGNMENTS[0]
        match = [a for a in assignments
                 if a["person"] == first["person"] and a["wp"] == first["wp"]]
        assert match, f"Expected assignment for {first['person']}/{first['wp']}"
        assert match[0]["total_fte"] == pytest.approx(first["total_fte"])

    def test_init_writes_config_externals(self, demo_wb):
        """Config sheet has an EXTERNAL COMMITMENTS section with data."""
        from budget_planner.reader import read_externals_from_config
        from budget_planner.config import EXTERNAL_COMMITMENTS

        externals = read_externals_from_config(demo_wb["Config"])
        assert len(externals) == len(EXTERNAL_COMMITMENTS)
        first = EXTERNAL_COMMITMENTS[0]
        match = [e for e in externals
                 if e["person"] == first["person"] and e["project"] == first["project"]]
        assert match, f"Expected external for {first['person']}/{first['project']}"
        assert match[0]["fte"] == pytest.approx(first["fte"])

    def test_rebuild_preserves_config_assignments(self, tmp_path_factory):
        """Rebuild reads assignments from Config and preserves them."""
        from budget_planner.config import DEMO_CAPACITY_ASSIGNMENTS, EXTERNAL_COMMITMENTS
        from budget_planner.reader import read_assignments_from_config, read_externals_from_config
        from budget_planner.workbook import rebuild_workbook

        out = tmp_path_factory.mktemp("wb_rebuild") / "test_rebuild.xlsx"
        init_workbook(str(out), assignments=DEMO_CAPACITY_ASSIGNMENTS)
        rebuild_workbook(str(out))

        wb = openpyxl.load_workbook(str(out), data_only=False)
        assignments = read_assignments_from_config(wb["Config"])
        externals = read_externals_from_config(wb["Config"])

        assert len(assignments) == len(DEMO_CAPACITY_ASSIGNMENTS)
        assert len(externals) == len(EXTERNAL_COMMITMENTS)

        # Verify data integrity after round-trip
        for orig in DEMO_CAPACITY_ASSIGNMENTS:
            match = [a for a in assignments
                     if a["person"] == orig["person"] and a["wp"] == orig["wp"]]
            assert match, f"Lost assignment {orig['person']}/{orig['wp']} after rebuild"
            assert match[0]["total_fte"] == pytest.approx(orig["total_fte"])


# ── Build-order enforcement ──

def test_build_order_is_stable(tmp_path):
    """Change detector: asserts _build_all_sheets calls builders in the order
    documented in the workbook.py module docstring.

    If this test fails, the orchestrator has been reordered. Read the
    build-order DAG in workbook.py before "fixing" this test — downstream
    builders consume layouts produced by upstream ones, so reordering
    silently breaks cross-sheet formulas.
    """
    from unittest.mock import patch
    from budget_planner import workbook as wb_mod

    call_order: list[str] = []

    def _spy(name, real):
        def _wrapped(*args, **kwargs):
            call_order.append(name)
            return real(*args, **kwargs)
        return _wrapped

    targets = [
        "build_config_sheet",
        "build_capacity_sheet",
        "build_plan_sheet",
        "_apply_planned_formulas",
        "build_actuals_hours_sheet",
        "build_plan_forecast_sheet",
        "build_actuals_sheet",
        "build_summary_sheet",
        "build_trends_sheet",
    ]
    with (
        patch.object(wb_mod, "build_config_sheet",
                     _spy("build_config_sheet", wb_mod.build_config_sheet)),
        patch.object(wb_mod, "build_plan_sheet",
                     _spy("build_plan_sheet", wb_mod.build_plan_sheet)),
        patch.object(wb_mod, "build_capacity_sheet",
                     _spy("build_capacity_sheet", wb_mod.build_capacity_sheet)),
        patch.object(wb_mod, "_apply_planned_formulas",
                     _spy("_apply_planned_formulas", wb_mod._apply_planned_formulas)),
        patch.object(wb_mod, "build_actuals_sheet",
                     _spy("build_actuals_sheet", wb_mod.build_actuals_sheet)),
        patch.object(wb_mod, "build_actuals_hours_sheet",
                     _spy("build_actuals_hours_sheet",
                          wb_mod.build_actuals_hours_sheet)),
        patch.object(wb_mod, "build_plan_forecast_sheet",
                     _spy("build_plan_forecast_sheet",
                          wb_mod.build_plan_forecast_sheet)),
        patch.object(wb_mod, "build_summary_sheet",
                     _spy("build_summary_sheet", wb_mod.build_summary_sheet)),
        patch.object(wb_mod, "build_trends_sheet",
                     _spy("build_trends_sheet", wb_mod.build_trends_sheet)),
    ):
        wb_mod.init_workbook(str(tmp_path / "order.xlsx"))

    assert call_order == targets, (
        "Builder call order changed. See workbook.py module docstring "
        "for the DAG before adjusting — downstream builders consume "
        "upstream layout dicts."
    )
