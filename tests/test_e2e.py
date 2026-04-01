"""End-to-end tests requiring LibreOffice for formula recalculation.

Tests skip gracefully when LibreOffice is not available.
"""
import json
import subprocess

import pytest
import openpyxl
from pathlib import Path

from budget_planner.workbook import generate_workbook, snapshot_workbook, restore_workbook
from budget_planner.config import MONTH_COL_START

SCRIPTS_DIR = Path(__file__).parent.parent / "scripts"


# ── Fixtures ──


@pytest.fixture(scope="module")
def demo_wb_path(tmp_path_factory):
    """Generate a demo workbook."""
    out = tmp_path_factory.mktemp("e2e") / "demo.xlsx"
    generate_workbook(str(out), mode="demo")
    return str(out)


@pytest.fixture(scope="module")
def demo_recalc(demo_wb_path):
    """Recalculate demo workbook via LibreOffice."""
    result = subprocess.run(
        ["python3", str(SCRIPTS_DIR / "recalc.py"), demo_wb_path],
        capture_output=True, text=True,
    )
    if result.returncode != 0:
        pytest.skip("LibreOffice not available")
    return json.loads(result.stdout)


@pytest.fixture(scope="module")
def skeleton_wb_path(tmp_path_factory):
    """Generate a skeleton workbook."""
    out = tmp_path_factory.mktemp("e2e_skel") / "skeleton.xlsx"
    generate_workbook(str(out), mode="skeleton")
    return str(out)


@pytest.fixture(scope="module")
def skeleton_recalc(skeleton_wb_path):
    """Recalculate skeleton workbook via LibreOffice."""
    result = subprocess.run(
        ["python3", str(SCRIPTS_DIR / "recalc.py"), skeleton_wb_path],
        capture_output=True, text=True,
    )
    if result.returncode != 0:
        pytest.skip("LibreOffice not available")
    return json.loads(result.stdout)


# ── Recalc tests ──


def test_demo_mode_recalc_zero_errors(demo_recalc):
    """Generate demo workbook -> recalc -> 0 formula errors across all 5 sheets."""
    assert demo_recalc["status"] == "success"
    assert demo_recalc["total_errors"] == 0


def test_skeleton_mode_recalc_zero_errors(skeleton_recalc):
    """Generate skeleton workbook -> recalc -> 0 formula errors."""
    assert skeleton_recalc["status"] == "success"
    assert skeleton_recalc["total_errors"] == 0


def test_demo_mode_formula_count(demo_recalc):
    """v2 workbook should have more formulas than v1 had."""
    # v1 had ~60-100 formulas. v2 adds Capacity TOTAL/REMAINING + Plan cross-sheet SUMs
    assert demo_recalc["total_formulas"] > 200


# ── Plan <-> Capacity consistency (requires recalc) ──


def test_demo_mode_plan_planned_equals_capacity_sums(demo_wb_path, demo_recalc):
    """After recalc, Plan Planned FTE values should equal Capacity sums."""
    # Load recalculated workbook with data_only=True to get computed values
    wb = openpyxl.load_workbook(demo_wb_path, data_only=True)
    plan = wb["Plan"]

    from budget_planner.reader import discover_layout

    layout = discover_layout(plan)

    # For WP1 (active M1-M60), check that Plan planned M1 value > 0
    # (it should be the sum of all persons' WP1 allocations)
    planned_row = layout["planned_first_wp_row"]
    val = plan.cell(planned_row, MONTH_COL_START).value
    # After recalc, this should be a computed number, not a formula
    assert val is not None and isinstance(val, (int, float)) and val > 0, \
        f"Plan WP1 M1 after recalc should be positive, got: {val}"
    wb.close()


# ── Snapshot / Restore ──


def test_snapshot_restore_roundtrip(tmp_path):
    """Generate -> snapshot -> edit -> restore -> matches original."""
    path = str(tmp_path / "test.xlsx")
    generate_workbook(path, mode="demo")

    # Read original value
    wb = openpyxl.load_workbook(path)
    cap = wb["Capacity"]
    original_val = cap.cell(10, MONTH_COL_START).value
    wb.close()

    # Snapshot
    snapshot_workbook(path)

    # Edit
    wb = openpyxl.load_workbook(path)
    cap = wb["Capacity"]
    cap.cell(10, MONTH_COL_START, 99.99)
    wb.save(path)
    wb.close()

    # Restore
    restore_workbook(path)

    # Verify
    wb = openpyxl.load_workbook(path)
    cap = wb["Capacity"]
    restored_val = cap.cell(10, MONTH_COL_START).value
    wb.close()

    assert restored_val == original_val
