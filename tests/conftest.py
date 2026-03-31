"""Shared fixtures for programme budget planner tests."""
import json
import subprocess

import pytest
import openpyxl
from pathlib import Path

from budget_planner.workbook import generate_workbook
from budget_planner.reader import discover_layout

SCRIPTS_DIR = Path(__file__).parent.parent / "scripts"


@pytest.fixture(scope="session")
def fresh_wb_path(tmp_path_factory):
    """Generate a fresh workbook for testing."""
    out = tmp_path_factory.mktemp("wb") / "test_budget.xlsx"
    generate_workbook(str(out))
    return str(out)


@pytest.fixture(scope="session")
def recalc_result(fresh_wb_path):
    """Recalculate formulas via LibreOffice. Skips if unavailable."""
    result = subprocess.run(
        ["python3", str(SCRIPTS_DIR / "recalc.py"), fresh_wb_path],
        capture_output=True, text=True
    )
    if result.returncode != 0:
        pytest.skip("LibreOffice not available")
    return json.loads(result.stdout)


@pytest.fixture
def wb_with_actuals(tmp_path):
    """Workbook with manually injected actual values for round-trip testing."""
    path = tmp_path / "with_actuals.xlsx"
    generate_workbook(str(path))
    wb = openpyxl.load_workbook(str(path))
    plan = wb["Plan"]
    layout = discover_layout(plan)
    row = layout["actual_first_wp_row"]
    plan.cell(row, 3, 0.15)   # WP1, M1
    plan.cell(row, 4, 0.20)   # WP1, M2
    wb.save(str(path))
    return str(path)
