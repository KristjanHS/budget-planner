"""Shared fixtures for programme budget planner tests."""
import json
import subprocess

import pytest
from pathlib import Path

from budget_planner.workbook import generate_workbook

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


