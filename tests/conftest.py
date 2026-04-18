"""Shared fixtures for programme budget planner tests."""
import json
import subprocess
from pathlib import Path

import pytest

from budget_planner.workbook import init_workbook, seed_demo_data

SCRIPTS_DIR = Path(__file__).parent.parent / "scripts"


@pytest.fixture(scope="session")
def fresh_wb_path(tmp_path_factory):
    """Generate a fresh workbook (empty Capacity — no WP assignments)."""
    out = tmp_path_factory.mktemp("wb") / "test_budget.xlsx"
    init_workbook(str(out))
    return str(out)


@pytest.fixture(scope="session")
def assigned_wb_path(tmp_path_factory):
    """Generate a workbook with demo assignments (WP rows per person)."""
    from budget_planner.config import DEMO_CAPACITY_ASSIGNMENTS
    out = tmp_path_factory.mktemp("wb_assigned") / "test_assigned.xlsx"
    init_workbook(str(out), assignments=DEMO_CAPACITY_ASSIGNMENTS)
    return str(out)


@pytest.fixture(scope="session")
def demo_wb_path(tmp_path_factory):
    """Generate a demo workbook (init with assignments + seed_demo_data).

    Read-only — tests that consume this fixture MUST NOT save or rebuild.
    """
    from budget_planner.config import DEMO_CAPACITY_ASSIGNMENTS
    out = tmp_path_factory.mktemp("demo_wb") / "demo.xlsx"
    init_workbook(str(out), assignments=DEMO_CAPACITY_ASSIGNMENTS)
    seed_demo_data(str(out))
    return str(out)


@pytest.fixture(scope="session")
def demo_recalc(demo_wb_path):
    """Recalculate demo workbook via LibreOffice (skips if LO unavailable)."""
    result = subprocess.run(
        ["python3", str(SCRIPTS_DIR / "recalc.py"), demo_wb_path],
        capture_output=True, text=True,
    )
    if result.returncode != 0:
        pytest.skip("LibreOffice not available")
    return json.loads(result.stdout)


@pytest.fixture(scope="session")
def demo_recalc_path(demo_wb_path, demo_recalc):
    """Path to the demo workbook AFTER session recalc ran.

    Depends on demo_recalc to enforce ordering: the recalc subprocess
    must have mutated demo_wb_path before callers open it with
    data_only=True. Returns the same path as demo_wb_path, just with
    a dependency edge that guarantees recalc has run.

    Read-only — consumers MUST NOT save or rebuild.
    """
    del demo_recalc  # consumed for ordering only; ensures recalc ran before return
    return demo_wb_path
