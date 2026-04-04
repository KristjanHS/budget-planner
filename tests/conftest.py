"""Shared fixtures for programme budget planner tests."""
import pytest

from budget_planner.workbook import init_workbook


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



