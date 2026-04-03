"""Shared fixtures for programme budget planner tests."""
import pytest

from budget_planner.workbook import init_workbook


@pytest.fixture(scope="session")
def fresh_wb_path(tmp_path_factory):
    """Generate a fresh workbook (empty Capacity) for testing."""
    out = tmp_path_factory.mktemp("wb") / "test_budget.xlsx"
    init_workbook(str(out))
    return str(out)



