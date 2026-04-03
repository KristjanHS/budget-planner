"""Shared fixtures for programme budget planner tests."""
import pytest

from budget_planner.workbook import generate_workbook


@pytest.fixture(scope="session")
def fresh_wb_path(tmp_path_factory):
    """Generate a fresh demo workbook for testing."""
    out = tmp_path_factory.mktemp("wb") / "test_budget.xlsx"
    generate_workbook(str(out))
    return str(out)


@pytest.fixture(scope="session")
def skeleton_wb_path(tmp_path_factory):
    """Generate a fresh skeleton workbook for testing."""
    out = tmp_path_factory.mktemp("wb_skeleton") / "test_budget_skeleton.xlsx"
    generate_workbook(str(out), mode="skeleton")
    return str(out)



