"""Tests that no builder imports config.py at module level."""
import ast
import inspect
from pathlib import Path

import pytest


BUILDER_MODULES = [
    "budget_planner.plan",
    "budget_planner.config_sheet",
    "budget_planner.capacity",
    "budget_planner.actuals",
    "budget_planner.summary",
    "budget_planner.trends",
    "budget_planner.meta",
]


def test_no_module_level_config_import():
    """No builder should import from budget_planner.config at module level."""
    src_dir = Path(__file__).parent.parent / "budget_planner"
    for module_name in BUILDER_MODULES:
        filename = module_name.split(".")[-1] + ".py"
        filepath = src_dir / filename
        tree = ast.parse(filepath.read_text())

        for node in ast.iter_child_nodes(tree):
            if isinstance(node, ast.ImportFrom):
                if node.module and "budget_planner.config" in node.module:
                    pytest.fail(
                        f"{module_name} has module-level import from "
                        f"budget_planner.config (line {node.lineno})"
                    )
            elif isinstance(node, ast.Import):
                for alias in node.names:
                    if "budget_planner.config" in alias.name:
                        pytest.fail(
                            f"{module_name} has module-level import of "
                            f"budget_planner.config (line {node.lineno})"
                        )


def test_builders_accept_project_params():
    """All builders that need config should accept project_params kwarg."""
    from budget_planner.plan import build_plan_sheet
    from budget_planner.capacity import build_capacity_sheet
    from budget_planner.actuals import build_actuals_sheet
    from budget_planner.summary import build_summary_sheet
    from budget_planner.trends import build_trends_sheet
    from budget_planner.meta import build_meta_sheet

    for func in [build_plan_sheet, build_capacity_sheet, build_actuals_sheet,
                 build_summary_sheet, build_trends_sheet, build_meta_sheet]:
        sig = inspect.signature(func)
        assert "project_params" in sig.parameters, \
            f"{func.__name__} missing project_params parameter"


def test_builders_accept_work_packages():
    """All builders should accept work_packages kwarg."""
    from budget_planner.plan import build_plan_sheet
    from budget_planner.config_sheet import build_config_sheet
    from budget_planner.capacity import build_capacity_sheet
    from budget_planner.actuals import build_actuals_sheet
    from budget_planner.summary import build_summary_sheet
    from budget_planner.trends import build_trends_sheet
    from budget_planner.meta import build_meta_sheet

    for func in [build_plan_sheet, build_config_sheet, build_capacity_sheet,
                 build_actuals_sheet, build_summary_sheet, build_trends_sheet,
                 build_meta_sheet]:
        sig = inspect.signature(func)
        assert "work_packages" in sig.parameters, \
            f"{func.__name__} missing work_packages parameter"



def test_actuals_accepts_preserved():
    """Actuals builder should accept preserved kwarg."""
    from budget_planner.actuals import build_actuals_sheet
    sig = inspect.signature(build_actuals_sheet)
    assert "preserved" in sig.parameters


def test_meta_accepts_rebuild_warnings():
    """Meta builder should accept rebuild_warnings kwarg."""
    from budget_planner.meta import build_meta_sheet
    sig = inspect.signature(build_meta_sheet)
    assert "rebuild_warnings" in sig.parameters
