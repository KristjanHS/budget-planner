"""Tests that no builder imports config.py at module level."""
import ast
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


