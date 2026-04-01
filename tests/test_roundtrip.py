"""Round-trip tests: reader + update preserves user data."""
import openpyxl
import pytest

from budget_planner.workbook import generate_workbook
from budget_planner.reader import discover_layout, read_workbook


def test_discover_layout_finds_all_blocks(fresh_wb_path):
    wb = openpyxl.load_workbook(fresh_wb_path)
    layout = discover_layout(wb["Plan"])
    expected_keys = [
        "config_header_row", "config_first_wp_row", "config_last_wp_row",
        "config_total_row", "planned_header_row", "planned_first_wp_row",
        "planned_last_wp_row", "planned_total_row", "actual_header_row",
        "actual_first_wp_row", "actual_last_wp_row", "actual_total_row",
    ]
    for key in expected_keys:
        assert key in layout, f"Missing layout key: {key}"
    wb.close()


def test_read_workbook_extracts_planned_values(fresh_wb_path):
    data = read_workbook(fresh_wb_path)
    assert "planned" in data
    assert len(data["planned"]) > 0
    # v2: planned cells are formulas, reader returns None for formula cells
    wp1_values = [v for v in data["planned"][0] if v is not None]
    assert len(wp1_values) == 0, "Planned formula cells should read as None"


def test_read_workbook_extracts_actual_values(wb_with_actuals):
    data = read_workbook(wb_with_actuals)
    assert "actual" in data
    # We injected 0.15 at M1 and 0.20 at M2 for WP1
    assert data["actual"][0][0] == pytest.approx(0.15)
    assert data["actual"][0][1] == pytest.approx(0.20)


def test_update_preserves_actuals(wb_with_actuals, tmp_path):
    """Regenerate with --update and verify actuals survive."""
    updated_path = str(tmp_path / "updated.xlsx")
    generate_workbook(output_path=updated_path, existing_path=wb_with_actuals)

    wb = openpyxl.load_workbook(updated_path)
    plan = wb["Plan"]
    layout = discover_layout(plan)
    row = layout["actual_first_wp_row"]
    assert plan.cell(row, 3).value == pytest.approx(0.15), "Actual M1 not preserved"
    assert plan.cell(row, 4).value == pytest.approx(0.20), "Actual M2 not preserved"
    wb.close()


