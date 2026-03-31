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
    # WP1 (index 0) is active M1-M60, should have non-None values
    wp1_values = [v for v in data["planned"][0] if v is not None]
    assert len(wp1_values) > 0, "WP1 planned values should not all be None"


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


def test_update_preserves_manual_planned_edits(tmp_path):
    """Manual edits to planned FTE should survive regeneration."""
    path = tmp_path / "edited.xlsx"
    generate_workbook(str(path))

    # Edit a planned value
    wb = openpyxl.load_workbook(str(path))
    plan = wb["Plan"]
    layout = discover_layout(plan)
    row = layout["planned_first_wp_row"]
    plan.cell(row, 3, 0.99)  # WP1, M1 = custom value
    wb.save(str(path))

    # Regenerate
    updated = str(tmp_path / "updated.xlsx")
    generate_workbook(output_path=updated, existing_path=str(path))

    wb2 = openpyxl.load_workbook(updated)
    plan2 = wb2["Plan"]
    layout2 = discover_layout(plan2)
    row2 = layout2["planned_first_wp_row"]
    assert plan2.cell(row2, 3).value == pytest.approx(0.99), "Planned edit not preserved"
    wb2.close()


def test_fresh_generation_has_even_distribution(fresh_wb_path):
    """Fresh workbook should have evenly distributed planned FTE."""
    data = read_workbook(fresh_wb_path)
    # WP1: 12.0 FTE over 60 months = 0.2 per month
    wp1_active = [v for v in data["planned"][0] if v is not None and v > 0]
    assert len(wp1_active) == 60
    assert all(abs(v - 0.2) < 0.001 for v in wp1_active)
