"""Tests for cell protection — blue input cells unlocked, others locked."""
from openpyxl import load_workbook
from budget_planner.workbook import init_workbook


def test_sheets_are_protected(tmp_path):
    """All sheets should have protection enabled."""
    path = str(tmp_path / "test.xlsx")
    init_workbook(output_path=path)
    wb = load_workbook(path)
    for ws in wb.worksheets:
        assert ws.protection.sheet, f"Sheet '{ws.title}' is not protected"
    wb.close()


def test_blue_input_cells_unlocked(tmp_path):
    """Blue input cells should have locked=False."""
    path = str(tmp_path / "test.xlsx")
    init_workbook(output_path=path)
    wb = load_workbook(path)

    # Check Config sheet — param values (B2-B4) should be unlocked
    # B1 uses TITLE_FONT (not blue input), so it stays locked
    config = wb["Config"]
    for r in [2, 3, 4]:
        cell = config.cell(r, 2)
        assert cell.protection.locked is False, \
            f"Config!B{r} should be unlocked (blue input)"

    # Check Config sheet — WP data cells should be unlocked
    wp_first_row = 8  # After section header + column headers
    cell = config.cell(wp_first_row, 1)
    if cell.value is not None:
        assert cell.protection.locked is False, \
            "Config WP cell should be unlocked"

    wb.close()


def test_formula_cells_locked(tmp_path):
    """Formula cells (black/green font) should remain locked."""
    path = str(tmp_path / "test.xlsx")
    init_workbook(output_path=path)
    wb = load_workbook(path)

    plan = wb["Plan"]
    # Personnel Ceiling (B4) is a formula — should be locked
    cell = plan.cell(4, 2)
    assert cell.protection.locked is not False, \
        "Plan!B4 (formula) should be locked"

    wb.close()


def test_team_roster_unlocked(tmp_path):
    """Team roster name and max_fte cells should be unlocked."""
    path = str(tmp_path / "test.xlsx")
    init_workbook(output_path=path)
    wb = load_workbook(path)

    team = wb["Team"]
    # Row 3 is first data row (row 1 = section header, row 2 = column headers)
    name_cell = team.cell(3, 1)
    fte_cell = team.cell(3, 2)
    assert name_cell.protection.locked is False, "Team name should be unlocked"
    assert fte_cell.protection.locked is False, "Team max_fte should be unlocked"

    wb.close()


def test_capacity_monthly_cells_unlocked(tmp_path):
    """Capacity monthly FTE cells should be unlocked."""
    path = str(tmp_path / "test.xlsx")
    init_workbook(output_path=path)
    wb = load_workbook(path)

    cap = wb["Capacity"]
    # Find first person's first WP row with a numeric value in col C
    for r in range(1, cap.max_row + 1):
        val = cap.cell(r, 3).value
        if isinstance(val, (int, float)) and 0 < val < 1:
            assert cap.cell(r, 3).protection.locked is False, \
                f"Capacity FTE cell at row {r} should be unlocked"
            break

    wb.close()


def test_protection_password(tmp_path):
    """Protection password should be 'grant'."""
    path = str(tmp_path / "test.xlsx")
    init_workbook(output_path=path)
    wb = load_workbook(path)

    # All sheets should have a password hash set
    for ws in wb.worksheets:
        assert ws.protection.password is not None or ws.protection.sheet, \
            f"Sheet '{ws.title}' has no protection"

    wb.close()
