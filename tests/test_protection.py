"""Tests for cell protection — blue input cells unlocked, others locked."""
import pytest
from openpyxl import load_workbook
from openpyxl.worksheet.protection import SheetProtection
from budget_planner.workbook import init_workbook

pytestmark = pytest.mark.skip("Protection temporarily disabled")


@pytest.fixture(scope="module")
def protection_wb(tmp_path_factory):
    """Generate a single workbook shared by all protection tests (read-only)."""
    path = str(tmp_path_factory.mktemp("protection") / "test.xlsx")
    init_workbook(output_path=path)
    wb = load_workbook(path)
    yield wb
    wb.close()


def test_sheets_are_protected(protection_wb):
    """All sheets should have protection enabled."""
    for ws in protection_wb.worksheets:
        assert ws.protection.sheet, f"Sheet '{ws.title}' is not protected"


def test_blue_input_cells_unlocked(protection_wb):
    """Blue input cells should have locked=False."""
    # Check Config sheet — param values (B2-B4) should be unlocked
    # B1 uses TITLE_FONT (not blue input), so it stays locked
    config = protection_wb["Config"]
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


def test_formula_cells_locked(protection_wb):
    """Formula cells (black/green font) should remain locked."""
    plan = protection_wb["Plan"]
    # Personnel Ceiling (B4) is a formula — should be locked
    cell = plan.cell(4, 2)
    assert cell.protection.locked is not False, \
        "Plan!B4 (formula) should be locked"

    # PLANNED FTE cells are overwritten by _apply_planned_formulas with
    # BLACK_FORMULA SUM formulas — they must be re-locked
    from budget_planner.reader import discover_layout
    layout = discover_layout(plan)
    first_wp = layout["planned_first_wp_row"]
    planned_cell = plan.cell(first_wp, 3)  # Col C = first month
    assert planned_cell.protection.locked is not False, \
        f"Planned FTE cell (row {first_wp}, col C) should be locked"


def test_team_roster_unlocked(protection_wb):
    """Team roster name and max_fte cells (in Config sheet) should be unlocked."""
    config = protection_wb["Config"]
    # Find TEAM ROSTER section
    roster_row = None
    for r in range(1, config.max_row + 1):
        val = config.cell(r, 1).value
        if val and "TEAM ROSTER" in str(val).upper():
            roster_row = r
            break
    assert roster_row is not None, "TEAM ROSTER section not found in Config"
    # First data row is 2 below section header (header + column headers)
    first_data = roster_row + 2
    name_cell = config.cell(first_data, 1)
    fte_cell = config.cell(first_data, 2)
    assert name_cell.protection.locked is False, "Roster name should be unlocked"
    assert fte_cell.protection.locked is False, "Roster max_fte should be unlocked"


def test_capacity_monthly_cells_unlocked(protection_wb):
    """Capacity monthly FTE cells should be unlocked."""
    cap = protection_wb["Capacity"]
    # Find first person's first WP row with a numeric value in col C
    for r in range(1, cap.max_row + 1):
        val = cap.cell(r, 3).value
        if isinstance(val, (int, float)) and 0 < val < 1:
            assert cap.cell(r, 3).protection.locked is False, \
                f"Capacity FTE cell at row {r} should be unlocked"
            break


def test_protection_allows_column_row_resizing(protection_wb):
    """Users must be able to resize columns and rows on all sheets.

    openpyxl defaults formatColumns/formatRows to True (restricted).
    We explicitly set them to False so users can auto-fit column widths.
    """
    for ws in protection_wb.worksheets:
        assert ws.protection.formatColumns is False, \
            f"Sheet '{ws.title}': formatColumns should be False (allow resizing)"
        assert ws.protection.formatRows is False, \
            f"Sheet '{ws.title}': formatRows should be False (allow resizing)"


def test_protection_password(protection_wb):
    """Protection password should be 'grant' on every sheet."""
    # Compute the expected hash for "grant"
    expected = SheetProtection()
    expected.set_password("grant")
    expected_hash = expected.password

    for ws in protection_wb.worksheets:
        assert ws.protection.password == expected_hash, \
            f"Sheet '{ws.title}' password hash {ws.protection.password!r} " \
            f"does not match expected hash {expected_hash!r} for password 'grant'"
