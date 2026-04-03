"""Tests for CLI modes, snapshot, and restore functionality."""

import filecmp

import pytest
import openpyxl

from budget_planner.workbook import init_workbook, seed_demo_data, snapshot_workbook, restore_workbook
from budget_planner.config import (
    TEAM, WORK_PACKAGES, EXTERNAL_COMMITMENTS,
    MONTH_COL_START, PROJECT_MONTHS,
)


# ── Test 1: Init produces workbook with expected structure ──

def test_cli_init_produces_workbook(tmp_path):
    """init_workbook() produces workbook with Capacity sheet and empty month cells."""
    out = tmp_path / "init.xlsx"
    init_workbook(str(out))
    wb = openpyxl.load_workbook(str(out), data_only=False)

    # Verify Capacity sheet exists and has person groups
    assert "Capacity" in wb.sheetnames
    cap = wb["Capacity"]
    assert cap.max_row > 1, "Capacity sheet should have content"


# ── Test 2: Init produces empty monthly cells ──

def test_cli_init_has_empty_capacity(tmp_path):
    """init_workbook() produces empty Capacity monthly cells."""
    out = tmp_path / "init_empty.xlsx"
    init_workbook(str(out))
    wb = openpyxl.load_workbook(str(out), data_only=False)
    cap = wb["Capacity"]

    # Check that ALL active month cells in programme WP rows are empty
    for member in TEAM:
        for wp in WORK_PACKAGES:
            wp_id = wp["id"]
            # Find the row for this WP by scanning column A
            for row in range(1, cap.max_row + 1):
                val = cap.cell(row=row, column=1).value
                if isinstance(val, str) and wp_id in val:
                    for mi in range(PROJECT_MONTHS):
                        month_num = mi + 1
                        c = MONTH_COL_START + mi
                        if wp["start"] <= month_num <= wp["end"]:
                            cell_val = cap.cell(row=row, column=c).value
                            assert cell_val is None or cell_val == "", (
                                f"{member['name']}/{wp_id} M{month_num}: "
                                f"expected empty, got {cell_val}"
                            )
                    break

    # Also check external project rows
    for entry in EXTERNAL_COMMITMENTS:
        for row in range(1, cap.max_row + 1):
            val = cap.cell(row=row, column=1).value
            if val == entry["project"]:
                for mi in range(PROJECT_MONTHS):
                    month_num = mi + 1
                    c = MONTH_COL_START + mi
                    if entry["start"] <= month_num <= entry["end"]:
                        cell_val = cap.cell(row=row, column=c).value
                        assert cell_val is None or cell_val == "", (
                            f"{entry['person']}/{entry['project']} M{month_num}: "
                            f"expected empty, got {cell_val}"
                        )
                break


# ── Test 3: Init has cell comments on empty capacity ──

def test_cli_init_has_cell_comments(tmp_path):
    """First editable cell in first person's WP row has a comment."""
    out = tmp_path / "init_comment.xlsx"
    init_workbook(str(out))
    wb = openpyxl.load_workbook(str(out))
    cap = wb["Capacity"]

    # For each person, find their first WP row and check comment on first active cell
    first_wp = WORK_PACKAGES[0]
    first_wp_id = first_wp["id"]
    first_active_col = MONTH_COL_START + first_wp["start"] - 1

    for member in TEAM:
        for row in range(1, cap.max_row + 1):
            val = cap.cell(row=row, column=1).value
            if isinstance(val, str) and first_wp_id in val:
                comment = cap.cell(row=row, column=first_active_col).comment
                assert comment is not None, (
                    f"{member['name']}/{first_wp_id}: expected comment on first "
                    f"active cell (row={row}, col={first_active_col})"
                )
                assert "Enter monthly FTE allocation" in comment.text
                assert comment.author == "Programme Budget Tool"
                break


# ── Test 4: Demo seeding populates Capacity and Actuals ──

def test_cli_demo_seeds_data(tmp_path):
    """seed_demo_data() populates Capacity FTE and Actuals EUR values."""
    out = tmp_path / "demo.xlsx"
    init_workbook(str(out))
    seed_demo_data(str(out))
    wb = openpyxl.load_workbook(str(out), data_only=False)

    # Capacity should have non-zero FTE values
    cap = wb["Capacity"]
    has_fte = False
    for row in range(1, cap.max_row + 1):
        for mi in range(PROJECT_MONTHS):
            c = MONTH_COL_START + mi
            val = cap.cell(row, c).value
            if isinstance(val, (int, float)) and val > 0 and val < 1.0:
                has_fte = True
                break
        if has_fte:
            break
    assert has_fte, "Demo seeding should populate Capacity with FTE values"

    # Actuals should have non-zero EUR values
    act = wb["Actuals"]
    has_eur = False
    for row in range(1, act.max_row + 1):
        for mi in range(PROJECT_MONTHS):
            c = MONTH_COL_START + mi
            val = act.cell(row, c).value
            if isinstance(val, (int, float)) and val > 0:
                has_eur = True
                break
        if has_eur:
            break
    assert has_eur, "Demo seeding should populate Actuals with EUR values"

    # Trends current_month should be set to DEMO_CURRENT_MONTH
    from budget_planner.config import DEMO_CURRENT_MONTH
    assert wb["Trends"]["B1"].value == DEMO_CURRENT_MONTH


# ── Test 5: Snapshot creates copy ──

def test_cli_snapshot_creates_copy(tmp_path):
    """snapshot_workbook creates <stem>.snapshot.xlsx."""
    src = tmp_path / "budget.xlsx"
    init_workbook(str(src))

    result = snapshot_workbook(str(src))
    expected = tmp_path / "budget.snapshot.xlsx"
    assert expected.exists()
    assert result == str(expected)


# ── Test 6: Restore overwrites from snapshot ──

def test_cli_restore_overwrites_from_snapshot(tmp_path):
    """restore_workbook copies snapshot back to file."""
    src = tmp_path / "budget.xlsx"
    init_workbook(str(src))

    # Create snapshot
    snapshot_workbook(str(src))
    snapshot_path = tmp_path / "budget.snapshot.xlsx"

    # Read original content for comparison
    original_bytes = src.read_bytes()

    # Overwrite the source file with different content
    init_workbook(str(src))
    seed_demo_data(str(src))
    modified_bytes = src.read_bytes()
    assert original_bytes != modified_bytes, "File should differ after demo seeding"

    # Restore
    result = restore_workbook(str(src))
    assert result == str(src)

    # Restored file should match original snapshot
    assert filecmp.cmp(str(src), str(snapshot_path), shallow=False)


# ── Test 7: Restore creates timestamped backup ──

def test_cli_restore_creates_timestamped_backup(tmp_path):
    """Before restoring, current file backed up with timestamp."""
    src = tmp_path / "budget.xlsx"
    init_workbook(str(src))
    snapshot_workbook(str(src))

    # Modify the source so it differs from the snapshot
    src.write_bytes(b"modified content placeholder")

    # Count files right before restore (after all setup is done)
    files_before = set(tmp_path.iterdir())

    # Restore
    restore_workbook(str(src))

    # A new timestamped backup file should exist
    files_after = set(tmp_path.iterdir())
    new_files = files_after - files_before
    assert len(new_files) == 1, f"Expected 1 new backup file, got {new_files}"
    backup_file = new_files.pop()
    assert backup_file.suffix == ".xlsx"
    assert "budget." in backup_file.name
    assert backup_file.name not in ("budget.xlsx", "budget.snapshot.xlsx")


# ── Test 8: Restore without snapshot raises error ──

def test_cli_restore_no_snapshot_raises_error(tmp_path):
    """restore without existing snapshot raises FileNotFoundError."""
    src = tmp_path / "budget.xlsx"
    init_workbook(str(src))

    with pytest.raises(FileNotFoundError, match="No snapshot found"):
        restore_workbook(str(src))


# ── Test 9: Init creates timestamped backup when file exists ──

def test_init_creates_backup_when_file_exists(tmp_path):
    """Second init_workbook() to same path creates a timestamped backup."""
    out = tmp_path / "budget.xlsx"
    init_workbook(str(out))  # first generation
    files_before = set(tmp_path.iterdir())
    init_workbook(str(out))  # second generation should create backup
    files_after = set(tmp_path.iterdir())
    new_files = files_after - files_before
    assert len(new_files) == 1, f"Expected 1 new backup file, got {new_files}"
    backup = new_files.pop()
    assert backup.suffix == ".xlsx"
    assert "budget." in backup.name
    assert backup.name != "budget.xlsx"


# ── Test 10: Init creates output directory if missing ──

def test_init_creates_output_directory(tmp_path):
    """init_workbook() creates intermediate directories for output path."""
    out = tmp_path / "subdir" / "deep" / "budget.xlsx"
    init_workbook(str(out))
    assert out.exists()
