"""Tests for CLI modes, snapshot, and restore functionality."""

import filecmp

import pytest
import openpyxl

from budget_planner.workbook import generate_workbook, snapshot_workbook, restore_workbook
from budget_planner.config import (
    TEAM, WORK_PACKAGES, EXTERNAL_COMMITMENTS,
    MONTH_COL_START, PROJECT_MONTHS,
)


# ── Helpers ──

def _first_active_col(wp):
    """Return the column index of the first active month for a WP."""
    return MONTH_COL_START + wp["start"] - 1


# ── Test 1: Default mode produces demo workbook with seed values ──

def test_cli_default_is_demo_mode(tmp_path):
    """Running generate_workbook() without mode produces Capacity with values > 0."""
    out = tmp_path / "demo.xlsx"
    generate_workbook(str(out))
    wb = openpyxl.load_workbook(str(out), data_only=False)
    cap = wb["Capacity"]

    # Check that at least one active month cell has a numeric value > 0.
    # Scan all cells in the month columns for any numeric value > 0.
    found_value = False
    for row in range(1, cap.max_row + 1):
        for mi in range(PROJECT_MONTHS):
            c = MONTH_COL_START + mi
            cell_val = cap.cell(row=row, column=c).value
            if isinstance(cell_val, (int, float)) and cell_val > 0:
                found_value = True
                break
        if found_value:
            break

    assert found_value, "Demo mode should have at least one Capacity cell with value > 0"


# ── Test 2: Skeleton mode produces empty monthly cells ──

def test_cli_skeleton_mode(tmp_path):
    """generate_workbook(mode='skeleton') produces empty Capacity monthly cells."""
    out = tmp_path / "skeleton.xlsx"
    generate_workbook(str(out), mode="skeleton")
    wb = openpyxl.load_workbook(str(out), data_only=False)
    cap = wb["Capacity"]

    # Check that ALL active month cells in programme WP rows and external rows are empty
    for member in TEAM:
        for wp_id in member["wps"]:
            wp = next(w for w in WORK_PACKAGES if w["id"] == wp_id)
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


# ── Test 3: Skeleton mode has cell comments ──

def test_cli_skeleton_has_cell_comments(tmp_path):
    """First editable cell in first person's WP row has a comment."""
    out = tmp_path / "skeleton_comment.xlsx"
    generate_workbook(str(out), mode="skeleton")
    wb = openpyxl.load_workbook(str(out))
    cap = wb["Capacity"]

    # For each person, find their first WP row and check comment on first active cell
    for member in TEAM:
        first_wp_id = list(member["wps"].keys())[0]
        wp = next(w for w in WORK_PACKAGES if w["id"] == first_wp_id)
        first_active_col = _first_active_col(wp)

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


# ── Test 4: Snapshot creates copy ──

def test_cli_snapshot_creates_copy(tmp_path):
    """snapshot_workbook creates <stem>.snapshot.xlsx."""
    src = tmp_path / "budget.xlsx"
    generate_workbook(str(src))

    result = snapshot_workbook(str(src))
    expected = tmp_path / "budget.snapshot.xlsx"
    assert expected.exists()
    assert result == str(expected)


# ── Test 5: Restore overwrites from snapshot ──

def test_cli_restore_overwrites_from_snapshot(tmp_path):
    """restore_workbook copies snapshot back to file."""
    src = tmp_path / "budget.xlsx"
    generate_workbook(str(src))

    # Create snapshot
    snapshot_workbook(str(src))
    snapshot_path = tmp_path / "budget.snapshot.xlsx"

    # Read original content for comparison
    original_bytes = src.read_bytes()

    # Overwrite the source file with different content
    generate_workbook(str(src), mode="skeleton")
    modified_bytes = src.read_bytes()
    assert original_bytes != modified_bytes, "File should differ after regeneration"

    # Restore
    result = restore_workbook(str(src))
    assert result == str(src)

    # Restored file should match original snapshot
    assert filecmp.cmp(str(src), str(snapshot_path), shallow=False)


# ── Test 7: Restore creates timestamped backup ──

def test_cli_restore_creates_timestamped_backup(tmp_path):
    """Before restoring, current file backed up with timestamp."""
    src = tmp_path / "budget.xlsx"
    generate_workbook(str(src))
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
    generate_workbook(str(src))

    with pytest.raises(FileNotFoundError, match="No snapshot found"):
        restore_workbook(str(src))
