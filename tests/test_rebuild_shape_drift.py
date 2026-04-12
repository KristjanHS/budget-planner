"""Round-trip tests for Excel-side row insertions in Config tables.

Mirrors the user's workflow: generate a workbook, hand-edit Config in Excel
(simulated via openpyxl), run rebuild, assert no data loss.
"""
from pathlib import Path

import openpyxl

from budget_planner import config as cfg
from budget_planner.reader import (
    read_assignments_from_config,
    read_team_from_config,
)
from budget_planner.workbook import init_workbook, rebuild_workbook


def _find_roster_first_data_row(ws):
    for r in range(1, ws.max_row + 1):
        if str(ws.cell(r, 1).value or "").strip().upper() == "TEAM ROSTER":
            return r + 2
    raise AssertionError("TEAM ROSTER section not found")


def test_blank_row_mid_roster_does_not_truncate_rebuild(tmp_path: Path):
    """User inserts an empty row mid-roster (common in Excel); rebuild must keep persons after the blank."""
    path = tmp_path / "wb.xlsx"
    init_workbook(str(path))

    wb = openpyxl.load_workbook(path)
    config_ws = wb["Config"]
    first_data = _find_roster_first_data_row(config_ws)
    # Capture existing roster names for later assertion.
    existing_names: list[str] = []
    r = first_data
    while r <= config_ws.max_row:
        v = config_ws.cell(r, 1).value
        if v is None or (isinstance(v, str) and not v.strip()):
            break
        existing_names.append(str(v).strip())
        r += 1
    assert len(existing_names) >= 2, "seed roster must have at least 2 people for this test"

    # Insert a truly blank row in the middle of the roster.
    insert_at = first_data + 1
    config_ws.insert_rows(insert_at, amount=1)
    wb.save(path)
    wb.close()

    # Pre-rebuild reader must see ALL original persons despite the blank row.
    wb = openpyxl.load_workbook(path)
    roster = read_team_from_config(wb["Config"])
    wb.close()
    names = [p["name"] for p in roster]
    for n in existing_names:
        assert n in names, (
            f"Pre-rebuild read dropped person {n!r} after blank-row insertion. "
            f"Roster read: {names}"
        )

    # Rebuild must succeed and preserve all persons.
    rebuild_workbook(str(path))

    wb = openpyxl.load_workbook(path)
    roster_after = read_team_from_config(wb["Config"])
    assignments_after = read_assignments_from_config(wb["Config"])
    wb.close()
    names_after = [p["name"] for p in roster_after]
    for n in existing_names:
        assert n in names_after, (
            f"Rebuild dropped person {n!r} across blank-row insertion. "
            f"Roster after rebuild: {names_after}"
        )
    dropped_assignees = {a["person"] for a in assignments_after} - set(names_after)
    assert not dropped_assignees, (
        f"Rebuild has assignments pointing at persons no longer in roster: {dropped_assignees}"
    )


def test_mid_insert_in_team_roster_preserved_through_rebuild(tmp_path: Path):
    """User inserts a new person row in the middle of Team Roster; rebuild must keep them."""
    path = tmp_path / "wb.xlsx"
    init_workbook(str(path))

    wb = openpyxl.load_workbook(path)
    config_ws = wb["Config"]
    first_data = _find_roster_first_data_row(config_ws)
    insert_at = first_data + 1
    config_ws.insert_rows(insert_at, amount=1)
    config_ws.cell(insert_at, 1, "Midway Mary")
    config_ws.cell(insert_at, 2, 0.8)
    config_ws.cell(insert_at, 3, next(iter(cfg.GRADE_RATES.keys())))
    wb.save(path)
    wb.close()

    wb = openpyxl.load_workbook(path)
    roster = read_team_from_config(wb["Config"])
    wb.close()
    names = [p["name"] for p in roster]
    assert "Midway Mary" in names, (
        f"Pre-rebuild read already loses inserted person — reader bug. "
        f"Roster read: {names}"
    )

    rebuild_workbook(str(path))

    wb = openpyxl.load_workbook(path)
    roster_after = read_team_from_config(wb["Config"])
    assignments_after = read_assignments_from_config(wb["Config"])
    wb.close()
    names_after = [p["name"] for p in roster_after]
    assert "Midway Mary" in names_after, (
        f"Rebuild dropped the mid-inserted person. Roster after rebuild: {names_after}"
    )
    dropped_assignees = {a["person"] for a in assignments_after} - set(names_after)
    assert not dropped_assignees, (
        f"Rebuild has assignments pointing at persons no longer in roster: {dropped_assignees}"
    )
