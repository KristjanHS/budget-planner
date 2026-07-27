"""Tests for the TEAM ROSTER Priority field round-trip and clamping.

Priority is the solver tie-breaker (1 = highest, 3 = lowest); 0/4/7 etc. would
silently exclude a person from every WP pass since the solver equality-matches
prio ∈ (1, 2, 3). These tests guard the read+rebuild contract for Priority.
"""

import openpyxl

from budget_planner.layout import COL_ROSTER_PRIORITY, SECTION_HEADER_LABELS
from budget_planner.reader import _find_header_row, read_team_from_config
from budget_planner.workbook import init_workbook, rebuild_workbook


def _apply_priority_edits(cfg, roster_pre, edits):
    """Write Priority cells for named persons in the TEAM ROSTER section."""
    roster_section = _find_header_row(cfg, SECTION_HEADER_LABELS["team_roster"], required=True)
    header_row = roster_section + 1
    for r in range(header_row + 1, header_row + 1 + len(roster_pre)):
        name = cfg.cell(r, 1).value
        if name in edits:
            cfg.cell(row=r, column=COL_ROSTER_PRIORITY, value=edits[name])


def test_read_team_priority_roundtrip(tmp_path):
    """User edits Priority → rebuild → priority preserved."""
    path = tmp_path / "wb.xlsx"
    init_workbook(str(path))

    wb = openpyxl.load_workbook(path)
    cfg = wb["Config"]
    roster_pre = read_team_from_config(cfg)
    # Pick the first two team members and set distinct priorities.
    target_names = [roster_pre[0]["name"], roster_pre[1]["name"]]
    edits = {target_names[0]: 1, target_names[1]: 3}

    _apply_priority_edits(cfg, roster_pre, edits)
    wb.save(path)
    wb.close()

    rebuild_workbook(input_path=str(path))

    wb2 = openpyxl.load_workbook(path)
    roster_post = read_team_from_config(wb2["Config"])
    wb2.close()

    got = {m["name"]: m["priority"] for m in roster_post}
    for name, expected in edits.items():
        assert got[name] == expected, f"{name}: expected priority {expected}, got {got[name]}"


def test_read_team_priority_clamps_out_of_range(tmp_path):
    """Out-of-range Priority values are clamped to [1, 3] — a user typing 0 or
    4 into the Priority cell would otherwise be silently excluded from every
    WP pass by the solver, since it equality-matches prio ∈ (1,2,3).
    """
    path = tmp_path / "wb.xlsx"
    init_workbook(str(path))

    wb = openpyxl.load_workbook(path)
    cfg = wb["Config"]
    roster_pre = read_team_from_config(cfg)
    assert len(roster_pre) >= 3
    targets = [roster_pre[0]["name"], roster_pre[1]["name"], roster_pre[2]["name"]]
    # 0 → clamp to 1; 4 → clamp to 3; 7 → clamp to 3.
    edits = {targets[0]: 0, targets[1]: 4, targets[2]: 7}
    expected = {targets[0]: 1, targets[1]: 3, targets[2]: 3}

    _apply_priority_edits(cfg, roster_pre, edits)
    wb.save(path)
    wb.close()

    wb2 = openpyxl.load_workbook(path)
    roster_post = read_team_from_config(wb2["Config"])
    wb2.close()

    got = {m["name"]: m["priority"] for m in roster_post}
    for name, expect in expected.items():
        assert got[name] == expect, f"{name}: expected clamp to {expect}, got {got[name]}"
