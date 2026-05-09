"""Round-trip tests for the Actuals-Hours timelog on rebuild.

Covers:
- Blue-input hours survive a rebuild (init -> Excel edit -> rebuild -> reopen).
- Internal blank WP rows inside a person section don't truncate the reader.
"""
from __future__ import annotations

from pathlib import Path

import openpyxl
from openpyxl.cell import Cell

from budget_planner.config import DEMO_CAPACITY_ASSIGNMENTS
from budget_planner.layout import MONTH_COL_START
from budget_planner.reader import read_actuals_hours_sheet
from budget_planner.workbook import init_workbook, rebuild_workbook


SHEET = "Actuals-Hours"


def _set(ws, row: int, col: int, value: float | int) -> None:
    cell = ws.cell(row, col)
    assert isinstance(cell, Cell)
    cell.value = value


def _find_person_header_row(ws, person_name: str) -> int:
    """Person headers carry the literal name in col A and are merged A:B."""
    merged = [(mr.min_row, mr.min_col, mr.max_col)
              for mr in ws.merged_cells.ranges]
    for r in range(1, ws.max_row + 1):
        if ws.cell(r, 1).value != person_name:
            continue
        if any(mr_r == r and mr_c1 == 1 and mr_c2 >= 2
               for mr_r, mr_c1, mr_c2 in merged):
            return r
    raise AssertionError(f"Person header {person_name!r} not found")


def _find_wp_row_under_person(ws, header_row: int, wp_id: str) -> int:
    """Walk rows under a person header; match col A that resolves to ``wp_id``.

    Col A on WP rows is ``=Config!A{n}``; resolve via the Config sheet.
    """
    config_ws = ws.parent["Config"]
    import re
    for r in range(header_row + 1, ws.max_row + 1):
        val = ws.cell(r, 1).value
        if val is None:
            continue
        b_val = ws.cell(r, 2).value
        if b_val and str(b_val).upper() == "TOTAL":
            return -1
        if not isinstance(val, str):
            continue
        m = re.match(r"=Config!A(\d+)", val)
        if m:
            resolved = config_ws.cell(int(m.group(1)), 1).value
            if resolved and str(resolved).strip() == wp_id:
                return r
    raise AssertionError(f"WP row for {wp_id!r} under row {header_row} not found")


def test_reader_round_trip_preserves_hours(tmp_path: Path):
    """Init -> write hours at specific cells -> rebuild -> hours preserved."""
    path = tmp_path / "wb.xlsx"
    init_workbook(str(path), assignments=DEMO_CAPACITY_ASSIGNMENTS)

    # Pick 3 different (person, wp, month_index, hours) tuples.
    # Drive test data off DEMO_CAPACITY_ASSIGNMENTS so hours land on
    # rows where the cell is a blue input (WP active for that person).
    by_person: dict[str, list[str]] = {}
    for a in DEMO_CAPACITY_ASSIGNMENTS:
        by_person.setdefault(a["person"], []).append(a["wp"])
    people = [p for p, wps in by_person.items() if wps]
    assert len(people) >= 2, "demo data needs at least 2 assigned people"
    p1, p2 = people[0], people[1]
    wp_p1 = by_person[p1][0]
    wp_p2 = by_person[p2][0]
    # Use assignment start month so the cell is definitely an active (blue) cell.
    asgn_p1 = next(a for a in DEMO_CAPACITY_ASSIGNMENTS
                   if a["person"] == p1 and a["wp"] == wp_p1)
    asgn_p2 = next(a for a in DEMO_CAPACITY_ASSIGNMENTS
                   if a["person"] == p2 and a["wp"] == wp_p2)
    edits = [
        (p1, wp_p1, asgn_p1["start"] - 1, 120.0),
        (p1, wp_p1, asgn_p1["start"],     80.0),
        (p2, wp_p2, asgn_p2["start"] - 1, 40.0),
    ]

    # Write edits via openpyxl.
    wb = openpyxl.load_workbook(path)
    ws = wb[SHEET]
    for person, wp_id, month_idx, hours in edits:
        hdr = _find_person_header_row(ws, person)
        wp_row = _find_wp_row_under_person(ws, hdr, wp_id)
        assert wp_row > 0
        _set(ws, wp_row, MONTH_COL_START + month_idx, hours)
    wb.save(path)
    wb.close()

    # Rebuild.
    rebuild_workbook(input_path=str(path))

    # Re-read directly via the new reader and assert edits survived.
    wb = openpyxl.load_workbook(path)
    hours_data = read_actuals_hours_sheet(wb[SHEET])
    wb.close()

    for person, wp_id, month_idx, expected in edits:
        assert person in hours_data, f"person {person!r} missing from reader"
        assert wp_id in hours_data[person], (
            f"WP {wp_id!r} missing for {person!r}: "
            f"{list(hours_data[person].keys())}"
        )
        got = hours_data[person][wp_id][month_idx]
        assert got == expected, (
            f"{person}/{wp_id}/month{month_idx}: expected {expected}, got {got}"
        )


def test_reader_tolerates_internal_blank_rows(tmp_path: Path):
    """Hours on later WP rows survive even when earlier WP rows for the same
    person have no data (internal blanks must not truncate the walk)."""
    path = tmp_path / "wb.xlsx"
    init_workbook(str(path), assignments=DEMO_CAPACITY_ASSIGNMENTS)

    # Find a person with 2+ assigned WPs so we can leave the first blank
    # and write hours on the second.
    by_person: dict[str, list[dict]] = {}
    for a in DEMO_CAPACITY_ASSIGNMENTS:
        by_person.setdefault(a["person"], []).append(a)
    target = next((p for p, alist in by_person.items() if len(alist) >= 2), None)
    assert target is not None, "demo data needs a person with 2+ WPs"
    asgns = by_person[target]
    second_asgn = asgns[1]
    wp_id = second_asgn["wp"]
    month_idx = second_asgn["start"] - 1
    hours = 55.5

    wb = openpyxl.load_workbook(path)
    ws = wb[SHEET]
    hdr = _find_person_header_row(ws, target)
    wp_row = _find_wp_row_under_person(ws, hdr, wp_id)
    _set(ws, wp_row, MONTH_COL_START + month_idx, hours)
    wb.save(path)
    wb.close()

    rebuild_workbook(input_path=str(path))

    wb = openpyxl.load_workbook(path)
    hours_data = read_actuals_hours_sheet(wb[SHEET])
    wb.close()

    assert target in hours_data
    assert wp_id in hours_data[target]
    assert hours_data[target][wp_id][month_idx] == hours

    # The first WP row was never populated → all months should be None.
    first_wp = asgns[0]["wp"]
    assert first_wp in hours_data[target]
    assert all(v is None for v in hours_data[target][first_wp])
