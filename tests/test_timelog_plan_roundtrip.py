"""Round-trip tests for Timelog Plan Assigned-grid user overrides.

Covers the contract in ``docs/plans/archived/2026-04-18-timelog-plan-design.md`` §6:
a user types a WP code into an Assigned-grid cell, saves, and rebuilds — the
override survives, while cells the user did NOT touch stay as
``=<Proposed cell>`` formulas (so a re-run of the solver on a changed roster
can still move defaults around).
"""
from __future__ import annotations

import re
from pathlib import Path

import openpyxl
from openpyxl.cell import Cell

from budget_planner.config import DEMO_CAPACITY_ASSIGNMENTS
from budget_planner.layout import (
    SHEET_TIMELOG_PLAN,
    timelog_assigned_col_start,
)
from budget_planner.reader import read_timelog_plan
from budget_planner.workbook import init_workbook, rebuild_workbook


def _project_months(wb) -> int:
    cfg = wb["Config"]
    for r in range(1, cfg.max_row + 1):
        if str(cfg.cell(r, 1).value or "").strip() == "Project Months":
            return int(cfg.cell(r, 2).value)
    raise AssertionError("Project Months row missing from Config")


def _person_first_row(ws) -> int:
    for r in range(1, ws.max_row + 1):
        if ws.cell(r, 1).value == "Person":
            return r + 4
    raise AssertionError("'Person' header not found in col A of Timelog Plan")


def _resolve_person_name(ws, row: int) -> str:
    v = ws.cell(row, 1).value
    if not isinstance(v, str) or not v.startswith("="):
        return str(v)
    m = re.match(r"=Config!A(\d+)", v)
    assert m, f"unexpected col-A formula at row {row}: {v!r}"
    config_ws = ws.parent["Config"]
    return str(config_ws.cell(int(m.group(1)), 1).value).strip()


def _set(ws, row: int, col: int, value) -> None:
    cell = ws.cell(row, col)
    assert isinstance(cell, Cell)
    cell.value = value


# ── Reader-only sanity ──────────────────────────────────────────────────────


def test_reader_empty_on_fresh_workbook(tmp_path: Path):
    """No user typed anything → every Assigned cell is a formula default → {}."""
    path = tmp_path / "wb.xlsx"
    init_workbook(str(path), assignments=DEMO_CAPACITY_ASSIGNMENTS)
    wb = openpyxl.load_workbook(path)
    months = _project_months(wb)
    assert read_timelog_plan(wb, months) == {}
    wb.close()


def test_reader_ignores_formula_cells(tmp_path: Path):
    """=<Proposed> formulas must never be returned as overrides."""
    path = tmp_path / "wb.xlsx"
    init_workbook(str(path), assignments=DEMO_CAPACITY_ASSIGNMENTS)
    wb = openpyxl.load_workbook(path)
    months = _project_months(wb)
    ws = wb[SHEET_TIMELOG_PLAN]
    # Write a literal override on exactly one cell; every other cell stays
    # as its default =<Proposed> formula (bystander check — selectivity).
    first = _person_first_row(ws)
    person_name = _resolve_person_name(ws, first)
    a_start = timelog_assigned_col_start(months)
    _set(ws, first, a_start, "WP2")
    wb.save(path)
    wb.close()

    wb2 = openpyxl.load_workbook(path)
    overrides = read_timelog_plan(wb2, months)
    wb2.close()
    assert overrides == {(person_name, 1): "WP2"}


# ── Full round-trip: init → edit → rebuild → preserved ─────────────────────


def test_override_survives_rebuild(tmp_path: Path):
    """User-typed WP code in Assigned grid survives a rebuild.

    Bystander check: a sibling cell left untouched remains a formula (i.e.
    the builder didn't clobber every cell with a literal on round-trip).
    """
    path = tmp_path / "wb.xlsx"
    init_workbook(str(path), assignments=DEMO_CAPACITY_ASSIGNMENTS)

    wb = openpyxl.load_workbook(path)
    months = _project_months(wb)
    ws = wb[SHEET_TIMELOG_PLAN]
    first = _person_first_row(ws)
    person_name = _resolve_person_name(ws, first)
    a_start = timelog_assigned_col_start(months)

    # Edit column a_start (month 1) on the first person row.
    edit_col = a_start
    edit_value = "WP2"
    _set(ws, first, edit_col, edit_value)
    wb.save(path)
    wb.close()

    # Rebuild and reopen.
    rebuild_workbook(input_path=str(path))
    wb2 = openpyxl.load_workbook(path)
    ws2 = wb2[SHEET_TIMELOG_PLAN]
    first2 = _person_first_row(ws2)
    # Person row order is (priority asc, name asc) — stable across rebuild
    # when the roster doesn't change, so the first person row holds the
    # same person. Resolve to assert explicitly.
    assert _resolve_person_name(ws2, first2) == person_name

    got = ws2.cell(first2, edit_col).value
    assert got == edit_value, (
        f"override lost on rebuild: expected {edit_value!r}, got {got!r}"
    )

    # Bystander — neighbouring cell should still be a formula default.
    neighbour = ws2.cell(first2, edit_col + 1).value
    assert isinstance(neighbour, str) and neighbour.startswith("="), (
        f"neighbour cell should be =<Proposed> default, got {neighbour!r}"
    )
    wb2.close()


def test_proposed_grid_regenerates_on_rebuild(tmp_path: Path):
    """Proposed grid is solver-written — a user typing into Assigned must not
    bleed through into Proposed. After a rebuild, Proposed cells still hold
    literal WP codes (solver output), not formula artifacts of the edit."""
    path = tmp_path / "wb.xlsx"
    init_workbook(str(path), assignments=DEMO_CAPACITY_ASSIGNMENTS)

    wb = openpyxl.load_workbook(path)
    months = _project_months(wb)
    ws = wb[SHEET_TIMELOG_PLAN]
    first = _person_first_row(ws)
    a_start = timelog_assigned_col_start(months)
    _set(ws, first, a_start, "WP3")
    wb.save(path)
    wb.close()

    rebuild_workbook(input_path=str(path))

    wb2 = openpyxl.load_workbook(path)
    ws2 = wb2[SHEET_TIMELOG_PLAN]
    first2 = _person_first_row(ws2)
    # Proposed grid starts at MONTH_COL_START (col 3). Pull a cell that
    # the solver placed — WP code literal, not a formula.
    from budget_planner.layout import MONTH_COL_START
    placed_found = False
    for mi in range(months):
        val = ws2.cell(first2, MONTH_COL_START + mi).value
        if isinstance(val, str) and val.startswith("WP"):
            placed_found = True
            break
    assert placed_found, "Proposed row had no solver-placed WP cell after rebuild"
    wb2.close()


# ── Orphan pruning ──────────────────────────────────────────────────────────


def test_override_orphaned_by_removed_wp_is_dropped(tmp_path: Path, capsys):
    """Overriding with a WP code the user later deletes from Config must drop
    the stale cell with a warning — parallel to the Actuals/Hours orphan
    pruning already in rebuild.

    Bystander: a sibling override with a valid WP code must survive.
    """
    path = tmp_path / "wb.xlsx"
    init_workbook(str(path), assignments=DEMO_CAPACITY_ASSIGNMENTS)

    wb = openpyxl.load_workbook(path)
    months = _project_months(wb)
    ws = wb[SHEET_TIMELOG_PLAN]
    first = _person_first_row(ws)
    person_name = _resolve_person_name(ws, first)
    a_start = timelog_assigned_col_start(months)

    # Write two overrides: one valid, one with a WP code that does not
    # exist in the WP Budget table.
    _set(ws, first, a_start, "WP2")           # valid — bystander
    _set(ws, first, a_start + 1, "WP_ZOMBIE")  # orphan
    wb.save(path)
    wb.close()

    rebuild_workbook(input_path=str(path))
    captured = capsys.readouterr()
    assert "Orphaned Timelog Plan override" in captured.out, captured.out
    assert "WP_ZOMBIE" in captured.out

    # Reopen and verify: bystander survived, orphan reverted to default formula.
    wb2 = openpyxl.load_workbook(path)
    ws2 = wb2[SHEET_TIMELOG_PLAN]
    first2 = _person_first_row(ws2)
    assert _resolve_person_name(ws2, first2) == person_name
    assert ws2.cell(first2, a_start).value == "WP2"  # bystander
    zombie = ws2.cell(first2, a_start + 1).value
    assert isinstance(zombie, str) and zombie.startswith("="), (
        f"orphan cell should revert to default formula, got {zombie!r}"
    )
    wb2.close()
