"""Tests for the Timelog Plan resolve_flag gate on rebuild.

Semantics (Phase 2):
- resolve_flag=TRUE  → solver runs, grid is re-derived, user edits discarded.
- resolve_flag=FALSE → solver skipped, grid written verbatim from the
  preserved-grid dict (reader output).
- resolve_flag missing (pre-v2 workbook) → reader defaults to TRUE.

Each test uses a distinct "edit" cell that the solver would NOT place itself,
so the assertion can distinguish "solver wiped" vs. "edit preserved".
"""
from __future__ import annotations

from pathlib import Path

import openpyxl
from openpyxl.cell import Cell
from openpyxl.workbook.defined_name import DefinedName

from budget_planner.config import DEMO_CAPACITY_ASSIGNMENTS
from budget_planner.layout import SHEET_TIMELOG_PLAN, TIMELOG_MONTH_COL_START
from budget_planner.reader import read_timelog_plan
from budget_planner.timelog_plan import SECTION_TITLE_GRID
from budget_planner.workbook import init_workbook, rebuild_workbook


def _grid_first_person_row(ws) -> int:
    for r in range(1, ws.max_row + 1):
        if ws.cell(r, 1).value == SECTION_TITLE_GRID:
            return r + 4
    raise AssertionError(f"{SECTION_TITLE_GRID!r} not found")


def _project_months(wb) -> int:
    cfg = wb["Config"]
    for r in range(1, cfg.max_row + 1):
        if str(cfg.cell(r, 1).value or "").strip() == "Project Months":
            return int(cfg.cell(r, 2).value)
    raise AssertionError("Project Months row missing from Config")


def _set(ws, row: int, col: int, value) -> None:
    cell = ws.cell(row, col)
    assert isinstance(cell, Cell)
    cell.value = value


def _pick_editable_cell(ws, months: int) -> tuple[int, int, str | None]:
    """Return (row, col, solver_value) of the first non-empty grid cell."""
    first = _grid_first_person_row(ws)
    for mi in range(months):
        col = TIMELOG_MONTH_COL_START + mi
        v = ws.cell(first, col).value
        if isinstance(v, str) and v.strip() and not v.startswith("="):
            return first, col, v
    raise AssertionError("no solver-placed grid cell found on first person row")


def test_flag_true_wipes_edits(tmp_path: Path):
    """resolve_flag=TRUE: rebuild re-solves the grid, user edit is discarded."""
    path = tmp_path / "wb.xlsx"
    init_workbook(str(path), assignments=DEMO_CAPACITY_ASSIGNMENTS)

    wb = openpyxl.load_workbook(path)
    ws = wb[SHEET_TIMELOG_PLAN]
    months = _project_months(wb)
    row, col, _ = _pick_editable_cell(ws, months)
    # Stamp a sentinel that the solver will NEVER place (WP0 is not a real code).
    _set(ws, row, col, "WP0_EDIT")
    wb.save(path)
    wb.close()

    # Flag is TRUE by default on init — confirm, then rebuild.
    wb2 = openpyxl.load_workbook(path)
    pre = read_timelog_plan(wb2, months)
    assert pre["resolve_flag"] is True
    assert pre["grid"].get((_first_person_name(wb2), col - TIMELOG_MONTH_COL_START + 1)) == "WP0_EDIT"
    wb2.close()

    rebuild_workbook(input_path=str(path))

    wb3 = openpyxl.load_workbook(path)
    ws3 = wb3[SHEET_TIMELOG_PLAN]
    # Solver re-ran → edit must not survive.
    assert ws3.cell(row, col).value != "WP0_EDIT"
    # Bystander: rebuild must have produced a real grid — at least one solver-
    # placed WP code on the same row. Without this the test passes even if the
    # rebuilt grid is entirely empty (edit removed because nothing was written).
    row_cells = [
        ws3.cell(row, TIMELOG_MONTH_COL_START + mi).value
        for mi in range(months)
    ]
    solver_placed = [
        v for v in row_cells
        if isinstance(v, str) and v.strip() and v != "WP0_EDIT"
    ]
    assert len(solver_placed) >= 1, (
        f"flag=TRUE should re-run solver; row {row} has no solver output: {row_cells!r}"
    )
    wb3.close()


def test_flag_false_preserves_edits(tmp_path: Path):
    """resolve_flag=FALSE: rebuild preserves user edits in the grid verbatim."""
    path = tmp_path / "wb.xlsx"
    init_workbook(str(path), assignments=DEMO_CAPACITY_ASSIGNMENTS)

    wb = openpyxl.load_workbook(path)
    ws = wb[SHEET_TIMELOG_PLAN]
    months = _project_months(wb)
    row, col, solver_value = _pick_editable_cell(ws, months)
    # Flip the edit to a DIFFERENT valid WP code — still within roster so it
    # survives orphan pruning. Pick any WP code other than the solver's choice.
    edit_value = "WP2" if solver_value != "WP2" else "WP1"
    _set(ws, row, col, edit_value)

    # Flip the resolve flag cell to FALSE.
    _set_resolve_flag(wb, False)
    wb.save(path)
    wb.close()

    rebuild_workbook(input_path=str(path))

    wb2 = openpyxl.load_workbook(path)
    ws2 = wb2[SHEET_TIMELOG_PLAN]
    # Edit must survive — flag was FALSE, solver did not run.
    assert ws2.cell(row, col).value == edit_value, (
        f"expected preserved edit {edit_value!r}, got {ws2.cell(row, col).value!r}"
    )
    # Bystander: at least one OTHER solver-placed cell on the same row survived too.
    preserved_row_cells = [
        ws2.cell(row, TIMELOG_MONTH_COL_START + mi).value
        for mi in range(months)
    ]
    non_empty = [v for v in preserved_row_cells if isinstance(v, str) and v.strip()]
    assert len(non_empty) >= 2, (
        f"flag=FALSE should preserve the whole grid row verbatim, got {non_empty!r}"
    )
    wb2.close()


def test_flag_missing_defaults_true(tmp_path: Path):
    """Pre-v2 workbook (no named range) → reader defaults to TRUE, solver runs."""
    path = tmp_path / "wb.xlsx"
    init_workbook(str(path), assignments=DEMO_CAPACITY_ASSIGNMENTS)

    # Simulate a pre-v2 workbook: delete the named range entirely.
    wb = openpyxl.load_workbook(path)
    if "timelog_resolve_flag" in wb.defined_names:
        del wb.defined_names["timelog_resolve_flag"]
    wb.save(path)
    wb.close()

    wb2 = openpyxl.load_workbook(path)
    months = _project_months(wb2)
    out = read_timelog_plan(wb2, months)
    assert out["resolve_flag"] is True, (
        "missing named range must default to TRUE (re-solve) for pre-v2 compat"
    )
    wb2.close()

    # Rebuild must succeed — the builder re-creates the named range + toggle cell.
    rebuild_workbook(input_path=str(path))
    wb3 = openpyxl.load_workbook(path)
    assert "timelog_resolve_flag" in wb3.defined_names
    wb3.close()


# ── Helpers ────────────────────────────────────────────────────────────────


def _first_person_name(wb) -> str:
    ws = wb[SHEET_TIMELOG_PLAN]
    row = _grid_first_person_row(ws)
    v = ws.cell(row, 1).value
    if isinstance(v, str) and v.startswith("=Config!A"):
        import re
        m = re.match(r"=Config!A(\d+)", v)
        assert m
        cfg = wb["Config"]
        return str(cfg.cell(int(m.group(1)), 1).value).strip()
    return str(v).strip()


def _set_resolve_flag(wb, value: bool) -> None:
    """Flip the timelog_resolve_flag toggle cell to True/False."""
    dn = wb.defined_names["timelog_resolve_flag"]
    # attr_text looks like "'Timelog Plan'!$J$2" — parse sheet + coord.
    assert isinstance(dn, DefinedName)
    dests = list(dn.destinations)
    assert len(dests) == 1, dests
    sheet_name, coord = dests[0]
    ws = wb[sheet_name]
    cell = ws[coord]
    assert isinstance(cell, Cell)
    cell.value = value
