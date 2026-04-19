"""Round-trip tests for Timelog Plan v2 merged-grid shape.

v2 merges the v1 PROPOSED / ASSIGNED bands into a single editable grid of
literal WP codes (solver-written on init). The `=<Proposed>` round-trip
mechanism is gone. The tests here cover:

* solver output survives openpyxl round-trip after init + recalc
* orphan pruning on rebuild drops stale (person, wp) cells with a bystander
  surviving — the pruning loop lives in ``workbook.py`` around the
  ``read_timelog_plan`` call.
"""
from __future__ import annotations

import re
import subprocess
from pathlib import Path

import openpyxl
from openpyxl.cell import Cell

from budget_planner.config import DEMO_CAPACITY_ASSIGNMENTS
from budget_planner.layout import (
    SHEET_TIMELOG_PLAN,
    TIMELOG_MONTH_COL_START,
)
from budget_planner.reader import read_timelog_plan
from budget_planner.timelog_plan import SECTION_TITLE_GRID
from budget_planner.workbook import init_workbook, rebuild_workbook


def _project_months(wb) -> int:
    cfg = wb["Config"]
    for r in range(1, cfg.max_row + 1):
        if str(cfg.cell(r, 1).value or "").strip() == "Project Months":
            return int(cfg.cell(r, 2).value)
    raise AssertionError("Project Months row missing from Config")


def _grid_first_person_row(ws) -> int:
    """First person row = grid title row + 4 (title, year, month, month-idx)."""
    for r in range(1, ws.max_row + 1):
        if ws.cell(r, 1).value == SECTION_TITLE_GRID:
            return r + 4
    raise AssertionError(f"{SECTION_TITLE_GRID!r} not found in col A")


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


def _recalc(path: Path) -> None:
    subprocess.run(
        ["python", "scripts/recalc.py", str(path)],
        check=True, capture_output=True,
    )


# ── Init + recalc round-trip ────────────────────────────────────────────────


def test_init_solver_output_preserved(tmp_path: Path):
    """Init → recalc → reread: grid holds solver-placed WP codes as literals.

    Bystander: at least one non-empty grid cell and at least one row anchored
    to Config via ``=Config!A{r}`` (person col-A cross-ref).
    """
    path = tmp_path / "wb.xlsx"
    init_workbook(str(path), assignments=DEMO_CAPACITY_ASSIGNMENTS)
    _recalc(path)

    wb = openpyxl.load_workbook(path)
    ws = wb[SHEET_TIMELOG_PLAN]
    months = _project_months(wb)
    first = _grid_first_person_row(ws)

    # Every non-empty cell in the merged grid is a literal WP code (not a formula).
    placed = 0
    for r in range(first, ws.max_row + 1):
        a = ws.cell(r, 1).value
        if a is None or (isinstance(a, str) and a.strip() == ""):
            break  # end of grid block
        for mi in range(months):
            v = ws.cell(r, TIMELOG_MONTH_COL_START + mi).value
            if v is None or (isinstance(v, str) and v.strip() == ""):
                continue
            assert isinstance(v, str) and not v.startswith("="), (
                f"grid cell at ({r},{TIMELOG_MONTH_COL_START + mi}) "
                f"should be a literal WP code, got {v!r}"
            )
            assert v.startswith("WP"), f"expected WP code, got {v!r}"
            placed += 1
    assert placed > 0, "solver placed no WP codes in the merged grid"
    wb.close()


# ── Orphan pruning on rebuild ───────────────────────────────────────────────


def test_orphan_wp_dropped_with_bystander(tmp_path: Path, capsys):
    """Rebuild drops a grid cell whose WP code no longer exists in the roster.

    Bystander: a valid WP-code cell on the same row must survive the rebuild.
    """
    path = tmp_path / "wb.xlsx"
    init_workbook(str(path), assignments=DEMO_CAPACITY_ASSIGNMENTS)

    wb = openpyxl.load_workbook(path)
    ws = wb[SHEET_TIMELOG_PLAN]
    first = _grid_first_person_row(ws)
    person_name = _resolve_person_name(ws, first)
    m_start = TIMELOG_MONTH_COL_START

    # Stamp a known-valid WP code (bystander) and a zombie code on the same row.
    _set(ws, first, m_start, "WP2")           # bystander (real WP)
    _set(ws, first, m_start + 1, "WP_ZOMBIE")  # orphan
    wb.save(path)
    wb.close()

    rebuild_workbook(input_path=str(path))
    captured = capsys.readouterr()
    assert "Orphaned Timelog Plan cell" in captured.out, captured.out
    assert "WP_ZOMBIE" in captured.out

    wb2 = openpyxl.load_workbook(path)
    ws2 = wb2[SHEET_TIMELOG_PLAN]
    first2 = _grid_first_person_row(ws2)
    assert _resolve_person_name(ws2, first2) == person_name
    # Bystander survives.
    assert ws2.cell(first2, m_start).value == "WP2"
    # Orphan has been cleared / re-solved; must not carry the zombie value forward.
    orphan_val = ws2.cell(first2, m_start + 1).value
    assert orphan_val != "WP_ZOMBIE", (
        f"orphan cell still carries zombie code: {orphan_val!r}"
    )
    wb2.close()


# ── Reader shape contract ──────────────────────────────────────────────────


def test_reader_returns_grid_and_flag(tmp_path: Path):
    """v2 reader returns {'grid': {(person, month): wp_code}, 'resolve_flag': bool}."""
    path = tmp_path / "wb.xlsx"
    init_workbook(str(path), assignments=DEMO_CAPACITY_ASSIGNMENTS)
    wb = openpyxl.load_workbook(path)
    months = _project_months(wb)
    out = read_timelog_plan(wb, months)
    assert set(out.keys()) == {"grid", "resolve_flag"}, out.keys()
    assert isinstance(out["grid"], dict)
    assert isinstance(out["resolve_flag"], bool)
    # Fresh init: flag defaults to TRUE (named range writes literal True).
    assert out["resolve_flag"] is True
    # Grid has ≥1 solver-placed cell keyed by (person, 1-indexed month).
    assert len(out["grid"]) > 0, "solver placed no cells in merged grid"
    for key, val in out["grid"].items():
        person, month = key
        assert isinstance(person, str) and person
        assert isinstance(month, int) and 1 <= month <= months
        assert isinstance(val, str) and val.startswith("WP"), val
    wb.close()
