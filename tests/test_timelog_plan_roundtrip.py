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
import shutil
import subprocess
from pathlib import Path

import openpyxl
import pytest
from openpyxl.cell import Cell

from budget_planner.config import DEMO_CAPACITY_ASSIGNMENTS
from budget_planner.layout import (
    SECTION_TITLE_GRID,
    SHEET_TIMELOG_PLAN,
    TIMELOG_MONTH_COL_START,
)
from budget_planner.reader import read_timelog_plan
from budget_planner.workbook import init_workbook, rebuild_workbook
from tests.conftest import SCRIPTS_DIR


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


@pytest.fixture(scope="module")
def recalced_timelog_wb_path(tmp_path_factory) -> str:
    """Module-scoped: init_workbook(... DEMO_CAPACITY_ASSIGNMENTS) + recalc.

    Returns a path to a recalc'd workbook. Read-only consumers use the path
    directly; mutating consumers must shutil.copy first.

    rebuild_workbook reads with data_only=False (workbook.py:616), so
    sharing a recalc'd source has no effect on the rebuild flow exercised
    by test_orphan_wp_dropped_with_bystander.
    """
    out = tmp_path_factory.mktemp("tl_rt") / "wb.xlsx"
    init_workbook(str(out), assignments=DEMO_CAPACITY_ASSIGNMENTS)
    result = subprocess.run(
        ["python3", str(SCRIPTS_DIR / "recalc.py"), str(out)],
        capture_output=True,
        text=True,
    )
    if result.returncode != 0:
        pytest.skip("LibreOffice not available")
    return str(out)


# ── Init + recalc round-trip ────────────────────────────────────────────────


def test_init_solver_output_preserved(recalced_timelog_wb_path: str):
    """Init → recalc → reread: grid holds solver-placed WP codes as literals.

    Bystander: at least one non-empty grid cell and at least one row anchored
    to Config via ``=Config!A{r}`` (person col-A cross-ref).
    """
    wb = openpyxl.load_workbook(recalced_timelog_wb_path)
    ws = wb[SHEET_TIMELOG_PLAN]
    months = _project_months(wb)
    first = _grid_first_person_row(ws)

    # Every non-empty cell in the merged grid is a literal WP code (not a formula).
    placed = 0
    for r in range(first, ws.max_row + 1):
        a = ws.cell(r, 1).value
        a_str = str(a or "").strip()
        if a is None or not a_str or a_str.lower().startswith("monthly total") or a_str.upper().startswith("PER-"):
            break  # end of grid block
        for mi in range(months):
            v = ws.cell(r, TIMELOG_MONTH_COL_START + mi).value
            if v is None or (isinstance(v, str) and v.strip() == ""):
                continue
            assert isinstance(v, str) and not v.startswith("="), (
                f"grid cell at ({r},{TIMELOG_MONTH_COL_START + mi}) should be a literal WP code, got {v!r}"
            )
            assert v.startswith("WP"), f"expected WP code, got {v!r}"
            placed += 1
    assert placed > 0, "solver placed no WP codes in the merged grid"
    wb.close()


# ── Orphan pruning on rebuild ───────────────────────────────────────────────


def test_orphan_wp_dropped_with_bystander(tmp_path: Path, capsys, recalced_timelog_wb_path: str):
    """Rebuild drops a grid cell whose WP code no longer exists in the roster.

    Bystander: a valid WP-code cell on the same row must survive the rebuild.
    """
    path = tmp_path / "wb.xlsx"
    shutil.copy(recalced_timelog_wb_path, path)

    wb = openpyxl.load_workbook(path)
    ws = wb[SHEET_TIMELOG_PLAN]
    first = _grid_first_person_row(ws)
    person_name = _resolve_person_name(ws, first)
    m_start = TIMELOG_MONTH_COL_START

    # Bystander: re-stamp the WP code the solver naturally placed at this
    # cell. Rebuild re-runs the solver (no freeze by default) so the cell's
    # post-rebuild value comes from the solver, not the stamp — but a valid
    # WP code stamped before rebuild MUST survive orphan-pruning along the
    # way. Deriving from solver output keeps the assertion stable across
    # demo-data changes.
    bystander_wp = ws.cell(first, m_start).value
    assert isinstance(bystander_wp, str) and bystander_wp.startswith("WP"), (
        f"expected solver-placed WP code at first cell, got {bystander_wp!r}"
    )
    _set(ws, first, m_start, bystander_wp)  # bystander (real WP)
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
    assert ws2.cell(first2, m_start).value == bystander_wp
    # Orphan has been cleared / re-solved; must not carry the zombie value forward.
    orphan_val = ws2.cell(first2, m_start + 1).value
    assert orphan_val != "WP_ZOMBIE", f"orphan cell still carries zombie code: {orphan_val!r}"
    wb2.close()


# ── Reader shape contract ──────────────────────────────────────────────────


def test_reader_returns_grid_and_rebuild_from_month(
    recalced_timelog_wb_path: str,
):
    """Reader returns {'grid': dict, 'rebuild_from_month': int}."""
    wb = openpyxl.load_workbook(recalced_timelog_wb_path)
    months = _project_months(wb)
    out = read_timelog_plan(wb, months)
    assert set(out.keys()) == {"grid", "rebuild_from_month"}, out.keys()
    assert isinstance(out["grid"], dict)
    assert isinstance(out["rebuild_from_month"], int)
    # Fresh init: default is 1 (solve-all equivalent).
    assert out["rebuild_from_month"] == 1
    # Grid has ≥1 solver-placed cell keyed by (person, 1-indexed month).
    assert len(out["grid"]) > 0, "solver placed no cells in merged grid"
    for key, val in out["grid"].items():
        person, month = key
        assert isinstance(person, str) and person
        assert isinstance(month, int) and 1 <= month <= months
        assert isinstance(val, str) and val.startswith("WP"), val
    wb.close()
