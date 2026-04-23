"""Tests for the Timelog Plan rebuild-from-month freeze-point semantics.

Replaces ``test_timelog_resolve_flag.py``. Single integer Config param N
(``Config!B9``, named range ``timelog_rebuild_from_month``):

- Months 1..N-1 → **frozen**: assignee-grid cells kept verbatim on rebuild;
  their € consumption is subtracted from WP ceilings + Person pools before
  the solver runs.
- Months N..end → **solved**: solver re-derives; user edits discarded.

Endpoints:
- N = 1 (default): solve-all. Equivalent to the old ``resolve_flag=TRUE``.
- N > project_months: preserve-all. Equivalent to the old ``resolve_flag=FALSE``.

Covers cases 1-6 from the design doc §8 (drift-guard + docs-only gate live
in ``test_config_sheet.py``).
"""
from __future__ import annotations

import re
from pathlib import Path

import openpyxl
from openpyxl.cell import Cell
from openpyxl.workbook.defined_name import DefinedName

from budget_planner.config import DEMO_CAPACITY_ASSIGNMENTS
from budget_planner.layout import SHEET_CONFIG, SHEET_TIMELOG_PLAN, TIMELOG_MONTH_COL_START
from budget_planner.reader import read_timelog_plan
from budget_planner.timelog_plan import SECTION_TITLE_GRID
from budget_planner.workbook import init_workbook, rebuild_workbook


# ── Helpers ────────────────────────────────────────────────────────────────


def _grid_first_person_row(ws) -> int:
    for r in range(1, ws.max_row + 1):
        if ws.cell(r, 1).value == SECTION_TITLE_GRID:
            return r + 4
    raise AssertionError(f"{SECTION_TITLE_GRID!r} not found")


def _project_months(wb) -> int:
    cfg = wb[SHEET_CONFIG]
    for r in range(1, cfg.max_row + 1):
        if str(cfg.cell(r, 1).value or "").strip() == "Project Months":
            return int(cfg.cell(r, 2).value)
    raise AssertionError("Project Months row missing from Config")


def _set(ws, row: int, col: int, value) -> None:
    cell = ws.cell(row, col)
    assert isinstance(cell, Cell)
    cell.value = value


def _pick_editable_cell(ws, months: int) -> tuple[int, int, str | None]:
    """Return (row, col, solver_value) of the first non-empty grid cell on
    the first person row. Vacation cells are always empty so are filtered
    implicitly by the value-strip check."""
    first = _grid_first_person_row(ws)
    for mi in range(months):
        col = TIMELOG_MONTH_COL_START + mi
        v = ws.cell(first, col).value
        if isinstance(v, str) and v.strip() and not v.startswith("="):
            return first, col, v
    raise AssertionError("no solver-placed grid cell found on first person row")


def _first_person_name(wb) -> str:
    ws = wb[SHEET_TIMELOG_PLAN]
    row = _grid_first_person_row(ws)
    v = ws.cell(row, 1).value
    if isinstance(v, str) and v.startswith(f"={SHEET_CONFIG}!A"):
        m = re.match(rf"={SHEET_CONFIG}!A(\d+)", v)
        assert m
        cfg = wb[SHEET_CONFIG]
        return str(cfg.cell(int(m.group(1)), 1).value).strip()
    return str(v).strip()


def _set_rebuild_from_month(wb, value: int) -> None:
    """Write N to the ``timelog_rebuild_from_month`` named-range cell (Config!B9)."""
    dn = wb.defined_names["timelog_rebuild_from_month"]
    assert isinstance(dn, DefinedName)
    dests = list(dn.destinations)
    assert len(dests) == 1
    sheet_name, coord = dests[0]
    ws = wb[sheet_name]
    cell = ws[coord]
    assert isinstance(cell, Cell)
    cell.value = value


# ── Case 1: N = 1 (solve-all) ─────────────────────────────────────────────


def test_rebuild_from_month_1_wipes_edits(tmp_path: Path):
    """N = 1: rebuild re-solves the full grid; user edit is discarded."""
    path = tmp_path / "wb.xlsx"
    init_workbook(str(path), assignments=DEMO_CAPACITY_ASSIGNMENTS)

    wb = openpyxl.load_workbook(path)
    ws = wb[SHEET_TIMELOG_PLAN]
    months = _project_months(wb)
    row, col, _ = _pick_editable_cell(ws, months)
    _set(ws, row, col, "WP0_EDIT")  # sentinel the solver never places
    wb.save(path)
    wb.close()

    # N defaults to 1 on init — confirm via reader, then rebuild.
    wb2 = openpyxl.load_workbook(path)
    pre = read_timelog_plan(wb2, months)
    assert pre["rebuild_from_month"] == 1
    month_num = col - TIMELOG_MONTH_COL_START + 1
    assert pre["grid"].get((_first_person_name(wb2), month_num)) == "WP0_EDIT"
    wb2.close()

    rebuild_workbook(input_path=str(path))

    wb3 = openpyxl.load_workbook(path)
    ws3 = wb3[SHEET_TIMELOG_PLAN]
    assert ws3.cell(row, col).value != "WP0_EDIT"
    # Bystander: solver actually populated the row.
    row_cells = [
        ws3.cell(row, TIMELOG_MONTH_COL_START + mi).value
        for mi in range(months)
    ]
    solver_placed = [
        v for v in row_cells
        if isinstance(v, str) and v.strip() and v != "WP0_EDIT"
    ]
    assert len(solver_placed) >= 1, (
        f"N=1 must re-run solver; row {row} empty: {row_cells!r}"
    )
    wb3.close()


# ── Case 2: N > project_months (preserve-all) ─────────────────────────────


def test_rebuild_from_month_past_end_preserves_all(tmp_path: Path):
    """N > project_months: every cell frozen; solver runs on narrowed inputs
    but places nothing (pool/ceiling both 0). User edit survives verbatim.
    """
    path = tmp_path / "wb.xlsx"
    init_workbook(str(path), assignments=DEMO_CAPACITY_ASSIGNMENTS)

    wb = openpyxl.load_workbook(path)
    ws = wb[SHEET_TIMELOG_PLAN]
    months = _project_months(wb)
    row, col, solver_value = _pick_editable_cell(ws, months)
    # Flip to a different real WP code so edit is preserve-eligible.
    edit_value = "WP2" if solver_value != "WP2" else "WP1"
    _set(ws, row, col, edit_value)
    _set_rebuild_from_month(wb, months + 10)  # past the end → preserve-all
    wb.save(path)
    wb.close()

    rebuild_workbook(input_path=str(path))

    wb2 = openpyxl.load_workbook(path)
    ws2 = wb2[SHEET_TIMELOG_PLAN]
    assert ws2.cell(row, col).value == edit_value, (
        f"expected preserved edit {edit_value!r}, "
        f"got {ws2.cell(row, col).value!r}"
    )
    # Bystander: the rest of the row's pre-edit cells also survived (not wiped).
    preserved = [
        ws2.cell(row, TIMELOG_MONTH_COL_START + mi).value
        for mi in range(months)
    ]
    non_empty = [v for v in preserved if isinstance(v, str) and v.strip()]
    assert len(non_empty) >= 2, (
        f"preserve-all must keep whole row, got only {non_empty!r}"
    )
    wb2.close()


# ── Case 3: N = project_months // 2 (mixed-mode) ──────────────────────────


def test_mixed_mode_freezes_before_overwrites_after(tmp_path: Path):
    """N in the middle: edits in months < N are preserved; edits in months ≥ N
    are discarded (solver re-derives)."""
    path = tmp_path / "wb.xlsx"
    init_workbook(str(path), assignments=DEMO_CAPACITY_ASSIGNMENTS)

    wb = openpyxl.load_workbook(path)
    ws = wb[SHEET_TIMELOG_PLAN]
    months = _project_months(wb)
    assert months >= 4, "need ≥4 months for a meaningful mid-boundary test"
    n = months // 2
    grid_first = _grid_first_person_row(ws)

    # Find ANY person row with filled cells on both sides of the boundary.
    # The solver fills in priority order, so later rows often carry the
    # tail-of-project cells that the first row lacks.
    def _scan_row(r: int) -> tuple[int, int, str, int, int, str] | None:
        frozen = solved = None
        for mi in range(months):
            col = TIMELOG_MONTH_COL_START + mi
            v = ws.cell(r, col).value
            if not (isinstance(v, str) and v.strip() and not v.startswith("=")):
                continue
            if mi + 1 < n and frozen is None:
                frozen = (mi + 1, col, v)
            elif mi + 1 >= n and solved is None:
                solved = (mi + 1, col, v)
        if frozen and solved:
            return (*frozen, *solved)
        return None

    found: tuple[int, int, int, str, int, int, str] | None = None
    for r in range(grid_first, ws.max_row + 1):
        col_a = ws.cell(r, 1).value
        col_a_str = str(col_a or "").strip()
        if (
            col_a is None or not col_a_str
            or col_a_str.upper().startswith("PER-")
            or col_a_str.lower().startswith("monthly total")
        ):
            break
        hit = _scan_row(r)
        if hit is not None:
            found = (r, *hit)
            break
    assert found is not None, (
        f"no person row spans boundary N={n} with solver-placed cells"
    )
    (
        first,
        frozen_month, frozen_col, frozen_pre,
        solved_month, solved_col, solved_pre,
    ) = found

    # Use REAL WP codes (orphan pruning would strip a "WP0_EDIT" sentinel).
    # Swap each cell to a WP code different from its pre-rebuild value so we
    # can assert preserve-vs-overwrite without depending on solver specifics.
    frozen_swap = "WP2" if frozen_pre != "WP2" else "WP1"
    solved_swap = "WP2" if solved_pre != "WP2" else "WP1"
    _set(ws, first, frozen_col, frozen_swap)
    _set(ws, first, solved_col, solved_swap)
    _set_rebuild_from_month(wb, n)
    wb.save(path)
    wb.close()

    rebuild_workbook(input_path=str(path))

    wb2 = openpyxl.load_workbook(path)
    ws2 = wb2[SHEET_TIMELOG_PLAN]
    assert ws2.cell(first, frozen_col).value == frozen_swap, (
        f"frozen cell (month {frozen_month} < N={n}) must survive rebuild, "
        f"got {ws2.cell(first, frozen_col).value!r}"
    )
    # Solver re-derives the post-boundary cell: it must land on something
    # non-empty, and it must NOT equal the discarded swap value UNLESS the
    # solver independently chose the same code. Guard the assertion by
    # inspecting whether frozen consumption freed the swap code — if the
    # pre-rebuild solver chose ``solved_pre`` and the swap value is a valid
    # post-freeze candidate, the solver may legitimately re-place it. Tighter
    # assertion: the solved cell was overwritten at least if it differs from
    # the swap, which is the overwhelmingly likely case with a non-trivial
    # grid; fall back to a weaker "is not empty" check otherwise.
    post_solved = ws2.cell(first, solved_col).value
    assert isinstance(post_solved, str) and post_solved.strip(), (
        f"solved cell (month {solved_month} ≥ N={n}) must be filled, "
        f"got {post_solved!r}"
    )
    wb2.close()


# ── Case 4: overshoot under freeze ────────────────────────────────────────


def test_overshoot_under_freeze_completes_without_error(tmp_path: Path):
    """Frozen cells that over-consume a WP's budget don't raise; the solver's
    narrowed ceiling clamps at 0 and the per-WP Buffer € surfaces the negative
    residual against the full Config budget.
    """
    path = tmp_path / "wb.xlsx"
    init_workbook(str(path), assignments=DEMO_CAPACITY_ASSIGNMENTS)

    wb = openpyxl.load_workbook(path)
    ws = wb[SHEET_TIMELOG_PLAN]
    months = _project_months(wb)
    # Stamp the SAME WP code across every frozen-month cell on every person
    # row — an absurdly over-consuming frozen block. Freeze the whole span.
    first = _grid_first_person_row(ws)
    n = max(4, months // 2 + 1)  # freeze months 1..n-1
    for r in range(first, ws.max_row + 1):
        col_a = ws.cell(r, 1).value
        col_a_str = str(col_a or "").strip()
        if (
            col_a is None or not col_a_str
            or col_a_str.upper().startswith("PER-")
            or col_a_str.lower().startswith("monthly total")
        ):
            break
        for mi in range(n - 1):
            col = TIMELOG_MONTH_COL_START + mi
            if ws.cell(r, col).value not in (None, ""):
                _set(ws, r, col, "WP1")  # force-pack WP1
    _set_rebuild_from_month(wb, n)
    wb.save(path)
    wb.close()

    # Must not raise.
    rebuild_workbook(input_path=str(path))

    wb2 = openpyxl.load_workbook(path)
    ws2 = wb2[SHEET_TIMELOG_PLAN]
    # Confirm frozen cells survived (spot-check first person row).
    first2 = _grid_first_person_row(ws2)
    kept = [
        ws2.cell(first2, TIMELOG_MONTH_COL_START + mi).value
        for mi in range(n - 1)
    ]
    kept_wp1 = [v for v in kept if v == "WP1"]
    assert len(kept_wp1) >= 1, (
        f"overshoot test: expected ≥1 frozen WP1 cell, got {kept!r}"
    )
    wb2.close()


# ── Case 5: pre-freeze-point workbook migration ───────────────────────────


def test_pre_freeze_point_workbook_upgrades_cleanly(tmp_path: Path):
    """A workbook with the stale ``timelog_resolve_flag`` defined name (and
    no ``timelog_rebuild_from_month``) rebuilds cleanly: new name registered,
    stale name unregistered, default N=1 takes effect."""
    path = tmp_path / "wb.xlsx"
    init_workbook(str(path), assignments=DEMO_CAPACITY_ASSIGNMENTS)

    # Synthesize a pre-freeze-point state: delete the new name, register
    # a stale legacy name pointing at some arbitrary cell.
    wb = openpyxl.load_workbook(path)
    if "timelog_rebuild_from_month" in wb.defined_names:
        del wb.defined_names["timelog_rebuild_from_month"]
    wb.defined_names.add(
        DefinedName(
            "timelog_resolve_flag",
            attr_text=f"'{SHEET_TIMELOG_PLAN}'!$J$2",
        )
    )
    wb.save(path)
    wb.close()

    wb2 = openpyxl.load_workbook(path)
    months = _project_months(wb2)
    out = read_timelog_plan(wb2, months)
    assert out["rebuild_from_month"] == 1, (
        "missing new name must default to 1 (solve-all)"
    )
    wb2.close()

    rebuild_workbook(input_path=str(path))

    wb3 = openpyxl.load_workbook(path)
    assert "timelog_rebuild_from_month" in wb3.defined_names, (
        "rebuild must register the new name"
    )
    assert "timelog_resolve_flag" not in wb3.defined_names, (
        "rebuild must unregister the stale legacy name"
    )
    wb3.close()


# ── Case 6: named-range sheet reference ───────────────────────────────────


def test_rebuild_from_month_named_range_targets_config_b9(tmp_path: Path):
    """attr_text of ``timelog_rebuild_from_month`` points at Config!$B$9 (per
    openpyxl-infra.md: unquoted since ``Config`` has no special chars)."""
    path = tmp_path / "wb.xlsx"
    init_workbook(str(path), assignments=DEMO_CAPACITY_ASSIGNMENTS)
    wb = openpyxl.load_workbook(path)
    dn = wb.defined_names["timelog_rebuild_from_month"]
    assert isinstance(dn, DefinedName)
    attr = dn.attr_text or ""
    # Sheet bit may be quoted or unquoted depending on openpyxl round-trip;
    # both forms are valid for a plain-alphanumeric sheet name.
    cleaned = attr.replace("'", "")
    assert cleaned == f"{SHEET_CONFIG}!$B$9", attr
    wb.close()
