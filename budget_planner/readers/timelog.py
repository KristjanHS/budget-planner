"""Read the Timelog Plan sheet — preserved per-person grid + rebuild-from-month."""

import re

from budget_planner.layout import (
    SECTION_TITLE_GRID,
    SHEET_CONFIG,
    SHEET_TIMELOG_PLAN,
    TIMELOG_MONTH_COL_START,
)

__all__ = ["read_timelog_plan"]


_MONTH_ABBREV_TO_NUM = {
    "jan": 1,
    "feb": 2,
    "mar": 3,
    "apr": 4,
    "may": 5,
    "jun": 6,
    "jul": 7,
    "aug": 8,
    "sep": 9,
    "oct": 10,
    "nov": 11,
    "dec": 12,
}


def _read_vacation_month_num(ws) -> int:
    """Resolve the ``vacation_month`` named range to a 1-indexed month int.

    Defaults to 7 (July) when the named range is missing, the cell is blank,
    or the value is unparseable — this is the one-shot migration for
    pre-enhancement workbooks.
    """
    wb = ws.parent
    if wb is None:
        return 7
    try:
        dn = wb.defined_names["vacation_month"]
    except (KeyError, TypeError):
        return 7
    if dn is None:
        return 7
    try:
        destinations = list(dn.destinations)
    except Exception:
        return 7
    if not destinations:
        return 7
    sheet_name, ref = destinations[0]
    if sheet_name not in wb.sheetnames:
        return 7
    try:
        val = wb[sheet_name][ref.replace("$", "")].value
    except Exception:
        return 7
    if isinstance(val, (int, float)) and not isinstance(val, bool):
        m = int(val)
        return m if 1 <= m <= 12 else 7
    if isinstance(val, str):
        s = val.strip().lower()[:3]
        return _MONTH_ABBREV_TO_NUM.get(s, 7)
    return 7


def _read_rebuild_from_month(wb) -> int:
    """Read the ``timelog_rebuild_from_month`` named-range cell.

    Returns 1 (solve-all — equivalent to the old ``resolve_flag=TRUE``
    default) when the named range is missing, the cell is blank, or the
    value is unparseable — safe default for pre-freeze-point workbooks
    and corruption cases. Shape mirrors ``_read_vacation_month_num``.
    """
    try:
        dn = wb.defined_names["timelog_rebuild_from_month"]
    except (KeyError, TypeError):
        return 1
    if dn is None:
        return 1
    try:
        destinations = list(dn.destinations)
    except Exception:
        return 1
    if not destinations:
        return 1
    sheet_name, ref = destinations[0]
    if sheet_name not in wb.sheetnames:
        return 1
    try:
        val = wb[sheet_name][ref.replace("$", "")].value
    except Exception:
        return 1
    if isinstance(val, bool):
        # bool is an int subclass — treat True/False as unparseable.
        return 1
    if isinstance(val, (int, float)):
        n = int(val)
        return max(1, n)
    if isinstance(val, str):
        s = val.strip()
        if not s:
            return 1
        try:
            return max(1, int(float(s)))
        except ValueError:
            return 1
    return 1


def read_timelog_plan(wb, project_months: int) -> dict:
    """Read preserved grid values and rebuild-from-month N from the Timelog Plan sheet.

    v2 surface: the merged "TIMELOG GRID" band holds literal WP codes (no
    `=<Proposed>` formulas anymore). The reader returns the ENTIRE preserved
    grid unconditionally; filtering against the freeze boundary is the
    builder's job (see ``docs/spec.md`` § Timelog Plan merged grid).

    Anchoring strategy: scan col A for the exact GRID section-title string
    (imported from the builder so edits don't drift). The band layout is
    ``title, year, month, midx, persons...``; the first person row is
    anchor + 4. The rebuild-from-month integer lives at the
    ``timelog_rebuild_from_month`` named range on Config (defaults to 1 —
    solve-all — when the name is absent on pre-freeze-point workbooks).

    Args:
        wb: Workbook (load with ``data_only=False`` so formulas are visible).
        project_months: Project month count — fixes the grid column span.

    Returns:
        Dict with keys:
          - ``grid``: ``{(person_name, month_num_1_indexed): wp_code}``, one
            entry per non-empty grid cell.
          - ``rebuild_from_month``: int, read from the Config named range or
            1 default.
        Returns ``{"grid": {}, "rebuild_from_month": 1}`` when the sheet is
        missing or the anchor cannot be located (degraded-but-safe).
    """
    rebuild_from_month = _read_rebuild_from_month(wb)
    empty: dict = {"grid": {}, "rebuild_from_month": rebuild_from_month}

    if SHEET_TIMELOG_PLAN not in wb.sheetnames:
        return empty
    ws = wb[SHEET_TIMELOG_PLAN]

    # Scan col A for the exact GRID section-title string.
    grid_title_row: int | None = None
    for r in range(1, ws.max_row + 1):
        val = ws.cell(r, 1).value
        if isinstance(val, str) and val == SECTION_TITLE_GRID:
            grid_title_row = r
            break
    if grid_title_row is None:
        return empty

    # GRID band layout relative to the title row:
    #   +1 year, +2 month, +3 midx, +4.. persons.
    person_first_row = grid_title_row + 4

    month_col_start = TIMELOG_MONTH_COL_START

    config_ws = wb[SHEET_CONFIG] if SHEET_CONFIG in wb.sheetnames else None

    grid: dict[tuple[str, int], str] = {}
    for r in range(person_first_row, ws.max_row + 1):
        col_a = ws.cell(r, 1).value
        if col_a is None:
            break
        col_a_str = str(col_a).strip()
        if not col_a_str:
            break
        # PER-PERSON SUMMARY title (or any PER- prefix) ends the grid band.
        # "Monthly total" label also ends the band — its cells hold formulas,
        # not preserved literals. The prefix guard breaks on the under-grid
        # "Monthly total -i" row; the always-visible "Monthly total +i" row
        # sits below a blank spacer further down and is never reached
        # because the -i row terminates the scan first.
        if col_a_str.upper().startswith("PER-"):
            break
        if col_a_str.lower().startswith("monthly total"):
            break

        # Resolve name: either =Config!A{n} cross-ref or a literal fallback.
        name: str | None = None
        if col_a_str.startswith("="):
            m = re.match(rf"={SHEET_CONFIG}!A(\d+)", col_a_str)
            if m and config_ws is not None:
                resolved = config_ws.cell(int(m.group(1)), 1).value
                if resolved:
                    name = str(resolved).strip()
        else:
            name = col_a_str
        if not name:
            continue

        for mi in range(project_months):
            col = month_col_start + mi
            val = ws.cell(r, col).value
            if val is None:
                continue
            if isinstance(val, str):
                s = val.strip()
                # Defensive: formulas shouldn't appear in v2 grid cells,
                # but skip them if a user pasted one in.
                if not s or s.startswith("="):
                    continue
                grid[(name, mi + 1)] = s
            else:
                # Defensive: numeric shouldn't happen (cells hold WP code
                # strings) but preserve rather than silently drop.
                grid[(name, mi + 1)] = str(val)

    return {"grid": grid, "rebuild_from_month": rebuild_from_month}
