"""Read the Capacity sheet — per-person WP/external grids and Capacity-side
config blocks (assignments + externals)."""

import re

from budget_planner.layout import (
    COL_ASSIGN_END,
    COL_ASSIGN_FTE,
    COL_ASSIGN_PERSON,
    COL_ASSIGN_START,
    COL_ASSIGN_WP,
    COL_EXT_END,
    COL_EXT_EU,
    COL_EXT_FTE,
    COL_EXT_PERSON,
    COL_EXT_PROJECT,
    COL_EXT_START,
    MONTH_COL_START,
    SECTION_HEADER_LABELS,
    SHEET_CONFIG,
    SHEET_PLAN,
)

from .config import read_assignments_from_config, read_externals_from_config
from .layout_discovery import (
    _data_row_after,
    _find_header_row,
    _parse_config_float,
    _read_monthly_values,
    _resolve_plan_ref,
)

__all__ = [
    "read_capacity_sheet",
    "read_capacity_assignments",
]


def read_capacity_assignments(ws) -> list[dict]:
    """Read Person-WP Assignments config block from Capacity sheet.

    Returns:
        List of dicts with keys: person, wp, total_fte, start, end.
        Empty list if the block is not found.
    """
    assignments_header = _find_header_row(
        ws, SECTION_HEADER_LABELS["assignments"], required=False
    )
    if assignments_header is None:
        return []

    # Skip section header + column headers
    r = _data_row_after(assignments_header)
    assignments: list[dict] = []
    while r <= ws.max_row:
        person = ws.cell(r, COL_ASSIGN_PERSON).value
        if person is None or str(person).strip() == "":
            break
        wp = ws.cell(r, COL_ASSIGN_WP).value
        total_fte = ws.cell(r, COL_ASSIGN_FTE).value
        total_fte_parsed = _parse_config_float(total_fte)
        if total_fte_parsed is None:
            break
        start = ws.cell(r, COL_ASSIGN_START).value
        end = ws.cell(r, COL_ASSIGN_END).value
        assignments.append(
            {
                "person": str(person).strip(),
                "wp": str(wp or "").strip(),
                "total_fte": total_fte_parsed,
                "start": int(start or 1),
                "end": int(end or 1),
            }
        )
        r += 1

    return assignments


def _read_externals_from_capacity(ws) -> list[dict]:
    """Read the EXTERNAL COMMITMENTS config block from the Capacity sheet.

    Returns:
        List of dicts with keys: person, project, fte, start, end, is_eu.
        Empty list if the section is not found.
    """
    ext_config_header = _find_header_row(
        ws, SECTION_HEADER_LABELS["externals"], required=False
    )
    if ext_config_header is None:
        return []

    external_commitments: list[dict] = []
    # Skip section header + column headers
    r = _data_row_after(ext_config_header)
    while r <= ws.max_row:
        person = ws.cell(r, COL_EXT_PERSON).value
        project = ws.cell(r, COL_EXT_PROJECT).value
        if person is None or str(person).strip() == "":
            break
        fte = ws.cell(r, COL_EXT_FTE).value
        fte_parsed = _parse_config_float(fte)
        if fte_parsed is None:
            break
        start = ws.cell(r, COL_EXT_START).value
        end = ws.cell(r, COL_EXT_END).value
        is_eu_val = ws.cell(r, COL_EXT_EU).value
        external_commitments.append(
            {
                "person": str(person).strip(),
                "project": str(project or "").strip(),
                "fte": fte_parsed,
                "start": int(start or 1),
                "end": int(end or 1),
                "is_eu": str(is_eu_val).strip().upper() in ("YES", "TRUE", "1"),
            }
        )
        r += 1

    return external_commitments


def _read_person_grids(
    ws,
    wp_ids: set[str],
    project_months: int,
    month_col_start: int = MONTH_COL_START,
) -> dict[str, dict]:
    """Find person header rows and parse per-person WP/external grids.

    Args:
        ws: Capacity worksheet.
        wp_ids: Set of known WP IDs (from config_wps).
        project_months: Number of project months.
        month_col_start: First month column (default from layout).

    Returns:
        Dict mapping person name to
        ``{"wp_values": {wp_id: [monthly]}, "external_values": {proj: [monthly]}}``.
    """
    # ── Find month header row: "WP ID" in col A ──
    month_header_row = _find_header_row(ws, SECTION_HEADER_LABELS["wp_id_column"])

    # ── Find person header rows ──
    # Person headers have bold font + SECTION_FILL (fgColor "2F5496")
    # and formula like =Config!A{n}
    person_headers: list[int] = []
    for r in range(month_header_row + 1, ws.max_row + 1):
        cell = ws.cell(r, 1)
        font = cell.font
        fill = cell.fill
        fg = ""
        if fill and fill.fgColor and fill.fgColor.rgb:
            fg = str(fill.fgColor.rgb).upper()
        if font.bold and "2F5496" in fg:
            person_headers.append(r)

    # Fallback: pattern-based detection (scan for TOTAL rows, person header
    # is above each group)
    if not person_headers:
        for r in range(month_header_row + 1, ws.max_row + 1):
            if ws.cell(r, 2).value and str(ws.cell(r, 2).value).upper() == "TOTAL":
                # Scan backward to find group start
                for rr in range(r - 1, month_header_row, -1):
                    cell_a = ws.cell(rr, 1).value
                    if cell_a is not None and ws.cell(rr, 2).value is None:
                        # Person header (merged cells: col A has name,
                        # col B is None/empty from merge)
                        if rr not in person_headers:
                            person_headers.append(rr)
                        break
        person_headers.sort()

    # ── Resolve helpers (closures over ws.parent sheets) ──
    wb = ws.parent
    config_ws = wb[SHEET_CONFIG] if SHEET_CONFIG in wb.sheetnames else None
    plan_ws = wb[SHEET_PLAN]

    def _resolve_person_name(cell_value):
        """Resolve person name from formula or literal value."""
        if cell_value is None:
            return None
        val = str(cell_value)
        m = re.match(rf"={SHEET_CONFIG}!A(\d+)", val)
        if m and config_ws:
            config_row = int(m.group(1))
            name_val = config_ws.cell(config_row, 1).value
            return str(name_val) if name_val else None
        if not val.startswith("="):
            return val
        return None

    def _resolve_wp_id(cell_value):
        """Resolve WP ID from formula or literal, validated against known WP IDs."""
        resolved = _resolve_plan_ref(cell_value, plan_ws, config_ws)
        if resolved is not None and resolved in wp_ids:
            return resolved
        return None

    # ── Parse each person group ──
    person_groups: dict[str, dict] = {}

    for hi, header_row in enumerate(person_headers):
        name = _resolve_person_name(ws.cell(header_row, 1).value)
        if name is None:
            continue

        # Find the end of this person's group (next person header or end of sheet)
        if hi + 1 < len(person_headers):
            group_end = person_headers[hi + 1] - 1
        else:
            group_end = ws.max_row

        wp_values: dict[str, list] = {}
        external_values: dict[str, list] = {}

        # Scan data rows between header+1 and group_end
        for dr in range(header_row + 1, group_end + 1):
            col_a = ws.cell(dr, 1).value
            col_b = ws.cell(dr, 2).value

            # Skip TOTAL and REMAINING rows
            if col_b and str(col_b).upper() in ("TOTAL", "REMAINING"):
                continue

            # Try to classify the row
            resolved_wp = _resolve_wp_id(col_a)
            if resolved_wp is not None:
                # Programme WP row
                monthly = _read_monthly_values(ws, dr, month_col_start, project_months)
                wp_values[resolved_wp] = monthly
            elif col_a is not None:
                # Check if it's an external project (literal string, not formula
                # pointing to Plan)
                col_a_str = str(col_a)
                if not col_a_str.startswith("="):
                    # External project row
                    monthly = _read_monthly_values(
                        ws, dr, month_col_start, project_months
                    )
                    external_values[col_a_str] = monthly

        person_groups[name] = {
            "wp_values": wp_values,
            "external_values": external_values,
        }

    return person_groups


def read_capacity_sheet(
    ws, config_wps, project_months=None, month_col_start=MONTH_COL_START
) -> dict:
    """Read Capacity sheet person groups.

    Prefers the Config-sheet assignments/externals sections when present:
    ``ws.parent[SHEET_CONFIG]`` is inspected first via
    ``read_assignments_from_config`` / ``read_externals_from_config``. Only if
    those return empty does the reader fall back to scanning Capacity itself
    (legacy workbooks pre-dating the Config sections).

    Args:
        ws: Capacity worksheet
        config_wps: list of WP dicts from read_plan_config (used to classify rows)
        project_months: Number of project months. If None, falls back to
            config.PROJECT_MONTHS for backward compatibility.
        month_col_start: First month column (default 3 = col C).

    Returns:
        {
            "assignments": [{"person": ..., "wp": ..., ...}],
            "external_commitments": [{"person": ..., "project": ..., "fte": ...,
                                      "start": ..., "end": ..., "is_eu": ...}],
            "person_groups": {
                "Alice": {
                    "wp_values": {"WP1": [60 monthly values], ...},
                    "external_values": {"ProjectX": [60 values], ...},
                },
                ...
            }
        }
    """
    if project_months is None:
        from budget_planner.config import PROJECT_MONTHS

        project_months = PROJECT_MONTHS

    wp_ids = {wp["id"] for wp in config_wps}

    # Try Config-based readers first; Capacity-based fallback below.
    wb = ws.parent
    config_ws = wb[SHEET_CONFIG] if SHEET_CONFIG in wb.sheetnames else None
    assignments: list[dict] = []
    external_commitments: list[dict] = []
    if config_ws is not None:
        assignments = read_assignments_from_config(config_ws)
        external_commitments = read_externals_from_config(config_ws)

    # Fall back to Capacity-based readers if Config sections not found
    if not assignments:
        assignments = read_capacity_assignments(ws)
    if not external_commitments:
        external_commitments = _read_externals_from_capacity(ws)
    person_groups = _read_person_grids(ws, wp_ids, project_months, month_col_start)

    return {
        "assignments": assignments,
        "external_commitments": external_commitments,
        "person_groups": person_groups,
    }
