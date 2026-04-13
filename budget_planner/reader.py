"""Read existing workbooks for round-trip updates."""

import re
from decimal import Decimal
from typing import Literal, overload

from openpyxl.utils import get_column_letter

from budget_planner.layout import (
    MONTH_COL_START,
    SHEET_PLAN,
    SHEET_CONFIG,
    SECTION_HEADER_LABELS,
    SECTION_HEADER_ROWS,
    COL_WP_FTE,
    COL_WP_RATE,
    COL_SCHED_START,
    COL_SCHED_END,
    COL_SCHED_TYPE,
    COL_ROSTER_FTE,
    COL_ROSTER_GRADE,
    COL_GRADE_RATE,
    COL_RP_START,
    COL_RP_END,
    COL_ASSIGN_PERSON,
    COL_ASSIGN_WP,
    COL_ASSIGN_FTE,
    COL_ASSIGN_START,
    COL_ASSIGN_END,
    COL_EXT_PERSON,
    COL_EXT_PROJECT,
    COL_EXT_FTE,
    COL_EXT_START,
    COL_EXT_END,
    COL_EXT_EU,
)


def _safe_float(value, default: float) -> float:
    """Convert value to float, returning default if it's a formula or None."""
    if value is None:
        return default
    try:
        return float(value)
    except (ValueError, TypeError):
        return default


def _parse_config_float(value) -> float | None:
    """Parse Config numeric cells: real numbers and text (e.g. European ``0,35``).

    Excel often stores locale-formatted values as *text* (green triangle). The
    roster reader used to stop at the first such cell, so later names were
    omitted from ``team`` and assignments looked "orphaned".

    Returns:
        ``0.0`` when *value* is ``None`` (empty cell on a row the caller has
        already decided is valid). A float when *value* is numeric or a
        parseable numeric string (US/EU locale, NBSP-tolerant). ``None`` when
        *value* is a bool, a formula string (``=...``), or any other non-numeric
        string — signalling the caller to stop scanning the section.
    """
    # bool is a subclass of int; reject it explicitly before the int/float branch
    # so True/False don't silently coerce to 1.0/0.0 in a numeric context.
    if isinstance(value, bool):
        return None
    if value is None:
        return 0.0
    if isinstance(value, (int, float)):
        return float(value)
    if isinstance(value, Decimal):
        return float(value)
    if not isinstance(value, str):
        return None
    s = value.strip()
    if not s:
        # Empty/whitespace string: treat as "non-numeric sentinel" (matches pre-
        # _parse_config_float break-on-type-guard behaviour). Returning 0.0 here
        # would append a zero-capacity ghost row for a named person with a blank
        # FTE cell.
        return None
    if s.startswith("="):
        return None
    s = s.replace("\u00a0", "").replace(" ", "")
    norm = s
    if "," in norm and "." not in norm:
        norm = norm.replace(",", ".")
    elif "," in norm and "." in norm:
        if norm.rfind(",") > norm.rfind("."):
            norm = norm.replace(".", "").replace(",", ".")
        else:
            norm = norm.replace(",", "")
    try:
        return float(norm)
    except ValueError:
        return None


def _parse_fte_rate(value) -> float:
    """Parse the FTE Rate cell value as a numeric fallback.

    The builder always regenerates B4 as a SUMPRODUCT over the live grade/roster
    layout, so formula strings in the cell are intentionally discarded — the
    rate is a readout of the grade table, not an independent input. Numeric
    values (and None) are coerced to float, defaulting to 7500.
    """
    if isinstance(value, (int, float)):
        return float(value or 7500)
    return 7500.0


@overload
def _find_header_row(
    ws,
    pattern: str,
    col: int = ...,
    max_row: int | None = ...,
    required: Literal[True] = ...,
) -> int: ...


@overload
def _find_header_row(
    ws,
    pattern: str,
    col: int = ...,
    max_row: int | None = ...,
    required: Literal[False] = ...,
) -> int | None: ...


def _find_header_row(ws, pattern, col=1, max_row=None, required=True):
    """Find row where cell exactly matches pattern (case-insensitive, stripped).

    Uses exact match (not substring) to avoid layout-drift silent failures —
    e.g. a renamed or nested section header won't accidentally satisfy the search.

    Args:
        ws: openpyxl worksheet.
        pattern: String to search for (exact match after strip + upper-case).
        col: Column number to scan (default 1 = col A).
        max_row: Upper row limit. If None, scans the whole sheet.
        required: If True and pattern not found, raises ValueError.

    Returns:
        Row number (int) or None if not found and required=False.

    Raises:
        ValueError: If required=True and pattern is not found.
    """
    limit = max_row if max_row is not None else ws.max_row
    target = pattern.strip().upper()
    for r in range(1, limit + 1):
        val = ws.cell(r, col).value
        if val is not None and str(val).strip().upper() == target:
            return r
    if required:
        raise ValueError(
            f"Section header {pattern!r} not found in column {col} "
            f"(rows 1-{limit}). Was the workbook structure modified?"
        )
    return None


def _data_row_after(header_row: int, header_rows: int = SECTION_HEADER_ROWS) -> int:
    """First data row after a section header (title row + column header row)."""
    return header_row + header_rows


_NAMED_RANGE_REF_RE = re.compile(r"\$?[A-Z]+\$?(\d+)(?::\$?[A-Z]+\$?(\d+))?$")


def _iter_named_range_rows(wb, name: str) -> tuple[int, int] | None:
    """Resolve a defined name to its (first_row, last_row) on the first target range.

    Returns None if the name is not defined or cannot be parsed. Callers fall back
    to sentinel-scan reading when this returns None. Excel auto-extends defined
    names on row insertion, so this is the authoritative live span after
    Excel-side edits.
    """
    dn = None
    try:
        dn = wb.defined_names[name]
    except (KeyError, TypeError):
        return None
    if dn is None:
        return None
    try:
        destinations = list(dn.destinations)
    except Exception:
        return None
    if not destinations:
        return None
    _sheet, ref = destinations[0]
    m = _NAMED_RANGE_REF_RE.match(ref)
    if not m:
        return None
    first = int(m.group(1))
    last = int(m.group(2)) if m.group(2) else first
    if last < first:
        return None
    return (first, last)


def read_config_sheet(ws) -> tuple[dict, list[dict], list[dict]]:
    """Read Config sheet -> (project_params, work_packages, wp_schedule).

    Returns:
        Tuple of (project_params, work_packages, wp_schedule) where:
        - project_params: dict with keys 'title', 'start', 'months', 'rate',
          'eu_ceiling_pm'. ``rate`` is a numeric fallback only — the builder
          always regenerates B4 from the live grade/roster layout.
        - work_packages: list of dicts with keys 'id', 'name', 'fte', 'rate',
          'rate_raw'
        - wp_schedule: list of dicts with keys 'id', 'start', 'end', 'type'
    """
    # Scan column A for known param labels (position-independent).
    _param_map = {
        "Project Title": ("title", None),
        "Project Start": ("start", None),
        "Project Months": ("months", lambda v: int(v or 60)),
        "FTE Rate": ("rate", _parse_fte_rate),
        "EU Annual Ceiling": ("eu_ceiling_pm", lambda v: float(v or 12.0)),
    }
    project_params: dict = {}
    matched_keys: set[str] = set()
    for r in range(1, ws.max_row + 1):
        label = str(ws.cell(r, 1).value or "")
        for prefix, (key, converter) in _param_map.items():
            if prefix.lower() in label.lower():
                raw = ws.cell(r, 2).value
                project_params[key] = converter(raw) if converter else raw
                matched_keys.add(key)
                break
        if len(matched_keys) == len(_param_map):
            break
    # Validate all params found
    for prefix, (key, _) in _param_map.items():
        if key not in project_params:
            raise ValueError(f"Config sheet missing param row matching '{prefix}'")

    # WP Budget table — find "WP BUDGET" section header
    wp_section = _find_header_row(ws, SECTION_HEADER_LABELS["wp_budget"])
    work_packages: list[dict] = []
    r = _data_row_after(wp_section)
    while r <= ws.max_row:
        wp_id = ws.cell(r, 1).value
        col_b = ws.cell(r, 2).value
        if wp_id is None or str(wp_id).strip() == "":
            break
        # Stop at TOTAL row
        if col_b and str(col_b).upper() == "TOTAL":
            break
        rate_cell = ws.cell(r, COL_WP_RATE).value
        work_packages.append(
            {
                "id": str(wp_id),
                "name": str(col_b or ""),
                "fte": float(ws.cell(r, COL_WP_FTE).value or 0),
                # ``rate`` is the numeric coercion for downstream consumers (falls
                # back to 6500 when the cell is a formula). ``rate_raw`` preserves
                # the literal cell value — formula string or numeric — so that
                # writers can restore the user's formula on round-trip.
                "rate": _safe_float(rate_cell, 6500),
                "rate_raw": rate_cell,
            }
        )
        r += 1

    # WP Schedule table — find "WP SCHEDULE" section header (optional)
    sched_section = _find_header_row(
        ws, SECTION_HEADER_LABELS["wp_schedule"], required=False
    )
    wp_schedule: list[dict] = []
    if sched_section is not None:
        r = _data_row_after(sched_section)
        while r <= ws.max_row:
            sched_id = ws.cell(r, 1).value
            if sched_id is None or str(sched_id).strip() == "":
                break
            wp_schedule.append(
                {
                    "id": str(sched_id),
                    "start": int(ws.cell(r, COL_SCHED_START).value or 1),
                    "end": int(ws.cell(r, COL_SCHED_END).value or 1),
                    "type": str(ws.cell(r, COL_SCHED_TYPE).value or "flat"),
                }
            )
            r += 1

    return project_params, work_packages, wp_schedule


def read_plan_config(ws, layout) -> list[dict]:
    """Read WP definitions from Config sheet (via Plan layout cross-refs).

    After the Plan config table was removed, WP data lives on the Config sheet.
    This function reads from Config using the config_first_wp_row from layout.
    """
    wb = ws.parent
    config_ws = wb[SHEET_CONFIG] if SHEET_CONFIG in wb.sheetnames else None
    if config_ws is None:
        raise ValueError("Workbook has no Config sheet")

    wps = []
    for row in range(layout["config_first_wp_row"], layout["config_last_wp_row"] + 1):
        wp_id = config_ws.cell(row, 1).value
        wp_name = config_ws.cell(row, 2).value
        fte = config_ws.cell(row, 3).value
        if wp_id is None:
            continue
        fte_parsed = _parse_config_float(fte)
        wps.append(
            {
                "id": str(wp_id).strip(),
                "name": str(wp_name or "").strip(),
                "fte": fte_parsed if fte_parsed is not None else 0.0,
            }
        )
    if not wps:
        raise ValueError(
            "Config table has zero WP rows. "
            "Regenerate from defaults with: python -m budget_planner"
        )
    return wps


def read_grade_rates(ws) -> dict[str, int | float]:
    """Read grade-to-rate mapping from the Config sheet's RATE BY GRADE section.

    Args:
        ws: Config worksheet.

    Returns:
        Dict mapping grade name to EUR/PM rate
        (e.g. {"Senior": 12500, "Mid": 10000, "Junior": 7500}).
        Returns default rates from config.py if section not found
        (backward compat with old workbooks).
    """
    section_row = _find_header_row(
        ws, SECTION_HEADER_LABELS["rate_by_grade"], required=False
    )
    if section_row is None:
        from budget_planner.config import GRADE_RATES

        return dict(GRADE_RATES)

    # Skip section header + column headers
    r = _data_row_after(section_row)
    rates: dict[str, int | float] = {}
    while r <= ws.max_row:
        grade_name = ws.cell(r, 1).value
        if grade_name is None or str(grade_name).strip() == "":
            break
        rate_val = ws.cell(r, COL_GRADE_RATE).value
        if rate_val is None or not isinstance(rate_val, (int, float)):
            break
        rates[str(grade_name).strip()] = rate_val
        r += 1

    return rates


def read_team_from_config(ws) -> list[dict]:
    """Read team roster from the Config sheet's TEAM ROSTER section.

    Args:
        ws: Config worksheet.

    Returns:
        List of dicts with keys 'name', 'max_fte', and 'grade'.
    """
    from budget_planner.config import DEFAULT_GRADE

    # Find "TEAM ROSTER" section header
    roster_section = _find_header_row(
        ws, SECTION_HEADER_LABELS["team_roster"], required=True
    )

    # TEAM ROSTER has a 1-row header (title + column-headers on the same row pair
    # where the section label row IS the title and the next row is the column
    # header). Data starts at person_header + 1 — do NOT use _data_row_after here.
    person_header = roster_section + 1

    # Detect whether Grade column exists (col C header)
    has_grade_col = False
    grade_hdr = ws.cell(person_header, COL_ROSTER_GRADE).value
    if grade_hdr and "grade" in str(grade_hdr).lower():
        has_grade_col = True

    # Use roster_ftes named range as authoritative bounds when present; within
    # the range, tolerate internal blanks (Excel-side mid-insert leaves gaps).
    # Past the range, fall back to sentinel scan to support append-style edits
    # the user may have made without manually extending the defined name.
    wb = ws.parent
    span = _iter_named_range_rows(wb, "roster_ftes") if wb is not None else None
    # Sanity check: the named range must not start before the first data row
    # (corrupt/stale workbooks otherwise yield an empty roster → rebuild would
    # silently drop everyone).
    if span is not None and span[0] < person_header + 1:
        span = None
    if span is not None:
        first_data, range_end = span
    else:
        first_data = person_header + 1
        range_end = None

    roster: list[dict] = []
    r = first_data
    while r <= ws.max_row:
        name = ws.cell(r, 1).value
        max_fte = ws.cell(r, COL_ROSTER_FTE).value
        max_fte_parsed = _parse_config_float(max_fte)
        name_blank = name is None or str(name).strip() == ""
        inside_range = range_end is not None and r <= range_end
        # Past range_end with no entries yet — the section is empty; break
        # rather than sentinel-scan into the next section's header.
        if range_end is not None and r > range_end and not roster:
            break
        if name_blank or max_fte_parsed is None:
            if inside_range:
                # Tolerate blanks inside the named range — Excel-side mid-insert.
                r += 1
                continue
            # Outside the range (or no range at all): sentinel termination.
            break
        grade = DEFAULT_GRADE
        if has_grade_col:
            grade_val = ws.cell(r, COL_ROSTER_GRADE).value
            if grade_val is not None and str(grade_val).strip():
                grade = str(grade_val).strip()
        roster.append(
            {
                "name": str(name).strip(),
                "max_fte": max_fte_parsed,
                "grade": grade,
            }
        )
        r += 1

    return roster


def read_reporting_periods(ws, fallback_months=None):
    """Read reporting period definitions from Config sheet.

    Args:
        ws: Config worksheet.
        fallback_months: If section not found, generate default periods
            for this many months. If None and section missing, returns [].

    Returns:
        List of dicts with keys: id, start, end.
    """
    rp_section = _find_header_row(
        ws, SECTION_HEADER_LABELS["reporting_periods"], required=False
    )
    if rp_section is None:
        if fallback_months is not None:
            from budget_planner.config import default_reporting_periods

            return default_reporting_periods(fallback_months)
        return []

    # Skip section header + column headers
    r = _data_row_after(rp_section)
    periods: list[dict] = []
    while r <= ws.max_row:
        period_id = ws.cell(r, 1).value
        if period_id is None or str(period_id).strip() == "":
            break
        start = ws.cell(r, COL_RP_START).value
        end = ws.cell(r, COL_RP_END).value
        periods.append(
            {
                "id": str(period_id),
                "start": int(start or 1),
                "end": int(end or 1),
            }
        )
        r += 1

    return periods


def read_assignments_from_config(ws) -> list[dict]:
    """Read Person-WP Assignments from Config sheet.

    Args:
        ws: Config worksheet.

    Returns:
        List of dicts with keys: person, wp, total_fte, start, end.
        Empty list if the section is not found.
    """
    assignments_header = _find_header_row(
        ws, SECTION_HEADER_LABELS["assignments"], required=False
    )
    if assignments_header is None:
        return []

    # Use assign_persons named range as authoritative span when present; skip
    # blanks inside the range (Excel mid-insert) and keep scanning past the
    # range end to support append-style edits that don't extend the name.
    wb = ws.parent
    span = _iter_named_range_rows(wb, "assign_persons") if wb is not None else None
    expected_first = _data_row_after(assignments_header)
    if span is not None and span[0] < expected_first:
        span = None
    if span is not None:
        first_data, range_end = span
    else:
        first_data = expected_first
        range_end = None

    assignments: list[dict] = []
    r = first_data
    while r <= ws.max_row:
        person = ws.cell(r, COL_ASSIGN_PERSON).value
        total_fte = ws.cell(r, COL_ASSIGN_FTE).value
        total_fte_parsed = _parse_config_float(total_fte)
        person_blank = person is None or str(person).strip() == ""
        inside_range = range_end is not None and r <= range_end
        if range_end is not None and r > range_end and not assignments:
            break
        if person_blank or total_fte_parsed is None:
            if inside_range:
                r += 1
                continue
            break
        wp = ws.cell(r, COL_ASSIGN_WP).value
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


def read_externals_from_config(ws) -> list[dict]:
    """Read External Commitments from Config sheet.

    Args:
        ws: Config worksheet.

    Returns:
        List of dicts with keys: person, project, fte, start, end, is_eu.
        Empty list if the section is not found.
    """
    ext_config_header = _find_header_row(
        ws, SECTION_HEADER_LABELS["externals"], required=False
    )
    if ext_config_header is None:
        return []

    # Use ext_persons named range as authoritative span when present; same
    # hybrid semantics as roster/assignments (tolerate internal blanks, allow
    # sentinel scan past range end for append-style edits).
    wb = ws.parent
    span = _iter_named_range_rows(wb, "ext_persons") if wb is not None else None
    expected_first = _data_row_after(ext_config_header)
    if span is not None and span[0] < expected_first:
        span = None
    if span is not None:
        first_data, range_end = span
    else:
        first_data = expected_first
        range_end = None

    external_commitments: list[dict] = []
    r = first_data
    while r <= ws.max_row:
        person = ws.cell(r, COL_EXT_PERSON).value
        fte = ws.cell(r, COL_EXT_FTE).value
        fte_parsed = _parse_config_float(fte)
        person_blank = person is None or str(person).strip() == ""
        inside_range = range_end is not None and r <= range_end
        if range_end is not None and r > range_end and not external_commitments:
            break
        if person_blank or fte_parsed is None:
            if inside_range:
                r += 1
                continue
            break
        project = ws.cell(r, COL_EXT_PROJECT).value
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

            # Skip TOTAL, REMAINING, and EU TOTAL rows
            if col_b and str(col_b).upper() in ("TOTAL", "REMAINING", "EU TOTAL"):
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


def _read_monthly_values(ws, row, col_start, num_months):
    """Read monthly values from a row, returning None for formulas/empty."""
    values = []
    for m in range(num_months):
        col = col_start + m
        val = ws.cell(row, col).value
        if val is not None and isinstance(val, str) and val.startswith("="):
            values.append(None)
        elif val is None or val == "":
            values.append(None)
        else:
            try:
                values.append(float(val))
            except (ValueError, TypeError):
                values.append(None)
    return values


def discover_layout(ws) -> dict:
    """Scan existing Plan worksheet, reconstruct layout dict.

    After the Plan config table was removed, the Plan sheet has:
    - Row 1: compact header
    - Row 2: blank
    - Row 3: PLANNED FTE section header
    - Rows 4-5: year/month headers
    - Rows 6+: WP data rows

    The config_first_wp_row/config_last_wp_row refer to the Config sheet's
    WP table rows (row 8+ by default).
    """
    # Find PLANNED block header
    planned_header_row = _find_header_row(
        ws, SECTION_HEADER_LABELS["planned_fte"], required=True
    )

    # Find the first WP data row after the PLANNED header.
    # WP rows have a formula (e.g. =Config!A8) or literal WP ID in column A.
    planned_first_wp_row = None
    for r in range(planned_header_row + 1, ws.max_row + 1):
        val = ws.cell(r, 1).value
        if val is None or str(val).strip() == "":
            continue
        val_str = str(val)
        # A formula referencing a Config row (e.g. =Config!A8) means WP data
        if val_str.startswith("="):
            planned_first_wp_row = r
            break
        # A literal WP ID (skip known header labels)
        if not any(
            kw in val_str.upper()
            for kw in ("PLANNED", "YEAR", "SUM", "BUDGET", "WP ID", "WP NAME")
        ):
            planned_first_wp_row = r
            break
    if planned_first_wp_row is None:
        raise ValueError(
            "Cannot find first WP data row in PLANNED block. "
            "Expected formula (=Config!A{n}) or literal WP ID after "
            f"PLANNED header at row {planned_header_row}."
        )

    # Count WP rows: scan until empty col A or TOTAL row
    planned_last_wp_row = planned_first_wp_row
    while True:
        next_val = ws.cell(planned_last_wp_row + 1, 1).value
        next_name = str(ws.cell(planned_last_wp_row + 1, 2).value or "")
        if not next_val or next_name.upper().startswith("TOTAL"):
            break
        planned_last_wp_row += 1
    num_wps = planned_last_wp_row - planned_first_wp_row + 1
    planned_total_row = planned_last_wp_row + 1

    # Derive Config sheet WP rows from the first WP formula reference
    # Formula pattern: =Config!A{row}
    first_val = ws.cell(planned_first_wp_row, 1).value
    cfg_wp_first = None
    if first_val and isinstance(first_val, str):
        m = re.match(rf"={SHEET_CONFIG}!A(\d+)", str(first_val))
        if m:
            cfg_wp_first = int(m.group(1))
    if cfg_wp_first is None:
        # Fallback: discover from Config sheet directly (legacy workbooks
        # with literal WP IDs instead of =Config!A{n} formulas)
        config_ws = (
            ws.parent[SHEET_CONFIG] if SHEET_CONFIG in ws.parent.sheetnames else None
        )
        wp_section = (
            _find_header_row(
                config_ws, SECTION_HEADER_LABELS["wp_budget"], required=False
            )
            if config_ws is not None
            else None
        )
        if wp_section is None:
            raise ValueError(
                "Cannot locate Config WP Budget section. "
                "Regenerate workbook with `python -m budget_planner init`."
            )
        cfg_wp_first = _data_row_after(wp_section)
    cfg_wp_last = cfg_wp_first + num_wps - 1

    return {
        "config_first_wp_row": cfg_wp_first,
        "config_last_wp_row": cfg_wp_last,
        "planned_header_row": planned_header_row,
        "planned_first_wp_row": planned_first_wp_row,
        "planned_last_wp_row": planned_last_wp_row,
        "planned_total_row": planned_total_row,
    }


def read_actuals_values(
    ws, project_months, month_col_start=MONTH_COL_START
) -> dict[str, list]:
    """Read Actuals EUR values for round-trip preservation.

    Args:
        ws: Actuals worksheet.
        project_months: Number of project months.
        month_col_start: First month column (default 3 = col C).

    Returns:
        Dict keyed by WP ID string -> list of monthly values (float or None).
        Every Actuals WP row encountered is preserved, including rows where
        all monthly cells are blank — a row with no monthly data still
        represents an explicit "user has touched this WP" intent (e.g.,
        zero-spend month, or a deliberate clear). Distinguishing that from
        "never entered" via openpyxl alone isn't possible, so we preserve
        all rows that appear in the Actuals table.
    """
    values: dict[str, list] = {}
    wb = ws.parent
    plan_ws = wb[SHEET_PLAN] if SHEET_PLAN in wb.sheetnames else None
    config_ws = wb[SHEET_CONFIG] if SHEET_CONFIG in wb.sheetnames else None

    # Discover WP data start by scanning for "WP ID" header row
    try:
        wp_header_row = _find_header_row(ws, SECTION_HEADER_LABELS["wp_id_column"])
    except ValueError:
        raise ValueError(
            "Cannot find 'WP ID' header row in Actuals sheet. "
            "Was the sheet structure manually edited?"
        )
    first_wp_row = wp_header_row + 1
    r = first_wp_row
    while r <= ws.max_row:
        col_a = ws.cell(r, 1).value
        col_b = ws.cell(r, 2).value

        # Stop at TOTAL row or empty
        if col_b and str(col_b).upper() == "TOTAL":
            break
        if col_a is None:
            break

        # Resolve WP ID (may be a =Plan!A{n} or =Config!A{n} formula)
        wp_id = _resolve_plan_ref(col_a, plan_ws, config_ws)
        if wp_id is None:
            r += 1
            continue

        monthly = _read_monthly_values(ws, r, month_col_start, project_months)

        # Preserve every Actuals WP row — a row with all-None monthlies
        # still documents user intent (explicit zero / deliberate clear).
        values[wp_id] = monthly

        r += 1

    return values


def read_actuals_hours_sheet(
    ws,
    project_months: int | None = None,
    month_col_start: int = MONTH_COL_START,
) -> dict[str, dict[str, list[float | None]]]:
    """Read Actuals-Hours blue-input hours for round-trip preservation.

    Anchors on the "WP ID" month-header row, then walks person groups. Each
    person section starts with a merged A:B header whose col A holds the
    literal person name (written as a literal, not a formula, so the
    forecast-EUR VLOOKUP works under ``data_only``). Under each person, WP
    rows carry ``=Config!A{n}`` in col A; external rows carry a literal
    project name. TOTAL and EU TOTAL rows are skipped.

    Columns with ``outline_level >= 1`` belong to the collapsible Forecast
    EUR band — those are derived formulas, never user input, so they are
    skipped when reading monthly values. Tolerates internal blank rows
    inside a person section (per the rebuild shape-drift design).

    Args:
        ws: Actuals-Hours worksheet.
        project_months: Number of project months. If None, falls back to
            ``config.PROJECT_MONTHS`` for backward compatibility.
        month_col_start: First month column (default 3 = col C).

    Returns:
        ``{person_name: {wp_id_or_external_project: [h_m1, h_m2, ...]}}``.
        Each monthly entry is a float (user input) or None (blank cell /
        greyed inactive cell / formula).
    """
    if project_months is None:
        from budget_planner.config import PROJECT_MONTHS

        project_months = PROJECT_MONTHS

    wb = ws.parent
    config_ws = wb[SHEET_CONFIG] if SHEET_CONFIG in wb.sheetnames else None

    # Anchor: the month-header row ("WP ID" in col A). Person sections
    # begin strictly below it.
    month_header_row = _find_header_row(ws, SECTION_HEADER_LABELS["wp_id_column"])

    # Pre-compute the set of hour columns (skip outline-level >= 1 columns
    # — those are the forecast-EUR band mirror, recomputed on rebuild).
    hour_cols: list[int] = []
    for mi in range(project_months):
        col = month_col_start + mi
        if ws.column_dimensions[get_column_letter(col)].outline_level >= 1:
            continue
        hour_cols.append(col)

    # ── Find person header rows ──
    # Convention (actuals_hours.py): person headers are A:B merged with a
    # literal name in col A (not a formula) + SECTION_FILL banner. Locate
    # them by matching the A:B merge anchored at col 1 — no reliance on
    # font colour, which openpyxl doesn't always round-trip cleanly.
    merged_a_b_rows: set[int] = set()
    for mr in ws.merged_cells.ranges:
        if mr.min_col == 1 and mr.max_col >= 2 and mr.min_row == mr.max_row:
            merged_a_b_rows.add(mr.min_row)

    person_headers: list[tuple[int, str]] = []
    for r in sorted(merged_a_b_rows):
        if r <= month_header_row:
            continue
        val = ws.cell(r, 1).value
        if val is None:
            continue
        if isinstance(val, str) and val.startswith("="):
            # Defensive: newer builder writes literal names, but tolerate
            # older formula-headers by resolving through Config.
            m = re.match(rf"={SHEET_CONFIG}!A(\d+)", val)
            if m and config_ws is not None:
                resolved = config_ws.cell(int(m.group(1)), 1).value
                if resolved:
                    person_headers.append((r, str(resolved).strip()))
            continue
        name = str(val).strip()
        if not name:
            continue
        # Skip merged section-header rows like "FORECAST EUR (from hours)"
        # that might also span A:B in future layouts. Person headers use
        # SECTION_FILL; the forecast-band label row uses LIGHT_GREEN_FILL.
        # Simple heuristic: section banner is on row 1; skip it.
        if r == 1:
            continue
        person_headers.append((r, name))

    def _read_hours_row(row: int) -> list[float | None]:
        """Read hour cells for one data row, skipping outline-level columns."""
        out: list[float | None] = []
        col_idx = 0
        for mi in range(project_months):
            col = month_col_start + mi
            if col_idx < len(hour_cols) and hour_cols[col_idx] == col:
                val = ws.cell(row, col).value
                if val is None or val == "":
                    out.append(None)
                elif isinstance(val, str):
                    if val.startswith("="):
                        out.append(None)
                    else:
                        try:
                            out.append(float(val))
                        except ValueError:
                            out.append(None)
                else:
                    try:
                        out.append(float(val))
                    except (TypeError, ValueError):
                        out.append(None)
                col_idx += 1
            else:
                # Column was suppressed by outline filter — reader would
                # have no hour-input to read here by design. Emit None to
                # keep the per-month list length stable.
                out.append(None)
        return out

    def _any_nonnull(series: list[float | None]) -> bool:
        """True if the row has any logged value (including 0.0). Used to
        prefer a later duplicate-key row that actually has data over an
        earlier all-blank one — a user logging 0 is still "present"."""
        return any(v is not None for v in series)

    hours: dict[str, dict[str, list[float | None]]] = {}

    for hi, (header_row, name) in enumerate(person_headers):
        if hi + 1 < len(person_headers):
            group_end = person_headers[hi + 1][0] - 1
        else:
            group_end = ws.max_row

        per_wp: dict[str, list[float | None]] = {}

        for dr in range(header_row + 1, group_end + 1):
            col_b = ws.cell(dr, 2).value
            # Stop-or-skip on TOTAL / EU TOTAL rows.
            if col_b is not None and str(col_b).strip().upper() in (
                "TOTAL",
                "EU TOTAL",
            ):
                continue

            col_a = ws.cell(dr, 1).value
            if col_a is None or (isinstance(col_a, str) and not col_a.strip()):
                # Internal blank row — tolerate and keep scanning.
                continue

            key: str | None = None
            if isinstance(col_a, str) and col_a.startswith("="):
                # WP row: =Config!A{n}
                m = re.match(rf"={SHEET_CONFIG}!A(\d+)", col_a)
                if m and config_ws is not None:
                    resolved = config_ws.cell(int(m.group(1)), 1).value
                    if resolved:
                        key = str(resolved).strip()
            else:
                # External project row: literal project name.
                key = str(col_a).strip()

            if not key:
                continue

            series = _read_hours_row(dr)
            # Preserve every identifiable row (including all-None) so the
            # builder sees an explicit "present but untouched" signal —
            # matches the ``read_actuals_values`` preservation contract.
            # First-wins if duplicate keys appear in one section.
            if key not in per_wp or _any_nonnull(series):
                per_wp[key] = series

        if per_wp:
            hours[name] = per_wp

    return hours


def _resolve_plan_ref(cell_value, plan_ws, config_ws=None):
    """Resolve a cell value that might be a =Plan!A{n} or =Config!A{n} formula."""
    if cell_value is None:
        return None
    val = str(cell_value)
    # Try =Plan!A{n} — resolve by reading the Plan cell (which itself may be =Config!A{n})
    m = re.match(rf"={SHEET_PLAN}!A(\d+)", val)
    if m and plan_ws:
        plan_row = int(m.group(1))
        resolved = plan_ws.cell(plan_row, 1).value
        if resolved and isinstance(resolved, str) and resolved.startswith("="):
            # Follow the chain: Plan cell is a formula like =Config!A8
            m2 = re.match(rf"={SHEET_CONFIG}!A(\d+)", resolved)
            if m2 and config_ws:
                config_row = int(m2.group(1))
                final = config_ws.cell(config_row, 1).value
                return str(final) if final else None
        return str(resolved) if resolved else None
    # Try =Config!A{n}
    m3 = re.match(rf"={SHEET_CONFIG}!A(\d+)", val)
    if m3 and config_ws:
        config_row = int(m3.group(1))
        resolved = config_ws.cell(config_row, 1).value
        return str(resolved) if resolved else None
    if not val.startswith("="):
        return val
    return None
