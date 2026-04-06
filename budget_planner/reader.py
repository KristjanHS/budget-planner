"""Read existing workbooks for round-trip updates."""
import re
from typing import Literal, overload

from budget_planner.layout import MONTH_COL_START, SHEET_PLAN, SHEET_CONFIG


@overload
def _find_header_row(
    ws, pattern: str, col: int = ..., max_row: int | None = ...,
    required: Literal[True] = ...,
) -> int: ...


@overload
def _find_header_row(
    ws, pattern: str, col: int = ..., max_row: int | None = ...,
    required: Literal[False] = ...,
) -> int | None: ...


def _find_header_row(ws, pattern, col=1, max_row=None, required=True):
    """Find row where cell contains pattern (case-insensitive substring).

    Args:
        ws: openpyxl worksheet.
        pattern: String to search for (case-insensitive substring match).
        col: Column number to scan (default 1 = col A).
        max_row: Upper row limit. If None, scans the whole sheet.
        required: If True and pattern not found, raises ValueError.

    Returns:
        Row number (int) or None if not found and required=False.

    Raises:
        ValueError: If required=True and pattern is not found.
    """
    limit = max_row if max_row is not None else ws.max_row
    for r in range(1, limit + 1):
        val = ws.cell(r, col).value
        if val and pattern.upper() in str(val).upper():
            return r
    if required:
        raise ValueError(
            f"Cannot find '{pattern}' in column {col}, rows 1-{limit}"
        )
    return None


def read_config_sheet(ws) -> tuple[dict, list[dict]]:
    """Read Config sheet -> (project_params, work_packages).

    Returns:
        Tuple of (project_params, work_packages) where:
        - project_params: dict with keys 'title', 'start', 'months', 'rate',
          'eu_ceiling_pm'
        - work_packages: list of dicts with keys 'id', 'name', 'fte',
          'start', 'end'
    """
    # Scan column A for known param labels (position-independent)
    _param_map = {
        "Project Title": ("title", None),
        "Project Start": ("start", None),
        "Project Months": ("months", lambda v: int(v or 60)),
        "FTE Rate": ("rate", lambda v: float(v or 7500) if not isinstance(v, str) or not v.startswith("=") else v),
        "EU Annual Ceiling": ("eu_ceiling_pm", lambda v: float(v or 12.0)),
    }
    project_params: dict = {}
    for r in range(1, ws.max_row + 1):
        label = str(ws.cell(r, 1).value or "")
        for prefix, (key, converter) in _param_map.items():
            if prefix.lower() in label.lower():
                raw = ws.cell(r, 2).value
                project_params[key] = converter(raw) if converter else raw
                break
        if len(project_params) == len(_param_map):
            break
    # Validate all params found
    for prefix, (key, _) in _param_map.items():
        if key not in project_params:
            raise ValueError(f"Config sheet missing param row matching '{prefix}'")

    # Work packages — find dynamically so adding param rows doesn't break
    wp_section = _find_header_row(ws, "WORK PACKAGES")
    work_packages: list[dict] = []
    r = wp_section + 2  # skip section header + column headers
    while r <= ws.max_row:
        wp_id = ws.cell(r, 1).value
        if wp_id is None or str(wp_id).strip() == "":
            break
        work_packages.append({
            "id": str(wp_id),
            "name": str(ws.cell(r, 2).value or ""),
            "fte": float(ws.cell(r, 3).value or 0),
            "start": int(ws.cell(r, 4).value or 1),
            "end": int(ws.cell(r, 5).value or 1),
            "type": str(ws.cell(r, 6).value or "flat"),
        })
        r += 1

    return project_params, work_packages


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
    for row in range(layout["config_first_wp_row"],
                     layout["config_last_wp_row"] + 1):
        wp_id = config_ws.cell(row, 1).value
        wp_name = config_ws.cell(row, 2).value
        fte = config_ws.cell(row, 3).value
        start = config_ws.cell(row, 4).value
        end = config_ws.cell(row, 5).value
        if wp_id is None:
            continue
        wps.append({
            "id": str(wp_id),
            "name": str(wp_name or ""),
            "fte": float(fte or 0),
            "start": int(start or 1),
            "end": int(end or 1),
        })
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
    section_row = _find_header_row(ws, "RATE BY GRADE", required=False)
    if section_row is None:
        from budget_planner.config import GRADE_RATES
        return dict(GRADE_RATES)

    # Skip section header + column headers
    r = section_row + 2
    rates: dict[str, int | float] = {}
    while r <= ws.max_row:
        grade_name = ws.cell(r, 1).value
        if grade_name is None or str(grade_name).strip() == "":
            break
        rate_val = ws.cell(r, 2).value
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
    roster_section = _find_header_row(ws, "TEAM ROSTER", required=True)

    # Person header is the next row after the section header
    person_header = _find_header_row(
        ws, "Person", col=1, max_row=roster_section + 5,
    )

    # Detect whether Grade column exists (col C header)
    has_grade_col = False
    grade_hdr = ws.cell(person_header, 3).value
    if grade_hdr and "grade" in str(grade_hdr).lower():
        has_grade_col = True

    # Read roster rows until empty name or non-numeric max_fte (next section)
    roster: list[dict] = []
    r = person_header + 1
    while r <= ws.max_row:
        name = ws.cell(r, 1).value
        if name is None or str(name).strip() == "":
            break
        max_fte = ws.cell(r, 2).value
        # Roster rows have numeric max_fte; a string means we hit a header row
        if max_fte is not None and not isinstance(max_fte, (int, float)):
            break
        grade = DEFAULT_GRADE
        if has_grade_col:
            grade_val = ws.cell(r, 3).value
            if grade_val is not None and str(grade_val).strip():
                grade = str(grade_val).strip()
        roster.append({
            "name": str(name).strip(),
            "max_fte": float(max_fte or 0),
            "grade": grade,
        })
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
    rp_section = _find_header_row(ws, "REPORTING PERIODS", required=False)
    if rp_section is None:
        if fallback_months is not None:
            from budget_planner.config import default_reporting_periods
            return default_reporting_periods(fallback_months)
        return []

    # Skip section header + column headers
    r = rp_section + 2
    periods: list[dict] = []
    while r <= ws.max_row:
        period_id = ws.cell(r, 1).value
        if period_id is None or str(period_id).strip() == "":
            break
        start = ws.cell(r, 2).value
        end = ws.cell(r, 3).value
        periods.append({
            "id": str(period_id),
            "start": int(start or 1),
            "end": int(end or 1),
        })
        r += 1

    return periods


def read_capacity_assignments(ws) -> list[dict]:
    """Read Person-WP Assignments config block from Capacity sheet.

    Returns:
        List of dicts with keys: person, wp, total_fte, start, end.
        Empty list if the block is not found.
    """
    assignments_header = _find_header_row(
        ws, "PERSON-WP ASSIGNMENTS", required=False
    )
    if assignments_header is None:
        return []

    # Skip section header + column headers
    r = assignments_header + 2
    assignments: list[dict] = []
    while r <= ws.max_row:
        person = ws.cell(r, 1).value
        if person is None or str(person).strip() == "":
            break
        wp = ws.cell(r, 2).value
        total_fte = ws.cell(r, 3).value
        start = ws.cell(r, 4).value
        end = ws.cell(r, 5).value
        assignments.append({
            "person": str(person),
            "wp": str(wp or ""),
            "total_fte": float(total_fte or 0),
            "start": int(start or 1),
            "end": int(end or 1),
        })
        r += 1

    return assignments


def read_capacity_sheet(ws, config_wps, project_months=None,
                        month_col_start=MONTH_COL_START) -> dict:
    """Read Capacity sheet person groups.

    Args:
        ws: Capacity worksheet
        config_wps: list of WP dicts from read_plan_config (used to classify rows)
        project_months: Number of project months. If None, falls back to
            config.PROJECT_MONTHS for backward compatibility.
        month_col_start: First month column (default 3 = col C).

    Returns:
        {
            "external_commitments": [{"person": ..., "project": ..., "fte": ...,
                                      "start": ..., "end": ...}],
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

    # Build WP ID lookup from config_wps
    wp_ids = {wp["id"] for wp in config_wps}

    # ── Read person-WP assignments config block ──
    assignments = read_capacity_assignments(ws)

    # ── Read external commitments config block ──
    external_commitments = []
    ext_config_header = _find_header_row(
        ws, "EXTERNAL COMMITMENTS", required=False
    )

    if ext_config_header is not None:
        # Skip header row (Person, Project, Total FTE, Start Month, End Month)
        r = ext_config_header + 2  # skip section header + column headers
        while r <= ws.max_row:
            person = ws.cell(r, 1).value
            project = ws.cell(r, 2).value
            if person is None or str(person).strip() == "":
                break
            fte = ws.cell(r, 3).value
            start = ws.cell(r, 4).value
            end = ws.cell(r, 5).value
            is_eu_val = ws.cell(r, 6).value
            external_commitments.append({
                "person": str(person),
                "project": str(project or ""),
                "fte": float(fte or 0),
                "start": int(start or 1),
                "end": int(end or 1),
                "is_eu": str(is_eu_val).strip().upper() in ("YES", "TRUE", "1"),
            })
            r += 1

    # ── Find month header row: "WP ID" in col A ──
    month_header_row = _find_header_row(ws, "WP ID")

    # ── Find person header rows ──
    # Person headers have bold font + SECTION_FILL (fgColor "2F5496")
    # and formula like =Config!A{n}
    person_headers = []
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
            if (ws.cell(r, 2).value
                    and str(ws.cell(r, 2).value).upper() == "TOTAL"):
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

    # ── Parse each person group ──
    person_groups: dict[str, dict] = {}

    # Resolve person names from =Config!A{n} formulas via the Config sheet.
    wb = ws.parent
    config_ws = wb[SHEET_CONFIG] if SHEET_CONFIG in wb.sheetnames else None

    # Get Plan sheet for WP ID resolution (Capacity WP rows reference Plan)
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
                monthly = _read_monthly_values(ws, dr, month_col_start,
                                               project_months)
                wp_values[resolved_wp] = monthly
            elif col_a is not None:
                # Check if it's an external project (literal string, not formula
                # pointing to Plan)
                col_a_str = str(col_a)
                if not col_a_str.startswith("="):
                    # External project row
                    monthly = _read_monthly_values(ws, dr, month_col_start,
                                                   project_months)
                    external_values[col_a_str] = monthly

        person_groups[name] = {
            "wp_values": wp_values,
            "external_values": external_values,
        }

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
    planned_header_row = _find_header_row(ws, "PLANNED", required=True)

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
        config_ws = ws.parent[SHEET_CONFIG] if SHEET_CONFIG in ws.parent.sheetnames else None
        if config_ws:
            wp_section = _find_header_row(config_ws, "WORK PACKAGES", required=False)
            cfg_wp_first = (wp_section + 2) if wp_section else 8
        else:
            cfg_wp_first = 8
    cfg_wp_last = cfg_wp_first + num_wps - 1

    return {
        "config_first_wp_row": cfg_wp_first,
        "config_last_wp_row": cfg_wp_last,
        "planned_header_row": planned_header_row,
        "planned_first_wp_row": planned_first_wp_row,
        "planned_last_wp_row": planned_last_wp_row,
        "planned_total_row": planned_total_row,
    }


def read_actuals_values(ws, project_months, month_col_start=MONTH_COL_START) -> dict[str, list]:
    """Read Actuals EUR values for round-trip preservation.

    Args:
        ws: Actuals worksheet.
        project_months: Number of project months.
        month_col_start: First month column (default 3 = col C).

    Returns:
        Dict keyed by WP ID string -> list of monthly values (float or None).
        Only includes rows with at least one non-None value.
    """
    values: dict[str, list] = {}
    wb = ws.parent
    plan_ws = wb[SHEET_PLAN] if SHEET_PLAN in wb.sheetnames else None
    config_ws = wb[SHEET_CONFIG] if SHEET_CONFIG in wb.sheetnames else None

    # Discover WP data start by scanning for "WP ID" header row
    try:
        wp_header_row = _find_header_row(ws, "WP ID")
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

        # Only include if there's at least one non-None value
        if any(v is not None for v in monthly):
            values[wp_id] = monthly

        r += 1

    return values


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
