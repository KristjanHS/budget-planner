"""Read existing workbooks for round-trip updates."""
import re

import openpyxl


def detect_version(wb) -> int:
    """Return 2 if both Team and Capacity sheets exist, else 1."""
    names = wb.sheetnames
    if "Team" in names and "Capacity" in names:
        return 2
    return 1


def read_plan_config(ws, layout) -> list[dict]:
    """Read WP definitions from Plan config table."""
    wps = []
    for row in range(layout["config_first_wp_row"], layout["config_last_wp_row"] + 1):
        wp_id = ws.cell(row, 1).value
        wp_name = ws.cell(row, 2).value
        fte = ws.cell(row, 3).value
        start = ws.cell(row, 4).value
        end = ws.cell(row, 5).value
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


def read_team_sheet(ws) -> dict:
    """Read Team roster and WP share matrix.

    Returns:
        {
            "roster": [{"name": "Alice", "max_fte": 1.0, "row": 3}, ...],
            "shares": {"Alice": {"WP1": 0.25, "WP2": 0.50}, ...},
        }
    """
    # Find roster header: "Person" in column A
    roster_header_row = None
    for r in range(1, ws.max_row + 1):
        if ws.cell(r, 1).value == "Person":
            roster_header_row = r
            break
    if roster_header_row is None:
        raise ValueError("Cannot find roster header ('Person' in column A)")

    # Read roster rows until empty name
    roster = []
    r = roster_header_row + 1
    while r <= ws.max_row:
        name = ws.cell(r, 1).value
        if name is None or str(name).strip() == "":
            break
        max_fte = ws.cell(r, 2).value
        roster.append({
            "name": str(name),
            "max_fte": float(max_fte or 0),
            "row": r,
        })
        r += 1

    # Find WP SHARE SEEDS section header
    share_section_row = None
    for r in range(roster_header_row + 1, ws.max_row + 1):
        val = ws.cell(r, 1).value
        if val and "WP SHARE SEEDS" in str(val).upper():
            share_section_row = r
            break

    shares: dict[str, dict[str, float]] = {}
    if share_section_row is None:
        return {"roster": roster, "shares": shares}

    # Read share header row (row after section header)
    share_header_row = share_section_row + 1
    wp_columns: dict[str, int] = {}
    for c in range(2, ws.max_column + 1):
        val = ws.cell(share_header_row, c).value
        if val is None:
            break
        wp_columns[str(val)] = c

    # Read share data rows — one per roster person
    r = share_header_row + 1
    for person in roster:
        person_name = person["name"]
        # Col A may be a formula ref like =A3 — resolve via roster name
        person_shares: dict[str, float] = {}
        for wp_id, col in wp_columns.items():
            val = ws.cell(r, col).value
            if val is not None and not (isinstance(val, str) and val.startswith("=")):
                try:
                    person_shares[wp_id] = float(val)
                except (ValueError, TypeError):
                    pass
        if person_shares:
            shares[person_name] = person_shares
        r += 1

    return {"roster": roster, "shares": shares}


def read_capacity_sheet(ws, config_wps) -> dict:
    """Read Capacity sheet person groups.

    Args:
        ws: Capacity worksheet
        config_wps: list of WP dicts from read_plan_config (used to classify rows)

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
    from budget_planner.config import MONTH_COL_START, PROJECT_MONTHS

    # Build WP ID lookup from config_wps: plan config row -> WP ID
    # config_first_wp_row is typically 6 (row 5 is header "WP ID")
    # We need to discover it from the Plan layout, but we receive config_wps
    # which are indexed 0..N-1.  The formula pattern is =Plan!A{row}.
    wp_ids = {wp["id"] for wp in config_wps}

    # ── Read external commitments config block ──
    external_commitments = []
    ext_config_header = None
    for r in range(1, min(ws.max_row + 1, 20)):
        val = ws.cell(r, 1).value
        if val and "EXTERNAL COMMITMENTS" in str(val).upper():
            ext_config_header = r
            break

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
            external_commitments.append({
                "person": str(person),
                "project": str(project or ""),
                "fte": float(fte or 0),
                "start": int(start or 1),
                "end": int(end or 1),
            })
            r += 1

    # ── Find month header row: "WP ID" in col A ──
    month_header_row = None
    for r in range(1, ws.max_row + 1):
        if ws.cell(r, 1).value == "WP ID":
            month_header_row = r
            break
    if month_header_row is None:
        raise ValueError("Cannot find month header row ('WP ID' in Capacity col A)")

    # ── Find person header rows ──
    # Person headers have bold font + SECTION_FILL (fgColor "2F5496")
    # and formula like =Team!A{n}
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

    # ── Parse each person group ──
    person_groups: dict[str, dict] = {}

    # Resolve person names from =Team!A{n} formulas via the Team sheet.
    wb = ws.parent
    team_ws = wb["Team"] if "Team" in wb.sheetnames else None

    # Cache Plan layout for WP ID resolution (avoid repeated discover_layout)
    plan_ws = wb["Plan"]
    plan_layout = discover_layout(plan_ws)
    config_first_wp_row = plan_layout["config_first_wp_row"]

    def _resolve_person_name(cell_value):
        """Resolve person name from formula or literal value."""
        if cell_value is None:
            return None
        val = str(cell_value)
        m = re.match(r"=Team!A(\d+)", val)
        if m and team_ws:
            team_row = int(m.group(1))
            name_val = team_ws.cell(team_row, 1).value
            return str(name_val) if name_val else None
        if not val.startswith("="):
            return val
        return None

    def _resolve_wp_id(cell_value):
        """Resolve WP ID from formula =Plan!A{n} or literal value."""
        if cell_value is None:
            return None
        val = str(cell_value)
        m = re.match(r"=Plan!A(\d+)", val)
        if m:
            plan_row = int(m.group(1))
            idx = plan_row - config_first_wp_row
            if 0 <= idx < len(config_wps):
                return config_wps[idx]["id"]
            return None
        if val in wp_ids:
            return val
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

            # Skip TOTAL and REMAINING rows
            if col_b and str(col_b).upper() in ("TOTAL", "REMAINING"):
                continue

            # Try to classify the row
            resolved_wp = _resolve_wp_id(col_a)
            if resolved_wp is not None:
                # Programme WP row
                monthly = _read_monthly_values(ws, dr, MONTH_COL_START,
                                               PROJECT_MONTHS)
                wp_values[resolved_wp] = monthly
            elif col_a is not None:
                # Check if it's an external project (literal string, not formula
                # pointing to Plan)
                col_a_str = str(col_a)
                if not col_a_str.startswith("="):
                    # External project row
                    monthly = _read_monthly_values(ws, dr, MONTH_COL_START,
                                                   PROJECT_MONTHS)
                    external_values[col_a_str] = monthly

        person_groups[name] = {
            "wp_values": wp_values,
            "external_values": external_values,
        }

    return {
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

    Heuristics:
    - Scan column A rows 1-20 for "WP ID" -> config header row
    - Scan for "PLANNED" / "ACTUAL" -> block headers
    - Count WP rows between headers
    """
    config_header_row = None
    for r in range(1, 21):
        val = ws.cell(r, 1).value
        if val == "WP ID":
            config_header_row = r
            break

    if config_header_row is None:
        raise ValueError("Cannot find config table header row (looking for 'WP ID' in column A)")

    # Count WP rows
    config_first_wp_row = config_header_row + 1
    config_last_wp_row = config_first_wp_row
    while True:
        next_val = ws.cell(config_last_wp_row + 1, 1).value
        next_name = str(ws.cell(config_last_wp_row + 1, 2).value or "")
        if not next_val or next_name.upper().startswith("TOTAL"):
            break
        config_last_wp_row += 1
    config_total_row = config_last_wp_row + 1

    # Find PLANNED and ACTUAL block headers
    planned_header_row = None
    actual_header_row = None
    for r in range(config_total_row + 1, ws.max_row + 1):
        val = str(ws.cell(r, 1).value or "")
        if "PLANNED" in val.upper() and planned_header_row is None:
            planned_header_row = r
        elif "ACTUAL" in val.upper():
            actual_header_row = r

    if planned_header_row is None:
        raise ValueError("Cannot find PLANNED FTE block header")
    if actual_header_row is None:
        raise ValueError("Cannot find ACTUAL FTE block header")

    num_wps = config_last_wp_row - config_first_wp_row + 1
    # Block structure: header row, year row, month row, then WP rows, total row
    planned_first_wp_row = planned_header_row + 3
    planned_last_wp_row = planned_first_wp_row + num_wps - 1
    planned_total_row = planned_last_wp_row + 1

    actual_first_wp_row = actual_header_row + 3
    actual_last_wp_row = actual_first_wp_row + num_wps - 1
    actual_total_row = actual_last_wp_row + 1

    return {
        "config_header_row": config_header_row,
        "config_first_wp_row": config_first_wp_row,
        "config_last_wp_row": config_last_wp_row,
        "config_total_row": config_total_row,
        "planned_header_row": planned_header_row,
        "planned_first_wp_row": planned_first_wp_row,
        "planned_last_wp_row": planned_last_wp_row,
        "planned_total_row": planned_total_row,
        "actual_header_row": actual_header_row,
        "actual_first_wp_row": actual_first_wp_row,
        "actual_last_wp_row": actual_last_wp_row,
        "actual_total_row": actual_total_row,
    }


def read_workbook(filepath: str) -> dict:
    """Read existing workbook, extract user-entered cell values.

    Returns:
        {
            "planned": {
                0: [0.2, 0.2, None, ...],   # 60 values per WP index
                1: [0.44, 0.44, ...],
            },
            "actual": {
                0: [0.15, 0.2, 0.0, ...],
                1: [0.5, 0.0, ...],
            },
        }

    None = inactive/empty cell. Values preserved exactly as entered.
    """
    wb = openpyxl.load_workbook(filepath, data_only=False)
    ws = wb["Plan"]
    layout = discover_layout(ws)

    from budget_planner.config import MONTH_COL_START, PROJECT_MONTHS

    result: dict = {"planned": {}, "actual": {}}

    num_wps = layout["config_last_wp_row"] - layout["config_first_wp_row"] + 1

    for block_key, first_row_key in [("planned", "planned_first_wp_row"), ("actual", "actual_first_wp_row")]:
        first_row = layout[first_row_key]
        for i in range(num_wps):
            row = first_row + i
            values = []
            for m in range(PROJECT_MONTHS):
                col = MONTH_COL_START + m
                cell = ws.cell(row, col)
                val = cell.value
                # Skip formula cells and empty/greyed cells
                if val is not None and isinstance(val, str) and val.startswith("="):
                    values.append(None)
                elif val is None or val == "":
                    values.append(None)
                else:
                    values.append(val)
            result[block_key][i] = values

    # v2 extensions
    version = detect_version(wb)
    result["version"] = version

    if version == 2:
        result["config_wps"] = read_plan_config(ws, layout)
        result["team"] = read_team_sheet(wb["Team"])
        result["capacity"] = read_capacity_sheet(wb["Capacity"],
                                                  result["config_wps"])

    wb.close()
    return result
