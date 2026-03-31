"""Read existing workbooks for round-trip updates."""
import openpyxl


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

    result = {"planned": {}, "actual": {}}

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
                elif val is None or val == "" or val == "":
                    values.append(None)
                else:
                    values.append(val)
            result[block_key][i] = values

    wb.close()
    return result
