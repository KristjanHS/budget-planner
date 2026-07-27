"""Read the Actuals sheet — through-month override + per-WP EUR values."""

from budget_planner.layout import (
    MONTH_COL_START,
    SECTION_HEADER_LABELS,
    SHEET_CONFIG,
    SHEET_PLAN,
)

from .layout_discovery import (
    _find_header_row,
    _read_monthly_values,
    _resolve_plan_ref,
)

__all__ = [
    "read_actuals_override",
    "read_actuals_values",
]


def read_actuals_override(ws) -> int | None:
    """Read the user-typed through-month override from row 1, column B.

    Row-1 layout on Actuals / Actuals-Hours:
        A = label, B = blue-input override, C = black-formula auto-derived,
        D = effective (IF(ISNUMBER(B),B,C), target of the named range).

    Returns:
        ``int`` when the user typed a numeric override into B1; ``None`` when
        the cell is blank, a formula, or any non-numeric value — signalling
        the builder to fall back to the auto-derived formula in C1.
    """
    val = ws.cell(row=1, column=2).value
    if val is None:
        return None
    if isinstance(val, bool):
        return None
    if isinstance(val, (int, float)):
        return int(val)
    return None


def read_actuals_values(ws, project_months, month_col_start=MONTH_COL_START) -> dict[str, list]:
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
        raise ValueError("Cannot find 'WP ID' header row in Actuals sheet. Was the sheet structure manually edited?")
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
