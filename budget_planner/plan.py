"""Plan sheet builder — compact header + planned FTE block.

Two-phase build. Phase 1 creates structure with empty FTE cells. Phase 2
(``_apply_planned_formulas()``) fills cells with ``=SUM(Capacity!...)``
after Capacity is built. Has validation columns: Sum, Budget FTE, Variance.

Role: monthly cells are formula-driven (``=SUM(Capacity!...)``) — **do not
edit directly**.

## Conditional formatting

- Planned variance column: **yellow** when ``ABS(value) > 0.01``.
- Actual remaining column: **red** when ``value < 0``.
"""

from openpyxl.styles import Font, Protection
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.formatting.rule import FormulaRule

from budget_planner.layout import (
    ConfigParamRows,
    MONTH_COL_START, SHEET_PLAN, SHEET_CAPACITY, SHEET_CONFIG,
    SECTION_HEADER_LABELS,
    build_year_month_headers, cell_ref, month_range_ref,
    month_col_end as _month_col_end, val_columns,
    wp_active_months,
)
from budget_planner.styles import (
    BLACK_FORMULA, GREEN_XREF, HEADER_FONT,
    GREY_FILL, LIGHT_GREEN_FILL,
    CF_RED_FILL, style_cell, col_letter, write_section_header,
    sized_comment,
)


def _apply_planned_formulas(ws, layout, capacity_layout, *,
                            work_packages=None, project_params=None):
    """Fill planned FTE cells with cross-sheet SUM formulas from Capacity.

    Called after build_capacity_sheet() so that capacity row positions are known.
    Cells are already styled (BLACK_FORMULA, locked) by build_plan_sheet();
    this function only sets the formula values.

    Args:
        ws: The Plan worksheet.
        layout: Layout dict returned by build_plan_sheet().
        capacity_layout: Layout dict returned by build_capacity_sheet().
        work_packages: WP list; falls back to config.WORK_PACKAGES.
        project_params: Project params dict; falls back to config values.
    """
    if project_params is None or work_packages is None:
        from budget_planner import config
        if project_params is None:
            project_params = config.default_project_params()
        if work_packages is None:
            work_packages = config.WORK_PACKAGES

    # Derived constants
    month_col_start = MONTH_COL_START

    for i, wp in enumerate(work_packages):
        row = layout["planned_first_wp_row"] + i
        wp_id = wp["id"]

        # Find all Capacity rows for this WP across all persons
        cap_rows = []
        for person_name, group in capacity_layout["person_groups"].items():
            if wp_id in group["wp_rows"]:
                cap_rows.append(group["wp_rows"][wp_id])
        cap_rows.sort()  # deterministic order

        active = wp_active_months(wp_id, work_packages)
        for m in range(project_params["months"]):
            c = month_col_start + m
            month_num = m + 1

            if month_num in active and cap_rows:
                cl = col_letter(c)
                refs = ",".join(f"{SHEET_CAPACITY}!{cl}{cr}" for cr in cap_rows)
                ws.cell(row=row, column=c).value = f"=SUM({refs})"


def build_plan_sheet(ws, wb, *, project_params=None, work_packages=None,
                     config_layout):
    """Build the Plan sheet: compact header + planned FTE block.

    Args:
        ws: The Plan worksheet to populate.
        wb: The parent Workbook (needed for named ranges).
        project_params: dict with keys 'title', 'start', 'months', 'rate';
            falls back to config.py values when None.
        work_packages: list of WP dicts; falls back to config.WORK_PACKAGES.
        config_layout: dict from build_config_sheet() with wp_first_row,
            wp_last_row, param_rows, etc.

    Returns:
        Layout dict with row numbers for all key positions.
    """
    if project_params is None or work_packages is None:
        from budget_planner import config
        if project_params is None:
            project_params = config.default_project_params()
        if work_packages is None:
            work_packages = config.WORK_PACKAGES

    # Config sheet WP row positions (config_layout always provided)
    cfg_wp_first = config_layout["wp_first_row"]
    cfg_wp_last = config_layout["wp_last_row"]

    # Derived constants
    month_col_start = MONTH_COL_START
    num_wps = len(work_packages)
    month_col_end = _month_col_end(project_params["months"])
    _vcols = val_columns(project_params["months"])
    val_col_sum = _vcols["sum"]
    val_col_budget = _vcols["budget"]
    val_col_diff = _vcols["diff"]

    # -- ROW 1: COMPACT HEADER --
    pr: ConfigParamRows = config_layout["param_rows"]
    # A1: "Title:", B1: =Config!B{title_row}
    style_cell(ws, 1, 1, "Title:", font=HEADER_FONT, border=False)
    style_cell(ws, 1, 2, font=Font(name="Arial", bold=True, size=14, color="008000"),
               border=False).value = f"={SHEET_CONFIG}!B{pr['Project Title']}"
    # D1: "Start:", E1: =Config!B{start_row}
    style_cell(ws, 1, 4, "Start:", font=HEADER_FONT, border=False)
    style_cell(ws, 1, 5, font=GREEN_XREF, border=False,
               fmt="MMM YYYY").value = f"={SHEET_CONFIG}!B{pr['Project Start']}"
    # F1: "Ceiling:", G1: ceiling formula
    style_cell(ws, 1, 6, "Ceiling:", font=HEADER_FONT, border=False)
    style_cell(ws, 1, 7, font=GREEN_XREF, border=False,
               fmt="\u20ac#,##0").value = (
        f"=SUM({SHEET_CONFIG}!E{cfg_wp_first}:E{cfg_wp_last})"
    )

    # partner_ceiling named range -> Plan!$G$1
    wb.defined_names.add(DefinedName(
        "partner_ceiling", attr_text=cell_ref(SHEET_PLAN, 1, 7)
    ))
    ws.cell(row=1, column=7).comment = sized_comment(
        "Personnel costs only (direct). The 25% indirect flat rate "
        "(the programme regulation) is added on top when computing "
        "total eligible costs."
    )

    # Row 2: blank separator (no action needed)

    # -- PLANNED FTE BLOCK (starts at row 3) --
    r = 3

    # Section header
    write_section_header(ws, r, SECTION_HEADER_LABELS["planned_fte"], merge_end=val_col_diff)
    r += 1

    year_row, month_row, r = build_year_month_headers(
        ws, r, project_params, month_col_start,
        val_headers=[
            (val_col_sum, "Sum"),
            (val_col_budget, "Budget FTE"),
            (val_col_diff, "Variance"),
        ],
    )
    first_wp_row = r

    # WP rows
    project_months = project_params["months"]
    for i, wp in enumerate(work_packages):
        row = r + i
        config_row = cfg_wp_first + i
        # WP ID/Name cross-ref to Config sheet
        style_cell(ws, row, 1, font=GREEN_XREF).value = (
            f"={SHEET_CONFIG}!A{config_row}"
        )
        style_cell(ws, row, 2, font=GREEN_XREF).value = (
            f"={SHEET_CONFIG}!B{config_row}"
        )

        active = wp_active_months(wp["id"], work_packages)
        for m in range(project_months):
            c = month_col_start + m
            month_num = m + 1

            if month_num in active:
                cell = style_cell(ws, row, c, font=BLACK_FORMULA, fmt="0.00")
                cell.protection = Protection(locked=True)
            else:
                cell = style_cell(ws, row, c, "", font=BLACK_FORMULA,
                                  fill=GREY_FILL, fmt="0.00")
                cell.protection = Protection(locked=True)

        # Validation columns
        sum_range = month_range_ref(row, month_col_start, month_col_end)
        style_cell(ws, row, val_col_sum, font=BLACK_FORMULA,
                   fmt="0.00").value = f"=SUM({sum_range})"
        # Budget: FTE Budget from Config WP table col C
        style_cell(ws, row, val_col_budget, font=GREEN_XREF,
                   fmt="0.00").value = f"={SHEET_CONFIG}!C{config_row}"
        style_cell(ws, row, val_col_diff, font=BLACK_FORMULA,
                   fmt="0.00").value = (
            f"={col_letter(val_col_budget)}{row}"
            f"-{col_letter(val_col_sum)}{row}"
        )

    last_wp_row = first_wp_row + num_wps - 1
    r = last_wp_row + 1

    # Totals row
    total_row = r
    style_cell(ws, r, 1, "", fill=LIGHT_GREEN_FILL)
    style_cell(ws, r, 2, "MONTHLY TOTAL", font=HEADER_FONT,
               fill=LIGHT_GREEN_FILL)
    for m in range(project_months):
        c = month_col_start + m
        col_range = f"{col_letter(c)}{first_wp_row}:{col_letter(c)}{last_wp_row}"
        style_cell(ws, r, c, font=BLACK_FORMULA, fill=LIGHT_GREEN_FILL,
                   fmt="0.00").value = f"=SUM({col_range})"
    sum_range = (f"{col_letter(val_col_sum)}{first_wp_row}:"
                 f"{col_letter(val_col_sum)}{last_wp_row}")
    style_cell(ws, r, val_col_sum, font=BLACK_FORMULA, fill=LIGHT_GREEN_FILL,
               fmt="0.00").value = f"=SUM({sum_range})"
    bud_range = (f"{col_letter(val_col_budget)}{first_wp_row}:"
                 f"{col_letter(val_col_budget)}{last_wp_row}")
    style_cell(ws, r, val_col_budget, font=BLACK_FORMULA, fill=LIGHT_GREEN_FILL,
               fmt="0.00").value = f"=SUM({bud_range})"
    diff_range = (f"{col_letter(val_col_diff)}{first_wp_row}:"
                  f"{col_letter(val_col_diff)}{last_wp_row}")
    style_cell(ws, r, val_col_diff, font=BLACK_FORMULA, fill=LIGHT_GREEN_FILL,
               fmt="0.00").value = f"=SUM({diff_range})"

    # Conditional formatting: red fill on negative variance
    diff_col_l = col_letter(val_col_diff)
    cell_range = f"{diff_col_l}{first_wp_row}:{diff_col_l}{last_wp_row}"
    ws.conditional_formatting.add(cell_range, FormulaRule(
        formula=[f"{diff_col_l}{first_wp_row}<0"], fill=CF_RED_FILL,
    ))

    # Freeze panes: freeze WP ID/Name columns and header rows
    ws.freeze_panes = f"C{first_wp_row}"
    ws.column_dimensions["B"].width = 25

    return {
        "config_first_wp_row": cfg_wp_first,
        "config_last_wp_row": cfg_wp_last,
        "planned_header_row": 3,
        "planned_first_wp_row": first_wp_row,
        "planned_last_wp_row": last_wp_row,
        "planned_total_row": total_row,
    }
