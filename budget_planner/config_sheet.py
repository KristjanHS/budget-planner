"""Config sheet builder — project parameters, work packages, and team roster."""

from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import FormulaRule
from openpyxl.utils import get_column_letter

from budget_planner.layout import (
    SHEET_CONFIG,
    SECTION_HEADER_LABELS,
    COL_WP_FTE, COL_WP_RATE, COL_WP_CEILING,
    COL_WP_ASSIGNED_PM, COL_WP_VARIANCE, WP_BUDGET_LAST_COL,
    COL_SCHED_START, COL_SCHED_END, COL_SCHED_TYPE, SCHED_LAST_COL,
    COL_ROSTER_FTE, COL_ROSTER_GRADE, ROSTER_LAST_COL,
    COL_GRADE_RATE, GRADE_LAST_COL,
    COL_RP_START, COL_RP_END, RP_LAST_COL,
    COL_ASSIGN_PERSON, COL_ASSIGN_WP, COL_ASSIGN_FTE,
    COL_ASSIGN_START, COL_ASSIGN_END,
    COL_ASSIGN_SIG_ASSIGNED, COL_ASSIGN_SIG_EXTERNAL, COL_ASSIGN_BALANCE,
    ASSIGN_LAST_COL,
    COL_EXT_PERSON, COL_EXT_PROJECT, COL_EXT_FTE,
    COL_EXT_START, COL_EXT_END, COL_EXT_EU, EXT_LAST_COL,
)
from budget_planner.styles import (
    BLACK_FORMULA,
    BLUE_INPUT,
    HEADER_FONT,
    SECTION_FONT,
    SECTION_FILL,
    TITLE_FONT,
    LIGHT_BLUE_FILL,
    auto_fit_columns,
    CF_YELLOW_FILL,
    CF_RED_FILL,
    CENTER,
    style_cell,
)


def _write_budget_section(ws, start_row, work_packages, project_params):
    """Write the WP BUDGET section (section header, column headers, data rows, total, Excel Table).

    Returns:
        (next_row, layout_dict) where next_row is the first row after the total row,
        and layout_dict has keys: wp_header_row, wp_first_row, wp_last_row, wp_total_row.
    """
    budget_section_row = start_row
    budget_header_row = budget_section_row + 1
    budget_first_data_row = budget_header_row + 1

    # Section header — merged A:E
    ws.merge_cells(start_row=budget_section_row, start_column=1,
                   end_row=budget_section_row, end_column=WP_BUDGET_LAST_COL)
    style_cell(ws, budget_section_row, 1, value=SECTION_HEADER_LABELS["wp_budget"],
               font=SECTION_FONT, fill=SECTION_FILL, alignment=CENTER, border=False)

    # Column headers
    budget_headers = [
        "WP ID", "WP Name", "FTE Budget", "Rate (€/PM)", "Ceiling (€)",
        "Assigned PM", "Variance",
    ]
    for col, hdr in enumerate(budget_headers, start=1):
        style_cell(ws, budget_header_row, col, value=hdr,
                   font=HEADER_FONT, fill=LIGHT_BLUE_FILL, alignment=CENTER)

    # WP Budget data rows
    default_rate = project_params.get("rate", 10000)
    if not isinstance(default_rate, (int, float)):
        default_rate = 10000  # fall back when rate is a formula
    budget_last_data_row = budget_first_data_row + len(work_packages) - 1
    for row_idx, wp in enumerate(work_packages, start=budget_first_data_row):
        # Col A: WP ID
        style_cell(ws, row_idx, 1, value=wp["id"], font=BLUE_INPUT, border=False)
        # Col B: WP Name
        style_cell(ws, row_idx, 2, value=wp["name"], font=BLUE_INPUT, border=False)
        # Col C: FTE Budget
        style_cell(ws, row_idx, COL_WP_FTE, value=wp["fte"], font=BLUE_INPUT,
                   fmt="0.00", border=False)
        # Col D: Rate (€/PM) — preserve user-authored formula when present,
        # else fall back to the numeric rate (or project default). Formulas are
        # styled BLACK_FORMULA per the font-color convention (CLAUDE.md Rule 3).
        rate_raw = wp.get("rate_raw")
        if isinstance(rate_raw, str) and rate_raw.startswith("="):
            style_cell(ws, row_idx, COL_WP_RATE, value=rate_raw,
                       font=BLACK_FORMULA, fmt="€#,##0", border=False)
        else:
            wp_rate = wp.get("rate", default_rate)
            if not isinstance(wp_rate, (int, float)):
                wp_rate = default_rate
            style_cell(ws, row_idx, COL_WP_RATE, value=wp_rate, font=BLUE_INPUT,
                       fmt="€#,##0", border=False)
        # Col E: Ceiling (€) — formula = FTE * Rate
        fte_cl = get_column_letter(COL_WP_FTE)
        rate_cl = get_column_letter(COL_WP_RATE)
        style_cell(ws, row_idx, COL_WP_CEILING, font=BLACK_FORMULA, fmt="€#,##0",
                   border=False).value = f"={fte_cl}{row_idx}*{rate_cl}{row_idx}"
        # Col F: Assigned PM — SUMIFS over assignment section
        style_cell(ws, row_idx, COL_WP_ASSIGNED_PM, font=BLACK_FORMULA, fmt="0.00",
                   border=False).value = f"=SUMIFS(assign_ftes,assign_wps,A{row_idx})"
        # Col G: Variance — FTE Budget minus Assigned PM
        fte_cl_g = get_column_letter(COL_WP_FTE)
        apm_cl = get_column_letter(COL_WP_ASSIGNED_PM)
        style_cell(ws, row_idx, COL_WP_VARIANCE, font=BLACK_FORMULA, fmt="0.00",
                   border=False).value = f"={fte_cl_g}{row_idx}-{apm_cl}{row_idx}"

    # CF: yellow fill on Variance column when ABS(value) > 0.01
    var_cl = get_column_letter(COL_WP_VARIANCE)
    for row_idx in range(budget_first_data_row, budget_last_data_row + 1):
        rule = FormulaRule(
            formula=[f"ABS({var_cl}{row_idx})>0.01"],
            fill=CF_YELLOW_FILL,
        )
        ws.conditional_formatting.add(f"{var_cl}{row_idx}", rule)

    # TOTAL row
    budget_total_row = budget_last_data_row + 1
    style_cell(ws, budget_total_row, 1, "", fill=LIGHT_BLUE_FILL)
    style_cell(ws, budget_total_row, 2, "TOTAL", font=HEADER_FONT, fill=LIGHT_BLUE_FILL)
    fte_cl = get_column_letter(COL_WP_FTE)
    style_cell(ws, budget_total_row, COL_WP_FTE, font=BLACK_FORMULA,
               fill=LIGHT_BLUE_FILL, fmt="0.00").value = (
        f"=SUM({fte_cl}{budget_first_data_row}:{fte_cl}{budget_last_data_row})"
    )
    style_cell(ws, budget_total_row, COL_WP_RATE, "", fill=LIGHT_BLUE_FILL)
    ceil_cl = get_column_letter(COL_WP_CEILING)
    style_cell(ws, budget_total_row, COL_WP_CEILING, font=BLACK_FORMULA,
               fill=LIGHT_BLUE_FILL, fmt="€#,##0").value = (
        f"=SUM({ceil_cl}{budget_first_data_row}:{ceil_cl}{budget_last_data_row})"
    )
    apm_cl = get_column_letter(COL_WP_ASSIGNED_PM)
    style_cell(ws, budget_total_row, COL_WP_ASSIGNED_PM, font=BLACK_FORMULA,
               fill=LIGHT_BLUE_FILL, fmt="0.00").value = (
        f"=SUM({apm_cl}{budget_first_data_row}:{apm_cl}{budget_last_data_row})"
    )
    var_cl = get_column_letter(COL_WP_VARIANCE)
    style_cell(ws, budget_total_row, COL_WP_VARIANCE, font=BLACK_FORMULA,
               fill=LIGHT_BLUE_FILL, fmt="0.00").value = (
        f"=SUM({var_cl}{budget_first_data_row}:{var_cl}{budget_last_data_row})"
    )

    # Excel Table over header + data range (NOT including total row)
    table_ref = f"A{budget_header_row}:{get_column_letter(WP_BUDGET_LAST_COL)}{budget_last_data_row}"
    table_style = TableStyleInfo(
        name="TableStyleLight1", showFirstColumn=False,
        showLastColumn=False, showRowStripes=False, showColumnStripes=False,
    )
    table = Table(displayName="tbl_WorkPackages", ref=table_ref)
    table.tableStyleInfo = table_style
    ws.add_table(table)

    next_row = budget_total_row + 1
    layout = {
        "wp_header_row": budget_header_row,
        "wp_first_row": budget_first_data_row,
        "wp_last_row": budget_last_data_row,
        "wp_total_row": budget_total_row,
    }
    return next_row, layout


def _write_schedule_section(ws, start_row, wp_schedule):
    """Write the WP SCHEDULE section (section header, column headers, data rows, validation, CF).

    Returns:
        (next_row, layout_dict) where next_row is the first row after the last data row,
        and layout_dict has keys: sched_header_row, sched_first_row, sched_last_row.
    """
    sched_section_row = start_row
    sched_header_row = sched_section_row + 1
    sched_first_data_row = sched_header_row + 1

    # Section header — merged A:D
    ws.merge_cells(start_row=sched_section_row, start_column=1,
                   end_row=sched_section_row, end_column=SCHED_LAST_COL)
    style_cell(ws, sched_section_row, 1, value=SECTION_HEADER_LABELS["wp_schedule"],
               font=SECTION_FONT, fill=SECTION_FILL, alignment=CENTER, border=False)

    # Column headers
    sched_headers = ["WP ID", "Start", "End", "Type"]
    for col, hdr in enumerate(sched_headers, start=1):
        style_cell(ws, sched_header_row, col, value=hdr,
                   font=HEADER_FONT, fill=LIGHT_BLUE_FILL, alignment=CENTER)

    # Schedule data rows
    sched_last_data_row = sched_first_data_row + len(wp_schedule) - 1
    for row_idx, entry in enumerate(wp_schedule, start=sched_first_data_row):
        style_cell(ws, row_idx, 1, value=entry["id"], font=BLUE_INPUT, border=False)
        style_cell(ws, row_idx, COL_SCHED_START, value=entry["start"],
                   font=BLUE_INPUT, fmt="0", border=False)
        style_cell(ws, row_idx, COL_SCHED_END, value=entry["end"],
                   font=BLUE_INPUT, fmt="0", border=False)
        style_cell(ws, row_idx, COL_SCHED_TYPE, value=entry["type"],
                   font=BLUE_INPUT, border=False)

    # Data validation dropdown on Type column
    dv_type = DataValidation(
        type="list",
        formula1='"flat,front-loaded,bell,back-loaded"',
        allow_blank=False,
    )
    dv_type.error = "Please select a valid distribution type"
    dv_type.errorTitle = "Invalid Type"
    type_col = get_column_letter(COL_SCHED_TYPE)
    dv_type.sqref = f"{type_col}{sched_first_data_row}:{type_col}{sched_last_data_row}"
    ws.add_data_validation(dv_type)

    # Conditional formatting: yellow fill if Start > End (per row)
    for row_idx in range(sched_first_data_row, sched_last_data_row + 1):
        rule = FormulaRule(
            formula=[f"$B{row_idx}>$C{row_idx}"],
            fill=CF_YELLOW_FILL,
        )
        ws.conditional_formatting.add(f"A{row_idx}:D{row_idx}", rule)

    next_row = sched_last_data_row + 1
    layout = {
        "sched_header_row": sched_header_row,
        "sched_first_row": sched_first_data_row,
        "sched_last_row": sched_last_data_row,
    }
    return next_row, layout


def _write_roster_section(ws, start_row, team):
    """Write the TEAM ROSTER section (section header, column headers, data rows).

    Returns:
        (next_row, layout_dict) where next_row is the first row after the last data row,
        and layout_dict has keys: roster_header_row, roster_first_row, roster_last_row, person_rows.
    """
    r = start_row

    # Section header — merged A:C
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=ROSTER_LAST_COL)
    style_cell(ws, r, 1, value=SECTION_HEADER_LABELS["team_roster"],
               font=SECTION_FONT, fill=SECTION_FILL, alignment=CENTER, border=False)
    style_cell(ws, r, 2, fill=SECTION_FILL)
    style_cell(ws, r, COL_ROSTER_GRADE, fill=SECTION_FILL)
    r += 1

    # Column headers
    roster_header_row = r
    style_cell(ws, r, 1, "Person", font=HEADER_FONT, fill=LIGHT_BLUE_FILL,
               alignment=CENTER)
    style_cell(ws, r, COL_ROSTER_FTE, "Max FTE", font=HEADER_FONT,
               fill=LIGHT_BLUE_FILL, alignment=CENTER)
    style_cell(ws, r, COL_ROSTER_GRADE, "Grade", font=HEADER_FONT,
               fill=LIGHT_BLUE_FILL, alignment=CENTER)
    r += 1

    # Team member data rows
    roster_first_row = r
    person_rows: dict[str, int] = {}
    for member in team:
        style_cell(ws, r, 1, member["name"], font=BLUE_INPUT)
        style_cell(ws, r, COL_ROSTER_FTE, member["max_fte"], font=BLUE_INPUT,
                   fmt="0.00")
        style_cell(ws, r, COL_ROSTER_GRADE, member.get("grade", "Mid"),
                   font=BLUE_INPUT)
        person_rows[member["name"]] = r
        r += 1
    roster_last_row = r - 1

    next_row = r
    layout = {
        "roster_header_row": roster_header_row,
        "roster_first_row": roster_first_row,
        "roster_last_row": roster_last_row,
        "person_rows": person_rows,
    }
    return next_row, layout


def _write_reporting_section(ws, start_row, reporting_periods):
    """Write the REPORTING PERIODS section (section header, column headers, data rows).

    Returns:
        (next_row, layout_dict) where next_row is the first row after the last data row,
        and layout_dict has keys: rp_first_row, rp_last_row.
    """
    if reporting_periods is None:
        from budget_planner.config import REPORTING_PERIODS
        reporting_periods = REPORTING_PERIODS

    r = start_row

    # Section header — merged A:C
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=RP_LAST_COL)
    style_cell(ws, r, 1, value=SECTION_HEADER_LABELS["reporting_periods"],
               font=SECTION_FONT, fill=SECTION_FILL, alignment=CENTER, border=False)
    for c in range(2, RP_LAST_COL + 1):
        style_cell(ws, r, c, fill=SECTION_FILL)
    r += 1

    # Column headers
    for col, hdr in enumerate(["Period", "Start Month", "End Month"], start=1):
        style_cell(ws, r, col, hdr, font=HEADER_FONT, fill=LIGHT_BLUE_FILL,
                   alignment=CENTER)
    r += 1

    # RP data rows
    rp_first_row = r
    for rp in reporting_periods:
        style_cell(ws, r, 1, rp["id"], font=BLUE_INPUT, border=False)
        style_cell(ws, r, COL_RP_START, rp["start"], font=BLUE_INPUT, fmt="0", border=False)
        style_cell(ws, r, COL_RP_END, rp["end"], font=BLUE_INPUT, fmt="0", border=False)
        r += 1
    rp_last_row = r - 1 if reporting_periods else rp_first_row

    next_row = r
    layout = {
        "rp_first_row": rp_first_row,
        "rp_last_row": rp_last_row,
    }
    return next_row, layout


def _write_rate_by_grade_section(ws, start_row, grade_rates):
    """Write the RATE BY GRADE section (section header, column headers, data rows).

    Returns:
        (next_row, layout_dict) where next_row is the first row after the last data row,
        and layout_dict has keys: grade_first_row, grade_last_row.
    """
    grade_section_row = start_row
    ws.merge_cells(start_row=grade_section_row, start_column=1,
                   end_row=grade_section_row, end_column=GRADE_LAST_COL)
    style_cell(ws, grade_section_row, 1, value=SECTION_HEADER_LABELS["rate_by_grade"],
               font=SECTION_FONT, fill=SECTION_FILL, alignment=CENTER, border=False)

    grade_header_row = grade_section_row + 1
    style_cell(ws, grade_header_row, 1, "Grade", font=HEADER_FONT,
               fill=LIGHT_BLUE_FILL, alignment=CENTER)
    style_cell(ws, grade_header_row, 2, "Rate (EUR/PM)", font=HEADER_FONT,
               fill=LIGHT_BLUE_FILL, alignment=CENTER)

    grade_first_row = grade_header_row + 1
    grade_names = list(grade_rates.keys())
    for gi, grade_name in enumerate(grade_names):
        gr = grade_first_row + gi
        style_cell(ws, gr, 1, grade_name, font=BLUE_INPUT, border=False)
        style_cell(ws, gr, COL_GRADE_RATE, grade_rates[grade_name],
                   font=BLUE_INPUT, fmt="€#,##0", border=False)
    grade_last_row = grade_first_row + len(grade_names) - 1

    next_row = grade_last_row + 1
    layout = {
        "grade_first_row": grade_first_row,
        "grade_last_row": grade_last_row,
    }
    return next_row, layout


def _write_assignments_section(ws, start_row, assignments):
    """Write the PERSON-WP ASSIGNMENTS section (section header, column headers, data rows, formulas).

    Returns:
        (next_row, layout_dict) where next_row is the first row after the last data row,
        and layout_dict has keys: assign_header_row, assign_first_row, assign_last_row,
        assign_config_rows.
    """
    r = start_row

    # Section header — merged A:H
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=ASSIGN_LAST_COL)
    style_cell(ws, r, 1, value=SECTION_HEADER_LABELS["assignments"],
               font=SECTION_FONT, fill=SECTION_FILL, alignment=CENTER, border=False)
    for c in range(2, ASSIGN_LAST_COL + 1):
        style_cell(ws, r, c, fill=SECTION_FILL)
    r += 1

    # Column headers
    assign_header_row = r
    assign_headers = [
        "Person", "WP", "Total FTE", "Start Month", "End Month",
        "Sig Assigned", "Sig External", "Balance",
    ]
    for col, hdr in enumerate(assign_headers, start=1):
        style_cell(ws, r, col, value=hdr,
                   font=HEADER_FONT, fill=LIGHT_BLUE_FILL, alignment=CENTER)
    r += 1

    # Data rows
    assign_first_row = r
    assign_config_rows: dict[tuple[str, str], int] = {}
    for entry in assignments:
        # Cols A-E: blue input
        style_cell(ws, r, COL_ASSIGN_PERSON, entry["person"], font=BLUE_INPUT, border=False)
        style_cell(ws, r, COL_ASSIGN_WP, entry["wp"], font=BLUE_INPUT, border=False)
        style_cell(ws, r, COL_ASSIGN_FTE, entry["total_fte"], font=BLUE_INPUT,
                   fmt="0.00", border=False)
        style_cell(ws, r, COL_ASSIGN_START, entry["start"], font=BLUE_INPUT,
                   fmt="0", border=False)
        style_cell(ws, r, COL_ASSIGN_END, entry["end"], font=BLUE_INPUT,
                   fmt="0", border=False)
        # Col F: Sig Assigned — SUMIFS over assign_ftes by person
        style_cell(ws, r, COL_ASSIGN_SIG_ASSIGNED, font=BLACK_FORMULA, fmt="0.00",
                   border=False).value = f"=SUMIFS(assign_ftes,assign_persons,A{r})"
        # Col G: Sig External — SUMIFS over ext_ftes by person
        style_cell(ws, r, COL_ASSIGN_SIG_EXTERNAL, font=BLACK_FORMULA, fmt="0.00",
                   border=False).value = f"=SUMIFS(ext_ftes,ext_persons,A{r})"
        # Col H: Balance — available PM minus assigned minus external
        bal_cl = get_column_letter(COL_ASSIGN_BALANCE)
        style_cell(ws, r, COL_ASSIGN_BALANCE, font=BLACK_FORMULA, fmt="0.00",
                   border=False).value = (
            f"=VLOOKUP(A{r},roster_lookup,2,FALSE)*project_months"
            f"-{get_column_letter(COL_ASSIGN_SIG_ASSIGNED)}{r}"
            f"-{get_column_letter(COL_ASSIGN_SIG_EXTERNAL)}{r}"
        )
        assign_config_rows[(entry["person"], entry["wp"])] = r
        r += 1
    assign_last_row = r - 1 if assignments else assign_first_row

    # CF on Balance column:
    #   red if < 0 (over-committed)
    #   yellow if < 10% of available PM (low remaining capacity)
    # Red is added first so it wins the overlap when Balance < 0.
    if assignments:
        bal_cl = get_column_letter(COL_ASSIGN_BALANCE)
        person_cl = get_column_letter(COL_ASSIGN_PERSON)
        for row_idx in range(assign_first_row, assign_last_row + 1):
            red_rule = FormulaRule(
                formula=[f"{bal_cl}{row_idx}<0"],
                fill=CF_RED_FILL,
            )
            ws.conditional_formatting.add(f"{bal_cl}{row_idx}", red_rule)
            yellow_rule = FormulaRule(
                formula=[
                    f"{bal_cl}{row_idx}<0.1*"
                    f"VLOOKUP({person_cl}{row_idx},roster_lookup,2,FALSE)"
                    f"*project_months"
                ],
                fill=CF_YELLOW_FILL,
            )
            ws.conditional_formatting.add(f"{bal_cl}{row_idx}", yellow_rule)

    next_row = r
    layout = {
        "assign_header_row": assign_header_row,
        "assign_first_row": assign_first_row,
        "assign_last_row": assign_last_row,
        "assign_config_rows": assign_config_rows,
    }
    return next_row, layout


def _write_external_section(ws, start_row, external_commitments):
    """Write the EXTERNAL COMMITMENTS section (section header, column headers, data rows).

    Returns:
        (next_row, layout_dict) where next_row is the first row after the last data row,
        and layout_dict has keys: ext_header_row, ext_first_row, ext_last_row,
        external_config_rows.
    """
    r = start_row

    # Section header — merged A:F
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=EXT_LAST_COL)
    style_cell(ws, r, 1, value=SECTION_HEADER_LABELS["externals"],
               font=SECTION_FONT, fill=SECTION_FILL, alignment=CENTER, border=False)
    for c in range(2, EXT_LAST_COL + 1):
        style_cell(ws, r, c, fill=SECTION_FILL)
    r += 1

    # Column headers
    ext_header_row = r
    ext_headers = ["Person", "Project", "Total FTE", "Start Month", "End Month", "EU Grant"]
    for col, hdr in enumerate(ext_headers, start=1):
        style_cell(ws, r, col, value=hdr,
                   font=HEADER_FONT, fill=LIGHT_BLUE_FILL, alignment=CENTER)
    r += 1

    # Data rows (all blue input)
    ext_first_row = r
    external_config_rows: dict[tuple[str, str], int] = {}
    for entry in external_commitments:
        style_cell(ws, r, COL_EXT_PERSON, entry["person"], font=BLUE_INPUT, border=False)
        style_cell(ws, r, COL_EXT_PROJECT, entry["project"], font=BLUE_INPUT, border=False)
        style_cell(ws, r, COL_EXT_FTE, entry["fte"], font=BLUE_INPUT, fmt="0.00",
                   border=False)
        style_cell(ws, r, COL_EXT_START, entry["start"], font=BLUE_INPUT, fmt="0",
                   border=False)
        style_cell(ws, r, COL_EXT_END, entry["end"], font=BLUE_INPUT, fmt="0",
                   border=False)
        style_cell(ws, r, COL_EXT_EU,
                   "Yes" if entry.get("is_eu", False) else "No",
                   font=BLUE_INPUT, border=False)
        external_config_rows[(entry["person"], entry["project"])] = r
        r += 1
    ext_last_row = r - 1 if external_commitments else ext_first_row

    next_row = r
    layout = {
        "ext_header_row": ext_header_row,
        "ext_first_row": ext_first_row,
        "ext_last_row": ext_last_row,
        "external_config_rows": external_config_rows,
    }
    return next_row, layout


def build_config_sheet(ws, wb, project_params, work_packages, wp_schedule=None,
                       team=None, reporting_periods=None, grade_rates=None,
                       assignments=None, external_commitments=None):
    """Build the Config sheet with project parameters, WP budget/schedule, team roster, and grade rates.

    Args:
        ws: Config worksheet to populate.
        wb: Parent Workbook (for named ranges).
        project_params: dict with keys 'title', 'start', 'months', 'rate'.
        work_packages: list of dicts with keys 'id', 'name', 'fte'.
        wp_schedule: list of dicts with keys 'id', 'start', 'end', 'type'.
            When None, lazy-imports from config.
        team: list of dicts with keys 'name', 'max_fte', 'grade'.
            When None, lazy-imports from config.
        reporting_periods: list of dicts with keys 'id', 'start', 'end'.
            When None, lazy-imports from config.
        grade_rates: dict mapping grade name to EUR/PM rate.
            When None, lazy-imports from config.
        assignments: list of dicts with keys 'person', 'wp', 'total_fte', 'start', 'end'.
            When None, defaults to empty list.
        external_commitments: list of dicts with keys 'person', 'project', 'fte',
            'start', 'end', 'is_eu'. When None, defaults to empty list.

    Returns:
        config_layout dict with keys including param_rows, wp_*, sched_*, roster_*,
        rp_*, grade_*, assign_*, ext_*.
    """
    if wp_schedule is None:
        from budget_planner import config
        wp_schedule = config.WP_SCHEDULE
    if team is None:
        from budget_planner import config
        team = config.TEAM
    if grade_rates is None:
        from budget_planner import config as _cfg
        grade_rates = _cfg.GRADE_RATES
    if assignments is None:
        assignments = []
    if external_commitments is None:
        external_commitments = []

    # ── Section 1: Project Parameters (rows 1-5) ──
    param_defs = [
        ("Project Title", project_params["title"], None, None),
        ("Project Start", project_params["start"], "MMM YYYY", None),
        ("Project Months", project_params["months"], "0", None),
        ("FTE Rate (EUR)", None, "€#,##0", None),  # value set later as formula
        ("EU Annual Ceiling (PM/FTE/year)", project_params.get("eu_ceiling_pm", 12.0), "0.00", None),
    ]
    param_rows: dict[str, int] = {}
    for i, (label, value, fmt, _) in enumerate(param_defs, start=1):
        # Label in column A — HEADER_FONT, no border
        style_cell(ws, i, 1, value=label, font=HEADER_FONT, border=False)
        # Value in column B — BLUE_INPUT, no border (except FTE Rate: set later)
        if label == "FTE Rate (EUR)":
            # FTE Rate is a formula — styled as BLACK_FORMULA, set after grade table
            style_cell(ws, i, 2, font=BLACK_FORMULA, fmt=fmt, border=False)
        else:
            font = TITLE_FONT if i == 1 else BLUE_INPUT
            style_cell(ws, i, 2, value=value, font=font, fmt=fmt, border=False)
        param_rows[label] = i

    # ── Section 2: WP Budget ──
    budget_section_row = len(param_defs) + 2  # gap row after last param
    next_row, budget_layout = _write_budget_section(ws, budget_section_row,
                                                    work_packages, project_params)

    # ── Section 2b: Person-WP Assignments ──
    next_row, assign_layout = _write_assignments_section(ws, next_row + 1, assignments)

    # ── Section 2c: External Commitments ──
    next_row, ext_layout = _write_external_section(ws, next_row + 1, external_commitments)

    # ── Section 3: WP Schedule ──
    next_row, sched_layout = _write_schedule_section(ws, next_row + 1, wp_schedule)

    # ── Section 4: Team Roster ──
    next_row, roster_layout = _write_roster_section(ws, next_row + 1, team)

    # ── Section 5: Reporting Periods ──
    next_row, rp_layout = _write_reporting_section(ws, next_row + 1, reporting_periods)

    # ── Section 6: Rate by Grade ──
    next_row, grade_layout = _write_rate_by_grade_section(ws, next_row + 1, grade_rates)

    # ── Data validation: Grade dropdown for roster (references grade table) ──
    dv_grade = DataValidation(
        type="list",
        formula1=f"=$A${grade_layout['grade_first_row']}:$A${grade_layout['grade_last_row']}",
        allow_blank=False,
    )
    dv_grade.error = "Please select a valid grade"
    dv_grade.errorTitle = "Invalid Grade"
    grade_col_letter = get_column_letter(COL_ROSTER_GRADE)
    dv_grade.sqref = (
        f"{grade_col_letter}{roster_layout['roster_first_row']}"
        f":{grade_col_letter}{roster_layout['roster_last_row']}"
    )
    ws.add_data_validation(dv_grade)

    # ── Named Ranges ──
    start_row = param_rows["Project Start"]
    dn_start = DefinedName("project_start", attr_text=f"{SHEET_CONFIG}!$B${start_row}")
    wb.defined_names.add(dn_start)

    rate_row = param_rows["FTE Rate (EUR)"]
    dn_rate = DefinedName("fte_rate", attr_text=f"{SHEET_CONFIG}!$B${rate_row}")
    wb.defined_names.add(dn_rate)

    ceiling_row = param_rows["EU Annual Ceiling (PM/FTE/year)"]
    dn_ceiling = DefinedName("eu_annual_ceiling_pm",
                              attr_text=f"{SHEET_CONFIG}!$B${ceiling_row}")
    wb.defined_names.add(dn_ceiling)

    # Grade table named range (2-column: grade name + rate)
    dn_grades = DefinedName(
        "tbl_GradeRates",
        attr_text=f"{SHEET_CONFIG}!$A${grade_layout['grade_first_row']}:$B${grade_layout['grade_last_row']}",
    )
    wb.defined_names.add(dn_grades)

    # Roster column named ranges
    dn_roster_ftes = DefinedName(
        "roster_ftes",
        attr_text=f"{SHEET_CONFIG}!$B${roster_layout['roster_first_row']}:$B${roster_layout['roster_last_row']}",
    )
    wb.defined_names.add(dn_roster_ftes)

    dn_roster_grades = DefinedName(
        "roster_grades",
        attr_text=f"{SHEET_CONFIG}!$C${roster_layout['roster_first_row']}:$C${roster_layout['roster_last_row']}",
    )
    wb.defined_names.add(dn_roster_grades)

    # Schedule table column named ranges
    dn_sched_ids = DefinedName(
        "sched_wp_ids",
        attr_text=f"{SHEET_CONFIG}!$A${sched_layout['sched_first_row']}:$A${sched_layout['sched_last_row']}",
    )
    wb.defined_names.add(dn_sched_ids)

    dn_sched_starts = DefinedName(
        "sched_starts",
        attr_text=f"{SHEET_CONFIG}!$B${sched_layout['sched_first_row']}:$B${sched_layout['sched_last_row']}",
    )
    wb.defined_names.add(dn_sched_starts)

    dn_sched_ends = DefinedName(
        "sched_ends",
        attr_text=f"{SHEET_CONFIG}!$C${sched_layout['sched_first_row']}:$C${sched_layout['sched_last_row']}",
    )
    wb.defined_names.add(dn_sched_ends)

    # Assignment column named ranges
    a_first = assign_layout["assign_first_row"]
    a_last = assign_layout["assign_last_row"]
    wb.defined_names.add(DefinedName(
        "assign_persons",
        attr_text=f"{SHEET_CONFIG}!$A${a_first}:$A${a_last}",
    ))
    wb.defined_names.add(DefinedName(
        "assign_wps",
        attr_text=f"{SHEET_CONFIG}!$B${a_first}:$B${a_last}",
    ))
    wb.defined_names.add(DefinedName(
        "assign_ftes",
        attr_text=f"{SHEET_CONFIG}!$C${a_first}:$C${a_last}",
    ))

    # External commitments column named ranges
    e_first = ext_layout["ext_first_row"]
    e_last = ext_layout["ext_last_row"]
    wb.defined_names.add(DefinedName(
        "ext_persons",
        attr_text=f"{SHEET_CONFIG}!$A${e_first}:$A${e_last}",
    ))
    wb.defined_names.add(DefinedName(
        "ext_ftes",
        attr_text=f"{SHEET_CONFIG}!$C${e_first}:$C${e_last}",
    ))

    # Roster lookup (2-column: person name + max FTE) for VLOOKUP in Balance formula
    wb.defined_names.add(DefinedName(
        "roster_lookup",
        attr_text=f"{SHEET_CONFIG}!$A${roster_layout['roster_first_row']}:$B${roster_layout['roster_last_row']}",
    ))

    # Project months named range (single cell) for Balance formula
    months_row = param_rows["Project Months"]
    wb.defined_names.add(DefinedName(
        "project_months",
        attr_text=f"{SHEET_CONFIG}!$B${months_row}",
    ))

    # ── Set FTE Rate formula ──
    # If the user has authored a custom formula (preserved by the reader as
    # project_params["rate"]), emit it verbatim. Otherwise fall back to the
    # default weighted average of grade rates by FTE.
    existing_rate = project_params.get("rate")
    if isinstance(existing_rate, str) and existing_rate.startswith("="):
        ws.cell(rate_row, 2).value = existing_rate
    else:
        roster_first_row = roster_layout["roster_first_row"]
        roster_last_row = roster_layout["roster_last_row"]
        grade_first_row = grade_layout["grade_first_row"]
        grade_last_row = grade_layout["grade_last_row"]

        grades_ref = f"$C${roster_first_row}:$C${roster_last_row}"
        ftes_ref = f"$B${roster_first_row}:$B${roster_last_row}"
        terms = "+".join(
            f"({grades_ref}=$A${row})*$B${row}"
            for row in range(grade_first_row, grade_last_row + 1)
        )
        ws.cell(rate_row, 2).value = (
            f"=SUMPRODUCT({terms},{ftes_ref})/SUM({ftes_ref})"
        )

    # ── Sheet Appearance ──
    ws.sheet_properties.tabColor = "4472C4"
    ws.freeze_panes = f"A{budget_section_row}"
    auto_fit_columns(ws, max_width=25)

    return {
        "param_rows": param_rows,
        "grade_first_row": grade_layout["grade_first_row"],
        "grade_last_row": grade_layout["grade_last_row"],
        "wp_header_row": budget_layout["wp_header_row"],
        "wp_first_row": budget_layout["wp_first_row"],
        "wp_last_row": budget_layout["wp_last_row"],
        "wp_total_row": budget_layout["wp_total_row"],
        "assign_header_row": assign_layout["assign_header_row"],
        "assign_first_row": assign_layout["assign_first_row"],
        "assign_last_row": assign_layout["assign_last_row"],
        "assign_config_rows": assign_layout["assign_config_rows"],
        "ext_header_row": ext_layout["ext_header_row"],
        "ext_first_row": ext_layout["ext_first_row"],
        "ext_last_row": ext_layout["ext_last_row"],
        "external_config_rows": ext_layout["external_config_rows"],
        "sched_header_row": sched_layout["sched_header_row"],
        "sched_first_row": sched_layout["sched_first_row"],
        "sched_last_row": sched_layout["sched_last_row"],
        "roster_header_row": roster_layout["roster_header_row"],
        "roster_first_row": roster_layout["roster_first_row"],
        "roster_last_row": roster_layout["roster_last_row"],
        "person_rows": roster_layout["person_rows"],
        "rp_first_row": rp_layout["rp_first_row"],
        "rp_last_row": rp_layout["rp_last_row"],
    }
