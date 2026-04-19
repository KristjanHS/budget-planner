"""Config sheet builder — project parameters, work packages, and team roster."""

from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import FormulaRule
from openpyxl.utils import get_column_letter

from budget_planner.layout import (
    ConfigParamRows,
    SHEET_CONFIG,
    SECTION_HEADER_LABELS,
    COL_WP_FTE, COL_WP_RATE, COL_WP_CEILING,
    COL_WP_ASSIGNED_PM, COL_WP_VARIANCE, WP_BUDGET_LAST_COL,
    COL_SCHED_START, COL_SCHED_END, COL_SCHED_TYPE, SCHED_LAST_COL,
    COL_ROSTER_FTE, COL_ROSTER_GRADE, COL_ROSTER_RATE, COL_ROSTER_PRIORITY,
    ROSTER_LAST_COL,
    COL_GRADE_RATE, GRADE_LAST_COL,
    COL_RP_START, COL_RP_END, RP_LAST_COL,
    COL_ASSIGN_PERSON, COL_ASSIGN_WP, COL_ASSIGN_FTE,
    COL_ASSIGN_START, COL_ASSIGN_END,
    COL_ASSIGN_SIG_ASSIGNED, COL_ASSIGN_SIG_EXTERNAL, COL_ASSIGN_BALANCE,
    ASSIGN_LAST_COL, ASSIGN_MIN_ROWS,
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
    sized_comment,
    style_cell,
)


def _apply_header_comments(ws, header_row: int, comments: dict[int, str]) -> None:
    """Attach explanatory comments to header cells. Keys are 1-based column numbers."""
    for col, text in comments.items():
        ws.cell(row=header_row, column=col).comment = sized_comment(text)


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
    _apply_header_comments(ws, budget_header_row, {
        COL_WP_FTE: (
            "Total person-months budgeted for this WP across its entire "
            "active period. Blue = user input."
        ),
        COL_WP_RATE: (
            "EUR per person-month for this WP. Overrides the project default "
            "rate. Blue = user input (can also be a formula referencing the "
            "grade-rates table)."
        ),
        COL_WP_CEILING: (
            "Auditor-visible budget cap in EUR. Formula: FTE Budget × Rate. "
            "Black = calculated — don't edit."
        ),
        COL_WP_ASSIGNED_PM: (
            "Total person-months committed to this WP across all assignment "
            "rows. Formula: SUMIFS(assign_ftes, assign_wps, WP ID). Tells you "
            "how much of this WP has been staffed."
        ),
        COL_WP_VARIANCE: (
            "FTE Budget − Assigned PM. Positive = budget remains unassigned. "
            "Zero = fully committed. Yellow highlight if |variance| > 0.01 "
            "(not fully staffed or over-committed)."
        ),
    })

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

    # Section header — merged across full section width
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=ROSTER_LAST_COL)
    style_cell(ws, r, 1, value=SECTION_HEADER_LABELS["team_roster"],
               font=SECTION_FONT, fill=SECTION_FILL, alignment=CENTER, border=False)
    style_cell(ws, r, 2, fill=SECTION_FILL)
    style_cell(ws, r, COL_ROSTER_GRADE, fill=SECTION_FILL)
    style_cell(ws, r, COL_ROSTER_RATE, fill=SECTION_FILL)
    style_cell(ws, r, COL_ROSTER_PRIORITY, fill=SECTION_FILL)
    r += 1

    # Column headers
    roster_header_row = r
    style_cell(ws, r, 1, "Person", font=HEADER_FONT, fill=LIGHT_BLUE_FILL,
               alignment=CENTER)
    style_cell(ws, r, COL_ROSTER_FTE, "Max FTE", font=HEADER_FONT,
               fill=LIGHT_BLUE_FILL, alignment=CENTER)
    style_cell(ws, r, COL_ROSTER_GRADE, "Grade", font=HEADER_FONT,
               fill=LIGHT_BLUE_FILL, alignment=CENTER)
    style_cell(ws, r, COL_ROSTER_RATE, "Rate (EUR/PM)", font=HEADER_FONT,
               fill=LIGHT_BLUE_FILL, alignment=CENTER)
    priority_header_cell = style_cell(
        ws, r, COL_ROSTER_PRIORITY, "Priority", font=HEADER_FONT,
        fill=LIGHT_BLUE_FILL, alignment=CENTER,
    )
    priority_header_cell.comment = sized_comment(
        "Solver fill order: 1=fill first, 3=fill last, 2=neutral default. "
        "Used by Timelog Plan.",
        author="Claude",
    )
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
        # Rate (EUR/PM) — black formula derived from Grade via tbl_GradeRates.
        grade_cl = get_column_letter(COL_ROSTER_GRADE)
        style_cell(ws, r, COL_ROSTER_RATE,
                   value=f"=VLOOKUP({grade_cl}{r},tbl_GradeRates,2,FALSE)",
                   font=BLACK_FORMULA, fmt="€#,##0")
        style_cell(ws, r, COL_ROSTER_PRIORITY, member.get("priority", 2),
                   font=BLUE_INPUT, fmt="0")
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

    Always writes at least ASSIGN_MIN_ROWS rows: populated entries first, then blank
    template rows (blue-input A-E, formulas in F-H with IF-guarded Balance) so the
    named ranges span a realistic area and users can type directly or insert rows
    without the "partial merge" error from the adjacent EXTERNAL COMMITMENTS header.

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
    _apply_header_comments(ws, assign_header_row, {
        COL_ASSIGN_FTE: (
            "Total person-months this person contributes to this WP across "
            "the window [Start Month, End Month]. The Capacity sheet "
            "distributes this amount ONLY across the WP's active months "
            "within the window — WP SCHEDULE gaps are excluded automatically."
        ),
        COL_ASSIGN_START: (
            "First project month (1-based) this person is allocated to this "
            "WP. Gaps in WP SCHEDULE inside this window are skipped "
            "automatically in the Capacity grid."
        ),
        COL_ASSIGN_END: (
            "Last project month (1-based) this person is allocated to this "
            "WP. Inclusive."
        ),
        COL_ASSIGN_SIG_ASSIGNED: (
            "Σ of Total FTE across all assignment rows with this Person. "
            "Formula: SUMIFS(assign_ftes, assign_persons, Person). "
            "This is the total programme-project commitment for this person."
        ),
        COL_ASSIGN_SIG_EXTERNAL: (
            "Σ of Total FTE across the Externals table for this Person. "
            "Captures commitments outside this programme project (other grants, "
            "contracts, internal work) that reduce available capacity."
        ),
        COL_ASSIGN_BALANCE: (
            "Remaining capacity in person-months = max_fte × project_months "
            "− Sig Assigned − Sig External. "
            "Red highlight if negative (over-committed); yellow if below "
            "10% of total capacity (low headroom)."
        ),
    })
    r += 1

    # Data rows — populated entries + blank template rows up to ASSIGN_MIN_ROWS
    assign_first_row = r
    assign_config_rows: dict[tuple[str, str], int] = {}
    total_rows = max(len(assignments), ASSIGN_MIN_ROWS)
    sig_assigned_cl = get_column_letter(COL_ASSIGN_SIG_ASSIGNED)
    sig_ext_cl = get_column_letter(COL_ASSIGN_SIG_EXTERNAL)
    for i in range(total_rows):
        entry = assignments[i] if i < len(assignments) else None
        # Cols A-E: blue input (empty for template rows)
        style_cell(ws, r, COL_ASSIGN_PERSON,
                   entry["person"] if entry else None,
                   font=BLUE_INPUT, border=False)
        style_cell(ws, r, COL_ASSIGN_WP,
                   entry["wp"] if entry else None,
                   font=BLUE_INPUT, border=False)
        style_cell(ws, r, COL_ASSIGN_FTE,
                   entry["total_fte"] if entry else None,
                   font=BLUE_INPUT, fmt="0.00", border=False)
        style_cell(ws, r, COL_ASSIGN_START,
                   entry["start"] if entry else None,
                   font=BLUE_INPUT, fmt="0", border=False)
        style_cell(ws, r, COL_ASSIGN_END,
                   entry["end"] if entry else None,
                   font=BLUE_INPUT, fmt="0", border=False)
        # Col F: Sig Assigned — SUMIFS returns 0 when Person cell is blank (safe)
        style_cell(ws, r, COL_ASSIGN_SIG_ASSIGNED, font=BLACK_FORMULA, fmt="0.00",
                   border=False).value = f"=SUMIFS(assign_ftes,assign_persons,A{r})"
        # Col G: Sig External — SUMIFS returns 0 when Person cell is blank (safe)
        style_cell(ws, r, COL_ASSIGN_SIG_EXTERNAL, font=BLACK_FORMULA, fmt="0.00",
                   border=False).value = f"=SUMIFS(ext_ftes,ext_persons,A{r})"
        # Col H: Balance — IF-guarded to avoid #N/A on blank-Person template rows
        style_cell(ws, r, COL_ASSIGN_BALANCE, font=BLACK_FORMULA, fmt="0.00",
                   border=False).value = (
            f'=IF(A{r}="","",'
            f"VLOOKUP(A{r},roster_lookup,2,FALSE)*project_months"
            f"-{sig_assigned_cl}{r}-{sig_ext_cl}{r})"
        )
        if entry is not None:
            assign_config_rows[(entry["person"], entry["wp"])] = r
        r += 1
    assign_last_row = r - 1

    # CF on Balance column (applied to all rows, template included — CF formulas
    # compare "" to numbers as FALSE so blank rows never trigger):
    #   red if < 0 (over-committed)
    #   yellow if < 10% of available PM (low remaining capacity)
    # Red is added first so it wins the overlap when Balance < 0.
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

    last_col_letter = get_column_letter(ASSIGN_LAST_COL)
    ws.auto_filter.ref = (
        f"A{assign_header_row}:{last_col_letter}{assign_last_row}"
    )

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
    _apply_header_comments(ws, ext_header_row, {
        COL_EXT_FTE: (
            "Person-months this person spends on the external project across "
            "[Start, End]. Rolled up into Sig External on the Assignments "
            "table — reduces capacity available for programme work."
        ),
        COL_EXT_EU: (
            "TRUE = this external project is another EU-funded grant. "
            "Used by the EU annual ceiling check (~12 PM/year total across "
            "ALL EU grants per person). FALSE = non-EU (still uses capacity "
            "but not counted against the EU ceiling)."
        ),
    })
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
                       assignments=None, external_commitments=None,
                       version_stamp=None):
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
        ("Hours per FTE Month", 143, "0", None),
        # Through-month params are user-edited integers. On rebuild the reader
        # surfaces any existing value via project_params; on fresh init it
        # defaults to 1 (= "only project month 1 is closed"). Hardcoding 1
        # here would clobber the user's edit across rebuild.
        (
            "EUR Actuals Through Month",
            project_params.get("eur_actuals_through_month", 1),
            "0",
            None,
        ),
        (
            "Hours Actuals Through Month",
            project_params.get("hours_actuals_through_month", 1),
            "0",
            None,
        ),
        # Code-written audit receipt — overwritten on every init/rebuild.
        # Default "unknown" survives when no git context is passed in.
        ("Version Stamp", version_stamp or "unknown", None, None),
    ]
    # Param rows are driven by the order of param_defs above. The literal
    # TypedDict below mirrors that order; the assertion inside the styling
    # loop catches any drift if the two ever fall out of sync. When adding a
    # new parameter row, also extend ``ConfigParamRows`` in ``layout.py`` —
    # otherwise the assertion fires with ``maps it to None``.
    param_rows: ConfigParamRows = {
        "Project Title": 1,
        "Project Start": 2,
        "Project Months": 3,
        "FTE Rate (EUR)": 4,
        "EU Annual Ceiling (PM/FTE/year)": 5,
        "Hours per FTE Month": 6,
        "EUR Actuals Through Month": 7,
        "Hours Actuals Through Month": 8,
        "Version Stamp": 9,
    }
    for i, (label, value, fmt, _) in enumerate(param_defs, start=1):
        assert param_rows.get(label) == i, (
            f"param_defs/param_rows drift: {label!r} at index {i} "
            f"but ConfigParamRows maps it to {param_rows.get(label)!r}"
        )
        # Label in column A — HEADER_FONT, no border
        style_cell(ws, i, 1, value=label, font=HEADER_FONT, border=False)
        # Value in column B — BLUE_INPUT, no border (except FTE Rate: set later)
        if label == "FTE Rate (EUR)":
            # FTE Rate is a formula — styled as BLACK_FORMULA, set after grade table
            style_cell(ws, i, 2, font=BLACK_FORMULA, fmt=fmt, border=False)
        elif label == "Version Stamp":
            # Code-written audit receipt — black = do-not-edit, but not a formula
            # (see label comment below). Overwritten by builder on every rebuild.
            style_cell(ws, i, 2, value=value, font=BLACK_FORMULA, fmt=fmt, border=False)
        else:
            font = TITLE_FONT if i == 1 else BLUE_INPUT
            style_cell(ws, i, 2, value=value, font=font, fmt=fmt, border=False)

    # ── Section 1b: Derived parameters (black, formula-driven) ──
    # These sit directly below the blue-input params. Labels in col A use
    # HEADER_FONT; formulas in col B use BLACK_FORMULA (same-sheet refs).
    # Through-month cells hold integer month indices (1 = project month 1),
    # not dates — so `months_elapsed` IS the through-month, floored at 1 to
    # keep to-date-rate denominators safe when no report has landed yet.
    derived_params = [
        ("Months Elapsed (Hours)", "=MAX(1,$B$8)"),
        ("Months Elapsed (EUR)", "=MAX(1,$B$7)"),
        ("Reconciled Through Month", "=MIN($B$7,$B$8)"),
    ]
    derived_first_row = len(param_defs) + 1  # first row after the blue-input block
    for offset, (label, formula) in enumerate(derived_params):
        r = derived_first_row + offset
        style_cell(ws, r, 1, value=label, font=HEADER_FONT, border=False)
        style_cell(ws, r, 2, value=formula, font=BLACK_FORMULA, fmt="0", border=False)

    # ── Label-cell comments for rows A4..A11 ──
    # Each param-row label gets an explanatory comment so the spreadsheet is
    # self-documenting in-tab. Comments are sized to fit their text via
    # ``sized_comment`` (default openpyxl boxes truncate).
    _param_label_comments = {
        param_rows["FTE Rate (EUR)"]: (
            "Project-wide default rate per person-month in EUR. The value "
            "here is a SUMPRODUCT-derived weighted average from the roster "
            "(grade × rate × max FTE). Black = formula — do not hand-edit. "
            "To change it, edit RATE BY GRADE or the per-person grade in "
            "TEAM ROSTER."
        ),
        param_rows["EU Annual Ceiling (PM/FTE/year)"]: (
            "Maximum person-months a team member may charge to EU-funded "
            "projects per calendar year — across ALL EU grants, not just "
            "this one. Drives the Capacity EU TOTAL per-year conditional "
            "formatting and the EU ceiling pro-rated check."
        ),
        param_rows["Hours per FTE Month"]: (
            "Convention for how many hours equal 1 person-month of "
            "full-time effort. EU/programme defaults to 143 h/month. Used to "
            "convert hours logged on Actuals-Hours into person-months on "
            "Capacity and into EUR on Plan-Forecast. Blue = user input."
        ),
        param_rows["EUR Actuals Through Month"]: (
            "Latest project month for which WP-level Actual EUR figures "
            "from finance close are trusted and complete. Integer month "
            "index, 1-based. Actual EUR typically lags hours by 2–3 months "
            "(finance-close cadence). Reconciliation and summary views use "
            "this as the trust boundary for EUR-side cumulative figures. "
            "Blue = user input."
        ),
        param_rows["Hours Actuals Through Month"]: (
            "Latest project month for which the hours timelog on "
            "Actuals-Hours is trusted and complete. Integer month index, "
            "1-based. Because hours are entered skip-and-cumulate "
            "(intermediate months blank, cumulative figure dropped into a "
            "later month), the only safe aggregate is cumulative through "
            "this month. Blue = user input."
        ),
        param_rows["Version Stamp"]: (
            "Commit identifier of the code that produced this workbook. "
            "Written by init/rebuild (git describe --always --dirty with "
            "fallback to 'unknown'). Maps an older xlsx back to the exact "
            "builder version for reproducing bugs. Black = do-not-edit "
            "(overwritten on every rebuild)."
        ),
        derived_first_row: (
            "Derived from Hours Actuals Through Month (=MAX(1,$B$8)). "
            "Used as the denominator for to-date rate calculations (Avg "
            "FTE to date, Utilisation to date). Floored at 1 to keep "
            "denominators safe when no hours have been reported yet. "
            "Black = formula."
        ),
        derived_first_row + 1: (
            "Derived from EUR Actuals Through Month (=MAX(1,$B$7)). "
            "Used by the Breach-In cumulative-slope extrapolation on the "
            "Actuals sheet. Floored at 1. Black = formula."
        ),
        derived_first_row + 2: (
            "=MIN($B$7,$B$8) — the boundary beyond which a "
            "Forecast-vs-Actual head-to-head reconciliation cannot be "
            "trusted. Summary table on Trends and the gap chart use this "
            "as their cutoff. Black = formula."
        ),
    }
    for row_num, text in _param_label_comments.items():
        ws.cell(row=row_num, column=1).comment = sized_comment(text)

    # ── Section 2: WP Budget ──
    # Gap row sits AFTER the derived block; budget section starts two rows below
    # the last derived row to preserve the existing visual separator.
    budget_section_row = len(param_defs) + len(derived_params) + 2
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

    hours_row = param_rows["Hours per FTE Month"]
    dn_hours = DefinedName(
        "hours_per_fte_month",
        attr_text=f"{SHEET_CONFIG}!$B${hours_row}",
    )
    wb.defined_names.add(dn_hours)

    eur_through_row = param_rows["EUR Actuals Through Month"]
    dn_eur_through = DefinedName(
        "eur_actuals_through_month",
        attr_text=f"{SHEET_CONFIG}!$B${eur_through_row}",
    )
    wb.defined_names.add(dn_eur_through)

    hours_through_row = param_rows["Hours Actuals Through Month"]
    wb.defined_names.add(DefinedName(
        "hours_actuals_through_month",
        attr_text=f"{SHEET_CONFIG}!$B${hours_through_row}",
    ))
    version_stamp_row = param_rows["Version Stamp"]
    wb.defined_names.add(DefinedName(
        "version_stamp",
        attr_text=f"{SHEET_CONFIG}!$B${version_stamp_row}",
    ))
    # Derived cells — positions follow derived_first_row contract above.
    wb.defined_names.add(DefinedName(
        "months_elapsed_hours",
        attr_text=f"{SHEET_CONFIG}!$B${derived_first_row}",
    ))
    wb.defined_names.add(DefinedName(
        "months_elapsed_eur",
        attr_text=f"{SHEET_CONFIG}!$B${derived_first_row + 1}",
    ))
    wb.defined_names.add(DefinedName(
        "reconciled_through_month",
        attr_text=f"{SHEET_CONFIG}!$B${derived_first_row + 2}",
    ))

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

    # Roster lookup (4-column: name, max FTE, grade, rate) — col 2 = Max FTE
    # used by the Assignments Balance formula; col 4 = Rate EUR/PM used by the
    # Actuals-Hours forecast-EUR leaves. Priority (col E) is intentionally
    # excluded — exposed separately via the `team_priorities` range below so
    # consumers that VLOOKUP into roster_lookup don't collide with the new col.
    wb.defined_names.add(DefinedName(
        "roster_lookup",
        attr_text=f"{SHEET_CONFIG}!$A${roster_layout['roster_first_row']}:$D${roster_layout['roster_last_row']}",
    ))

    # Roster Priority column — consumed by the Timelog Plan solver (fill order).
    priority_cl = get_column_letter(COL_ROSTER_PRIORITY)
    wb.defined_names.add(DefinedName(
        "team_priorities",
        attr_text=(
            f"{SHEET_CONFIG}!${priority_cl}${roster_layout['roster_first_row']}"
            f":${priority_cl}${roster_layout['roster_last_row']}"
        ),
    ))

    # Project months named range (single cell) for Balance formula
    months_row = param_rows["Project Months"]
    wb.defined_names.add(DefinedName(
        "project_months",
        attr_text=f"{SHEET_CONFIG}!$B${months_row}",
    ))

    # ── Set FTE Rate formula ──
    # Always regenerate from the live grade/roster layout. B4 is a read-only
    # dashboard readout (no downstream formulas consume ``fte_rate`` after the
    # 2026-04 Config restructure), so preserving a user-authored formula
    # across rebuilds would only break silently when section rows shift.
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
