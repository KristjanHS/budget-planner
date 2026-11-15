"""Tests for Config sheet builder."""

from openpyxl import Workbook
from budget_planner.config import PROJECT_START
from budget_planner.config_sheet import build_config_sheet


def _sample_params():
    return {
        "title": "Test Project",
        "start": PROJECT_START,
        "months": 60,
        "rate": 7500,
        "eu_ceiling_pm": 12.0,
    }


def _sample_wps():
    return [
        {"id": "WP1", "name": "Management", "fte": 12.0, "start": 1, "end": 60},
        {"id": "WP2", "name": "Design", "fte": 8.0, "start": 1, "end": 18},
    ]


def _find_param_value(ws, label_prefix):
    """Find param value by scanning column A for label prefix."""
    for r in range(1, ws.max_row + 1):
        label = str(ws.cell(r, 1).value or "")
        if label_prefix.lower() in label.lower():
            return ws.cell(r, 2).value
    return None


def test_config_sheet_params():
    """Project parameters are written correctly."""
    wb = Workbook()
    ws = wb.active
    assert ws is not None
    ws.title = "Config"
    build_config_sheet(ws, wb, _sample_params(), _sample_wps())
    assert _find_param_value(ws, "Project Title") == "Test Project"
    assert _find_param_value(ws, "Project Start") == PROJECT_START
    assert _find_param_value(ws, "Project Months") == 60
    # FTE Rate is now a formula (weighted average of grade rates)
    from openpyxl.worksheet.formula import ArrayFormula

    fte_rate_val = _find_param_value(ws, "FTE Rate")
    assert isinstance(fte_rate_val, (str, ArrayFormula))


def test_config_sheet_wp_table():
    """Work packages are written as a table."""
    wb = Workbook()
    ws = wb.active
    assert ws is not None
    ws.title = "Config"
    layout = build_config_sheet(ws, wb, _sample_params(), _sample_wps())
    first = layout["wp_first_row"]
    assert ws.cell(first, 1).value == "WP1"
    assert ws.cell(first + 1, 1).value == "WP2"
    # Table exists
    assert len(ws.tables) == 1
    assert "tbl_WorkPackages" in [t.name for t in ws.tables.values()]


def test_config_sheet_named_ranges():
    """Named ranges for project_start and fte_rate exist."""
    wb = Workbook()
    ws = wb.active
    assert ws is not None
    ws.title = "Config"
    build_config_sheet(ws, wb, _sample_params(), _sample_wps())
    names = set(wb.defined_names.keys())
    assert "project_start" in names
    assert "fte_rate" in names


def test_config_sheet_layout_keys():
    """Layout dict has expected keys including roster."""
    wb = Workbook()
    ws = wb.active
    assert ws is not None
    ws.title = "Config"
    layout = build_config_sheet(ws, wb, _sample_params(), _sample_wps())
    assert "wp_header_row" in layout
    assert "wp_first_row" in layout
    assert "wp_last_row" in layout
    assert layout["wp_last_row"] == layout["wp_first_row"] + 1  # 2 WPs
    assert "roster_header_row" in layout
    assert "roster_first_row" in layout
    assert "roster_last_row" in layout
    assert "person_rows" in layout
    assert "rp_first_row" in layout
    assert "rp_last_row" in layout


def test_config_sheet_blue_font():
    """WP cells and param values use blue input font."""
    wb = Workbook()
    ws = wb.active
    assert ws is not None
    ws.title = "Config"
    layout = build_config_sheet(ws, wb, _sample_params(), _sample_wps())
    # Param values should be blue (except row 1 which uses TITLE_FONT)
    start_row = layout["param_rows"]["Project Start"]
    assert ws.cell(start_row, 2).font.color.rgb in ("000000FF", "0000FF")
    # WP cells should be blue
    first = layout["wp_first_row"]
    font_rgb = ws.cell(first, 1).font.color.rgb
    assert "0000FF" in str(font_rgb)


def test_config_eu_ceiling_param():
    """EU Annual Ceiling (PM) parameter row exists with default value."""
    wb = Workbook()
    ws = wb.active
    assert ws is not None
    ws.title = "Config"
    build_config_sheet(ws, wb, _sample_params(), _sample_wps())
    found = False
    for r in range(1, 10):
        if ws.cell(r, 1).value and "EU Annual Ceiling" in str(ws.cell(r, 1).value):
            assert ws.cell(r, 2).value == 12.0
            found = True
            break
    assert found, "EU Annual Ceiling param not found"


def test_config_eu_ceiling_named_range():
    """Named range eu_annual_ceiling_pm exists."""
    wb = Workbook()
    ws = wb.active
    assert ws is not None
    ws.title = "Config"
    build_config_sheet(ws, wb, _sample_params(), _sample_wps())
    names = set(wb.defined_names.keys())
    assert "eu_annual_ceiling_pm" in names


def test_reporting_periods_section():
    """Config sheet has REPORTING PERIODS section with correct data."""
    wb = Workbook()
    ws = wb.active
    assert ws is not None
    ws.title = "Config"
    rps = [
        {"id": "RP1", "start": 1, "end": 6},
        {"id": "RP2", "start": 7, "end": 12},
    ]
    layout = build_config_sheet(ws, wb, _sample_params(), _sample_wps(), reporting_periods=rps)

    # Find REPORTING PERIODS section header
    rp_header_row = None
    for r in range(1, ws.max_row + 1):
        val = ws.cell(r, 1).value
        if val and "REPORTING PERIODS" in str(val).upper():
            rp_header_row = r
            break
    assert rp_header_row is not None, "REPORTING PERIODS section not found"

    # Column headers row
    col_header_row = rp_header_row + 1
    assert ws.cell(col_header_row, 1).value == "Period"
    assert ws.cell(col_header_row, 2).value == "Start Month"
    assert ws.cell(col_header_row, 3).value == "End Month"

    # Data rows — should match seed data count
    first_data_row = col_header_row + 1
    for i, rp in enumerate(rps):
        row = first_data_row + i
        assert ws.cell(row, 1).value == rp["id"]
        assert ws.cell(row, 2).value == rp["start"]
        assert ws.cell(row, 3).value == rp["end"]
        # Blue input font
        assert ws.cell(row, 1).font.color.rgb == "000000FF"

    # Layout dict includes RP rows
    assert layout["rp_first_row"] == first_data_row
    assert layout["rp_last_row"] == first_data_row + len(rps) - 1


def test_config_round_trip_eu_ceiling():
    """EU ceiling param survives write → read round-trip."""
    from budget_planner.reader import read_config_sheet

    wb = Workbook()
    ws = wb.active
    assert ws is not None
    ws.title = "Config"
    params = {**_sample_params(), "eu_ceiling_pm": 10.0}
    build_config_sheet(ws, wb, params, _sample_wps())
    read_params, read_wps = read_config_sheet(ws)
    assert read_params["eu_ceiling_pm"] == 10.0
    assert len(read_wps) == 2  # WPs still read correctly after row shift


# ── Grade-based rates tests ──


def _sample_team():
    return [
        {"name": "Alice", "max_fte": 1.0, "grade": "Senior"},
        {"name": "Bob", "max_fte": 1.0, "grade": "Mid"},
        {"name": "Carol", "max_fte": 0.5, "grade": "Junior"},
    ]


def _sample_grade_rates():
    return {"Senior": 12500, "Mid": 10000, "Junior": 7500}


def _build_with_grades():
    """Build a Config sheet with grade rates and team grades."""
    wb = Workbook()
    ws = wb.active
    assert ws is not None
    ws.title = "Config"
    layout = build_config_sheet(
        ws,
        wb,
        _sample_params(),
        _sample_wps(),
        team=_sample_team(),
        grade_rates=_sample_grade_rates(),
    )
    return wb, ws, layout


def test_grade_rate_table_values():
    """RATE BY GRADE section has correct grade names and rates."""
    wb, ws, layout = _build_with_grades()
    grade_rates = _sample_grade_rates()
    first = layout["grade_first_row"]
    for i, (grade, rate) in enumerate(grade_rates.items()):
        assert ws.cell(first + i, 1).value == grade
        assert ws.cell(first + i, 2).value == rate


def test_roster_grade_column():
    """Roster has Grade column with correct values."""
    _wb, ws, layout = _build_with_grades()
    team = _sample_team()
    # Header
    assert ws.cell(layout["roster_header_row"], 3).value == "Grade"
    # Data
    for i, member in enumerate(team):
        row = layout["roster_first_row"] + i
        assert ws.cell(row, 3).value == member["grade"]


def test_roster_has_grade_and_rate_columns():
    """Roster has Grade (col C, blue input) and Rate (col D, black VLOOKUP)."""
    _wb, ws, layout = _build_with_grades()
    header_row = layout["roster_header_row"]
    assert ws.cell(header_row, 3).value == "Grade"
    assert ws.cell(header_row, 4).value == "Rate (EUR/PM) -i"
    first = layout["roster_first_row"]
    # Grade is blue user input.
    grade_font = str(ws.cell(first, 3).font.color.rgb).upper()
    assert grade_font.endswith("0000FF"), f"Grade should be blue input, got {grade_font}"
    # Rate is a black formula doing VLOOKUP against tbl_GradeRates.
    rate_val = ws.cell(first, 4).value
    assert isinstance(rate_val, str) and rate_val.startswith("="), f"Expected Rate formula, got {rate_val!r}"
    assert "VLOOKUP" in rate_val
    assert "tbl_GradeRates" in rate_val
    rate_font = str(ws.cell(first, 4).font.color.rgb).upper()
    assert rate_font.endswith("000000"), f"Rate should be black formula, got {rate_font}"
    # Rate number format is EUR.
    assert "€" in ws.cell(first, 4).number_format


def test_roster_lookup_spans_cols_a_to_d():
    """roster_lookup named range covers A:D (name, max_fte, grade, rate)."""
    wb, _ws, layout = _build_with_grades()
    target = wb.defined_names["roster_lookup"].attr_text
    expected = f"Config!$A${layout['roster_first_row']}:$D${layout['roster_last_row']}"
    assert target == expected, f"Expected {expected!r}, got {target!r}"


def test_fte_rate_is_formula():
    """FTE Rate cell contains a SUMPRODUCT weighted-average formula, not a static value."""
    _wb, ws, layout = _build_with_grades()
    rate_row = layout["param_rows"]["FTE Rate (EUR)"]
    val = ws.cell(rate_row, 2).value
    from openpyxl.worksheet.formula import ArrayFormula

    assert isinstance(val, (str, ArrayFormula)), f"Expected formula, got: {val}"
    formula_text = str(val.text if isinstance(val, ArrayFormula) else val)
    assert "SUMPRODUCT" in formula_text


def test_fte_rate_font_is_black():
    """FTE Rate cell uses BLACK_FORMULA font, not BLUE_INPUT."""
    _wb, ws, layout = _build_with_grades()
    rate_row = layout["param_rows"]["FTE Rate (EUR)"]
    font = ws.cell(rate_row, 2).font
    rgb = str(font.color.rgb).upper()
    assert rgb.endswith("000000"), f"Expected black font, got rgb={rgb}"


def test_tbl_grade_rates_named_range():
    """tbl_GradeRates named range exists and covers grade data rows."""
    wb, _ws, layout = _build_with_grades()
    names = {dn.name: dn.attr_text for name in wb.defined_names for dn in [wb.defined_names[name]]}
    assert "tbl_GradeRates" in names
    expected = f"Config!$A${layout['grade_first_row']}:$B${layout['grade_last_row']}"
    assert names["tbl_GradeRates"] == expected


def test_roster_ftes_named_range():
    """roster_ftes named range exists and covers FTE data column."""
    wb, _ws, layout = _build_with_grades()
    names = {dn.name: dn.attr_text for name in wb.defined_names for dn in [wb.defined_names[name]]}
    assert "roster_ftes" in names
    expected = f"Config!$B${layout['roster_first_row']}:$B${layout['roster_last_row']}"
    assert names["roster_ftes"] == expected


def test_roster_grades_named_range():
    """roster_grades named range exists and covers Grade data column."""
    wb, _ws, layout = _build_with_grades()
    names = {dn.name: dn.attr_text for name in wb.defined_names for dn in [wb.defined_names[name]]}
    assert "roster_grades" in names
    expected = f"Config!$C${layout['roster_first_row']}:$C${layout['roster_last_row']}"
    assert names["roster_grades"] == expected


def test_grade_data_validation():
    """Grade cells have data validation dropdown with grade names."""
    _wb, ws, layout = _build_with_grades()
    # Check that data validation exists on the worksheet
    assert len(ws.data_validations.dataValidation) > 0
    # Find the list validation
    dv_found = False
    for dv in ws.data_validations.dataValidation:
        if dv.type == "list" and str(dv.formula1).startswith("=$A$"):
            dv_found = True
            break
    assert dv_found, "Data validation dropdown for grades not found"


def test_layout_includes_grade_rows():
    """config_layout dict includes grade_first_row and grade_last_row."""
    _wb, _ws, layout = _build_with_grades()
    assert "grade_first_row" in layout
    assert "grade_last_row" in layout
    grade_rates = _sample_grade_rates()
    assert layout["grade_last_row"] - layout["grade_first_row"] + 1 == len(grade_rates)


# ── Assignment and External section tests ──


def _sample_assignments():
    return [
        {"person": "Alice", "wp": "WP1", "total_fte": 4.5, "start": 1, "end": 60},
        {"person": "Bob", "wp": "WP2", "total_fte": 3.0, "start": 1, "end": 18},
    ]


def _sample_externals():
    return [
        {"person": "Alice", "project": "ProjectX", "fte": 2.0, "start": 1, "end": 11, "is_eu": False},
    ]


def _build_with_assignments():
    """Build a Config sheet with assignments and external commitments."""
    wb = Workbook()
    ws = wb.active
    assert ws is not None
    ws.title = "Config"
    layout = build_config_sheet(
        ws,
        wb,
        _sample_params(),
        _sample_wps(),
        team=_sample_team(),
        grade_rates=_sample_grade_rates(),
        assignments=_sample_assignments(),
        external_commitments=_sample_externals(),
    )
    return wb, ws, layout


def test_config_assignments_section_exists():
    """PERSON-WP ASSIGNMENTS section header and column headers are present."""
    _wb, ws, layout = _build_with_assignments()
    # Section header is one row above the column header row
    section_row = layout["assign_header_row"] - 1
    assert ws.cell(section_row, 1).value == "PERSON-WP ASSIGNMENTS"
    # Column headers
    header_row = layout["assign_header_row"]
    assert ws.cell(header_row, 1).value == "Person"
    assert ws.cell(header_row, 2).value == "WP"
    assert ws.cell(header_row, 3).value == "Total FTE"
    assert ws.cell(header_row, 8).value == "Balance"


def test_config_assignments_data_rows():
    """Assignment data rows have blue input cells with correct values."""
    _wb, ws, layout = _build_with_assignments()
    assignments = _sample_assignments()
    first = layout["assign_first_row"]
    for i, entry in enumerate(assignments):
        row = first + i
        assert ws.cell(row, 1).value == entry["person"]
        assert ws.cell(row, 2).value == entry["wp"]
        assert ws.cell(row, 3).value == entry["total_fte"]
        assert ws.cell(row, 4).value == entry["start"]
        assert ws.cell(row, 5).value == entry["end"]
        # Blue input font on cols A-E
        assert "0000FF" in str(ws.cell(row, 1).font.color.rgb)


def test_config_assignments_formula_columns():
    """Assignment cols F-H have formulas with correct patterns."""
    _wb, ws, layout = _build_with_assignments()
    first = layout["assign_first_row"]
    # Col F: Sig Assigned — SUMIFS
    val_f = str(ws.cell(first, 6).value)
    assert "SUMIFS" in val_f
    assert "assign_ftes" in val_f
    assert "assign_persons" in val_f
    # Col G: Sig External — SUMIFS
    val_g = str(ws.cell(first, 7).value)
    assert "SUMIFS" in val_g
    assert "ext_ftes" in val_g
    assert "ext_persons" in val_g
    # Col H: Balance — VLOOKUP
    val_h = str(ws.cell(first, 8).value)
    assert "VLOOKUP" in val_h
    assert "roster_lookup" in val_h
    assert "project_months" in val_h
    # Black formula font on cols F-H
    rgb_f = str(ws.cell(first, 6).font.color.rgb).upper()
    assert rgb_f.endswith("000000")


def test_config_externals_section_exists():
    """EXTERNAL COMMITMENTS section header and column headers are present."""
    _wb, ws, layout = _build_with_assignments()
    section_row = layout["ext_header_row"] - 1
    assert ws.cell(section_row, 1).value == "EXTERNAL COMMITMENTS"
    header_row = layout["ext_header_row"]
    assert ws.cell(header_row, 1).value == "Person"
    assert ws.cell(header_row, 2).value == "Project"
    assert ws.cell(header_row, 6).value == "EU Grant"


def test_config_externals_data_rows():
    """External commitment data rows match input."""
    _wb, ws, layout = _build_with_assignments()
    externals = _sample_externals()
    first = layout["ext_first_row"]
    for i, entry in enumerate(externals):
        row = first + i
        assert ws.cell(row, 1).value == entry["person"]
        assert ws.cell(row, 2).value == entry["project"]
        assert ws.cell(row, 3).value == entry["fte"]
        assert ws.cell(row, 4).value == entry["start"]
        assert ws.cell(row, 5).value == entry["end"]
        assert ws.cell(row, 6).value == "No"  # is_eu=False
        # Blue input font
        assert "0000FF" in str(ws.cell(row, 1).font.color.rgb)


def test_config_wp_budget_expanded_to_11_cols():
    """WP Budget has 11 columns (Budget -i + Budget +i + Start/End/Type); table ref covers A:K."""
    _wb, ws, layout = _build_with_assignments()
    header_row = layout["wp_header_row"]
    # Col 5/6 = Budget -i (orange, formula) / Budget +i (green, blue-input).
    assert ws.cell(header_row, 5).value == "Budget -i"
    assert ws.cell(header_row, 6).value == "Budget +i"
    # Assignment/variance/schedule cols shift right by one.
    assert ws.cell(header_row, 7).value == "Assigned PM"
    assert ws.cell(header_row, 8).value == "Variance"
    assert ws.cell(header_row, 9).value == "Start"
    assert ws.cell(header_row, 10).value == "End"
    assert ws.cell(header_row, 11).value == "Type"
    # Table ref covers A:K
    tables = list(ws.tables.values())
    assert len(tables) == 1
    table_ref = tables[0].ref
    assert table_ref.endswith(f"K{layout['wp_last_row']}")


def test_config_wp_budget_assigned_pm_formula():
    """WP Budget col G has SUMIFS formula referencing assign named ranges."""
    _wb, ws, layout = _build_with_assignments()
    first = layout["wp_first_row"]
    val = str(ws.cell(first, 7).value)
    assert "SUMIFS" in val
    assert "assign_ftes" in val
    assert "assign_wps" in val


def test_config_wp_budget_variance_formula():
    """WP Budget col H = C - G (FTE Budget minus Assigned PM)."""
    _wb, ws, layout = _build_with_assignments()
    first = layout["wp_first_row"]
    val = str(ws.cell(first, 8).value)
    # Should reference col C (FTE) and col G (Assigned PM) in same row
    assert f"C{first}" in val
    assert f"G{first}" in val


def test_config_wp_budget_minus_i_is_derived_from_plus_i():
    """Budget -i (col E) = Budget +i / 1.25; Budget +i (col F) = structured-ref
    formula into ``tbl_WorkPackages`` computing FTE × Rate(-i) × 1.25."""
    _wb, ws, layout = _build_with_assignments()
    first = layout["wp_first_row"]
    val = str(ws.cell(first, 5).value)
    assert val.startswith("="), f"Budget -i should be a formula, got {val!r}"
    assert f"F{first}/1.25" in val
    # Budget +i is a structured-ref formula — resolves on recalc via Excel
    # Table column names. Literal seeding was retired in the Aiko-port (Q-F).
    plus_i = ws.cell(first, 6).value
    plus_i_str = str(plus_i)
    assert plus_i_str.startswith("="), f"Budget +i should be a formula, got {plus_i!r}"
    assert "tbl_WorkPackages" in plus_i_str
    assert "[FTE Budget]" in plus_i_str
    assert "[Rate (€/PM) -i]" in plus_i_str
    assert "1.25" in plus_i_str


def test_config_wp_budget_with_indirect_named_range():
    """``wp_budget_with_indirect`` named range targets the Budget +i column (F)."""
    wb, _ws, layout = _build_with_assignments()
    assert "wp_budget_with_indirect" in wb.defined_names
    attr = wb.defined_names["wp_budget_with_indirect"].attr_text
    assert f"Config!$F${layout['wp_first_row']}:$F${layout['wp_last_row']}" in attr, (
        f"unexpected named-range target: {attr!r}"
    )


def test_config_new_named_ranges():
    """New named ranges for assignments and externals exist."""
    wb, _ws, _layout = _build_with_assignments()
    names = set(wb.defined_names.keys())
    expected = {
        "assign_persons",
        "assign_wps",
        "assign_ftes",
        "ext_persons",
        "ext_ftes",
        "roster_lookup",
        "project_months",
    }
    for name in expected:
        assert name in names, f"Named range '{name}' not found"


def test_config_assignments_empty_list():
    """When assignments=[], section has ASSIGN_MIN_ROWS blank template rows with formulas."""
    from budget_planner.layout import ASSIGN_MIN_ROWS

    wb = Workbook()
    ws = wb.active
    assert ws is not None
    ws.title = "Config"
    layout = build_config_sheet(
        ws,
        wb,
        _sample_params(),
        _sample_wps(),
        team=_sample_team(),
        grade_rates=_sample_grade_rates(),
        assignments=[],
        external_commitments=[],
    )
    first = layout["assign_first_row"]
    last = layout["assign_last_row"]
    assert last - first + 1 == ASSIGN_MIN_ROWS
    # Person col is blank in template rows; F/G/H carry formulas.
    # Balance IF-guards blank Person so VLOOKUP doesn't return #N/A and CF
    # rules evaluate "" < 0 as FALSE (no false-positive red fill).
    for r in range(first, last + 1):
        assert ws.cell(r, 1).value is None
        assert str(ws.cell(r, 6).value).startswith("=SUMIFS(assign_ftes")
        assert str(ws.cell(r, 7).value).startswith("=SUMIFS(ext_ftes")
        bal = str(ws.cell(r, 8).value)
        assert bal.startswith(f'=IF(A{r}="","",')
        assert f"VLOOKUP(A{r}" in bal


def test_config_assignments_cf_refs_row_local():
    """CF rules on blank template rows reference row-local cells, so blank
    Person yields ""<0 (FALSE) and never triggers the red/yellow fill."""
    from budget_planner.layout import ASSIGN_MIN_ROWS

    wb = Workbook()
    ws = wb.active
    assert ws is not None
    ws.title = "Config"
    layout = build_config_sheet(
        ws,
        wb,
        _sample_params(),
        _sample_wps(),
        team=_sample_team(),
        grade_rates=_sample_grade_rates(),
        assignments=[],
        external_commitments=[],
    )
    first = layout["assign_first_row"]
    last = layout["assign_last_row"]
    assert last - first + 1 == ASSIGN_MIN_ROWS

    # Collect CF rule formulas keyed by the cell they apply to.
    rules_by_cell: dict[str, list[str]] = {}
    for cf in ws.conditional_formatting:
        rng_str = str(cf.sqref)
        for rule in ws.conditional_formatting[cf]:
            if rule.formula:
                rules_by_cell.setdefault(rng_str, []).extend(rule.formula)

    # Every assignment row must have at least the red (<0) + yellow (<0.1*...) rule,
    # and each rule must reference the SAME row it applies to (row-local).
    for r in range(first, last + 1):
        target = f"H{r}"
        assert target in rules_by_cell, f"No CF on {target}"
        formulas = rules_by_cell[target]
        assert any(f"H{r}<0" in f for f in formulas), f"Row {r} missing red rule or ref is not row-local: {formulas}"
        assert any(f"A{r}" in f and "VLOOKUP" in f for f in formulas), (
            f"Row {r} yellow rule missing row-local Person ref: {formulas}"
        )


def test_config_layout_includes_new_keys():
    """Layout dict has assign_* and ext_* keys."""
    _wb, _ws, layout = _build_with_assignments()
    for key in [
        "assign_header_row",
        "assign_first_row",
        "assign_last_row",
        "assign_config_rows",
        "ext_header_row",
        "ext_first_row",
        "ext_last_row",
        "external_config_rows",
    ]:
        assert key in layout, f"Layout key '{key}' not found"
    # Verify assign_config_rows has expected entries
    assert ("Alice", "WP1") in layout["assign_config_rows"]
    assert ("Bob", "WP2") in layout["assign_config_rows"]
    # Verify external_config_rows has expected entries
    assert ("Alice", "ProjectX") in layout["external_config_rows"]


def test_hours_per_fte_month_constant_and_named_range(tmp_path):
    from budget_planner.workbook import init_workbook
    from openpyxl import load_workbook

    path = tmp_path / "wb.xlsx"
    init_workbook(str(path))
    wb = load_workbook(str(path))

    assert "hours_per_fte_month" in wb.defined_names
    target = wb.defined_names["hours_per_fte_month"].attr_text
    assert target.startswith("Config!")

    sheet, cell_ref = target.split("!")
    cfg = wb["Config"]
    row = int(cell_ref.replace("$", "").lstrip("B"))
    assert cfg.cell(row=row, column=2).value == 143

    # Label in col A on the same row
    assert cfg.cell(row=row, column=1).value == "Hours per FTE Month"


def test_through_month_named_ranges_point_at_actuals_sheets(tmp_path):
    """Post 2026-04-18 cleanup: both through-month named ranges live on the
    Actuals / Actuals-Hours sheets (row 1 col D = effective cell), not Config.

    Derived helpers (months_elapsed_hours, months_elapsed_eur,
    reconciled_through_month) were dropped entirely — each callsite inlines
    MAX(1, ...) / MIN(..., ...) directly.
    """
    from budget_planner.workbook import init_workbook
    from openpyxl import load_workbook

    path = tmp_path / "wb.xlsx"
    init_workbook(str(path))
    wb = load_workbook(str(path))

    names = wb.defined_names
    assert names["eur_actuals_through_month"].attr_text == "Actuals!$D$1"
    assert names["hours_actuals_through_month"].attr_text == ("'Actuals-Hours'!$D$1")
    # Derived helpers are gone.
    for dropped in (
        "months_elapsed_hours",
        "months_elapsed_eur",
        "reconciled_through_month",
    ):
        assert dropped not in names, f"{dropped} should be inlined at callsites, not a named range"


# ── Version stamp (post-v2 follow-up) ──


def test_version_stamp_default_and_named_range():
    """Default (no version passed) is 'unknown'; named range points at B10.

    Row shifted from B7 → B9 (2026-04-19 Timelog Plan enhancements) → B10
    (2026-04-23 solver-freeze-point design inserted ``Rebuild from month #``
    immediately before Version Stamp).
    """
    wb = Workbook()
    ws = wb.active
    assert ws is not None
    ws.title = "Config"
    build_config_sheet(ws, wb, _sample_params(), _sample_wps())

    # Default value is the "unknown" sentinel (not None, not "").
    # Proves the builder always writes a non-empty string so auditors see
    # a concrete stamp even when the orchestrator doesn't pass git context
    # (e.g. build_from_fingerprint on the synthetic-rebuild path).
    assert ws.cell(10, 1).value == "Version Stamp"
    assert ws.cell(10, 2).value == "unknown"
    assert "version_stamp" in wb.defined_names
    assert wb.defined_names["version_stamp"].attr_text == "Config!$B$10"


def test_version_stamp_explicit_value_round_trips():
    """Explicit value threaded from the orchestrator lands in B10 verbatim."""
    wb = Workbook()
    ws = wb.active
    assert ws is not None
    ws.title = "Config"
    build_config_sheet(
        ws,
        wb,
        _sample_params(),
        _sample_wps(),
        version_stamp="v1.2.3-4-gabcdef1-dirty",
    )
    # The value is written as a literal string (not as a formula starting
    # with "=") — the cell is audit metadata, not a derived formula.
    assert ws.cell(10, 2).value == "v1.2.3-4-gabcdef1-dirty"


def test_resolve_version_stamp_fallback_on_git_failure(monkeypatch):
    """Resolver returns 'unknown' when the git subprocess can't run."""
    from budget_planner import workbook as _wb

    def _boom(*args, **kwargs):
        raise FileNotFoundError("git: not found")

    monkeypatch.setattr(_wb.subprocess, "run", _boom)
    assert _wb._resolve_version_stamp() == "unknown"


# ── Roster Priority column (Task T1 of Timelog Plan) ──


def test_roster_priority_header():
    """Priority column header has label, HEADER_FONT, and LIGHT_BLUE_FILL."""
    from budget_planner.layout import COL_ROSTER_PRIORITY

    _wb, ws, layout = _build_with_grades()
    header_row = layout["roster_header_row"]
    cell = ws.cell(header_row, COL_ROSTER_PRIORITY)
    assert cell.value == "Priority"
    # HEADER_FONT is bold
    assert cell.font.bold, "Priority header should use HEADER_FONT (bold)"
    # LIGHT_BLUE_FILL — match by fill fgColor (same fill used on other roster headers)
    grade_fill_rgb = ws.cell(header_row, 3).fill.fgColor.rgb
    assert cell.fill.fgColor.rgb == grade_fill_rgb, (
        "Priority header should share LIGHT_BLUE_FILL with other roster headers"
    )


def test_roster_priority_data_defaults_to_2_blue():
    """Priority data cells default to 2 with BLUE_INPUT font when unspecified."""
    from budget_planner.layout import COL_ROSTER_PRIORITY

    _wb, ws, layout = _build_with_grades()
    team = _sample_team()
    first = layout["roster_first_row"]
    for i in range(len(team)):
        row = first + i
        cell = ws.cell(row, COL_ROSTER_PRIORITY)
        assert cell.value == 2, f"row {row}: expected default priority 2, got {cell.value!r}"
        rgb = str(cell.font.color.rgb).upper()
        assert rgb.endswith("0000FF"), f"row {row}: priority should be BLUE_INPUT, got {rgb}"


def test_team_priorities_named_range():
    """team_priorities named range resolves to the Priority column of roster rows."""
    from openpyxl.utils import get_column_letter
    from budget_planner.layout import COL_ROSTER_PRIORITY

    wb, _ws, layout = _build_with_grades()
    assert "team_priorities" in wb.defined_names
    cl = get_column_letter(COL_ROSTER_PRIORITY)
    expected = f"Config!${cl}${layout['roster_first_row']}:${cl}${layout['roster_last_row']}"
    assert wb.defined_names["team_priorities"].attr_text == expected


def test_roster_priority_header_has_comment():
    """Priority header cell carries an explanatory comment for solver fill order."""
    from budget_planner.layout import COL_ROSTER_PRIORITY

    _wb, ws, layout = _build_with_grades()
    header_row = layout["roster_header_row"]
    cell = ws.cell(header_row, COL_ROSTER_PRIORITY)
    assert cell.comment is not None, "Priority header should have a comment"
    text = cell.comment.text or ""
    # Explain the 1/2/3 semantics and reference the consumer.
    assert "1" in text and "3" in text
    assert "Timelog Plan" in text


def test_config_named_range_targets_match_spec_doc():
    """Docs-drift detector: Config named-range targets in the generated
    workbook must match the row numbers documented in ``docs/spec.md``
    (Per-Sheet Notes → Config prose).

    When the builder shifts a param row or ``spec.md`` claims a wrong
    row, this test fails — forcing either the builder or the doc to resync.
    The doc is the source of truth that future agents load; a silent skew
    would mislead them about where each named range points.
    """
    import re
    from pathlib import Path

    wb = Workbook()
    ws = wb.active
    assert ws is not None
    ws.title = "Config"
    build_config_sheet(ws, wb, _sample_params(), _sample_wps())

    internals_path = Path(__file__).parent.parent / "docs" / "spec.md"
    text = internals_path.read_text(encoding="utf-8")

    # Named-range cell mappings live in the Cross-Sheet Reference Map table
    # (Config-hosted rows only). Table row form:
    #     | `name` | Config!$B$N | Named range |
    # The three-column shape + literal `Named range` tail makes the pattern
    # specific enough to skip doc-scoping — no other table in the file emits
    # it. (Earlier revisions parsed a `` `name` (B<row>) `` prose form from
    # the **Config** per-sheet paragraph; 2026-04-20 Doc Trim removed that
    # enumeration and redirected readers to the reference map.)
    pairs = re.findall(
        r"\|\s*`([a-z_]+)`\s*\|\s*Config!\$B\$(\d+)\s*\|\s*Named range\s*\|",
        text,
    )
    assert pairs, (
        "No named-range → cell mappings parsed from spec.md Cross-Sheet "
        "Reference Map — doc format drifted from the "
        "`| `name` | Config!$B$N | Named range |` table convention"
    )
    doc_map = {name: int(row) for name, row in pairs}

    # Minimum coverage — these must all be documented. The through-month
    # named ranges moved off Config (2026-04-18 cleanup) and are verified
    # against the Actuals/Actuals-Hours paragraphs elsewhere; the derived
    # helpers were dropped entirely and are not doc-worthy.
    required = {
        "project_start",
        "fte_rate",
    }
    missing = required - doc_map.keys()
    assert not missing, f"docs/spec.md Config paragraph is missing named-range rows for {sorted(missing)}"

    for name, expected_row in doc_map.items():
        assert name in wb.defined_names, (
            f"Named range {name!r} documented in internals.md but missing from the workbook"
        )
        expected_attr = f"Config!$B${expected_row}"
        actual = wb.defined_names[name].attr_text
        assert actual == expected_attr, (
            f"Named range {name!r}: spec.md says {expected_attr!r}, workbook has {actual!r} — update whichever is wrong"
        )
