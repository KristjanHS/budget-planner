"""Tests for Config sheet builder."""
from datetime import datetime
from openpyxl import Workbook
from budget_planner.config_sheet import build_config_sheet


def _sample_params():
    return {
        "title": "Test Project",
        "start": datetime(2026, 7, 1),
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
    assert _find_param_value(ws, "Project Start") == datetime(2026, 7, 1)
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
    layout = build_config_sheet(ws, wb, _sample_params(), _sample_wps(),
                                reporting_periods=rps)

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
        {"name": "Bob",   "max_fte": 1.0, "grade": "Mid"},
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
        ws, wb, _sample_params(), _sample_wps(),
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


def test_grade_rate_table_section_header():
    """RATE BY GRADE section header exists."""
    _wb, ws, _layout = _build_with_grades()
    found = False
    for r in range(1, ws.max_row + 1):
        val = ws.cell(r, 1).value
        if val and "RATE BY GRADE" in str(val).upper():
            found = True
            break
    assert found, "RATE BY GRADE section header not found"


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


def test_fte_rate_is_formula():
    """FTE Rate cell contains a SUMPRODUCT/VLOOKUP formula, not a static value."""
    _wb, ws, layout = _build_with_grades()
    rate_row = layout["param_rows"]["FTE Rate (EUR)"]
    val = ws.cell(rate_row, 2).value
    from openpyxl.worksheet.formula import ArrayFormula
    assert isinstance(val, (str, ArrayFormula)), (
        f"Expected formula, got: {val}"
    )
    formula_text = val.text if isinstance(val, ArrayFormula) else val
    assert "SUMPRODUCT" in formula_text
    assert "VLOOKUP" in formula_text
    assert "roster_grades" in formula_text
    assert "tbl_GradeRates" in formula_text


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
    names = {dn.name: dn.attr_text for name in wb.defined_names
             for dn in [wb.defined_names[name]]}
    assert "tbl_GradeRates" in names
    expected = f"Config!$A${layout['grade_first_row']}:$B${layout['grade_last_row']}"
    assert names["tbl_GradeRates"] == expected


def test_roster_ftes_named_range():
    """roster_ftes named range exists and covers FTE data column."""
    wb, _ws, layout = _build_with_grades()
    names = {dn.name: dn.attr_text for name in wb.defined_names
             for dn in [wb.defined_names[name]]}
    assert "roster_ftes" in names
    expected = f"Config!$B${layout['roster_first_row']}:$B${layout['roster_last_row']}"
    assert names["roster_ftes"] == expected


def test_roster_grades_named_range():
    """roster_grades named range exists and covers Grade data column."""
    wb, _ws, layout = _build_with_grades()
    names = {dn.name: dn.attr_text for name in wb.defined_names
             for dn in [wb.defined_names[name]]}
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
