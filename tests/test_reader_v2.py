"""Tests for reader extensions: read_plan_config,
read_team_from_config, read_capacity_sheet, read_capacity_assignments."""
import pathlib

import pytest
from openpyxl import Workbook, load_workbook

from budget_planner.reader import (
    discover_layout,
    read_plan_config,
    read_team_from_config,
    read_capacity_sheet,
    read_assignments_from_config,
    read_config_sheet,
)
from budget_planner.config import (
    WORK_PACKAGES, TEAM, PROJECT_MONTHS, EXTERNAL_COMMITMENTS,
    DEMO_CAPACITY_ASSIGNMENTS,
)
from budget_planner.layout import COL_WP_RATE, SECTION_HEADER_LABELS
from budget_planner.workbook import init_workbook


# ── read_plan_config ──


def test_read_plan_config_extracts_wp_definitions(fresh_wb_path):
    """read_plan_config returns list matching WORK_PACKAGES."""
    import openpyxl
    wb = openpyxl.load_workbook(fresh_wb_path)
    layout = discover_layout(wb["Plan"])
    wps = read_plan_config(wb["Plan"], layout)

    assert len(wps) == len(WORK_PACKAGES)
    for got, expected in zip(wps, WORK_PACKAGES):
        assert got["id"] == expected["id"]
        assert got["name"] == expected["name"]
        assert got["fte"] == pytest.approx(expected["fte"])
    wb.close()


# ── read_team_from_config ──


def test_read_team_from_config_names_and_fte(fresh_wb_path):
    """read_team_from_config returns roster with correct names and max_fte."""
    import openpyxl
    wb = openpyxl.load_workbook(fresh_wb_path)
    roster = read_team_from_config(wb["Config"])

    assert len(roster) == len(TEAM)
    for got, expected in zip(roster, TEAM):
        assert got["name"] == expected["name"]
        assert got["max_fte"] == pytest.approx(expected["max_fte"])
    wb.close()


# ── read_capacity_sheet (with assignments) ──


def test_read_capacity_discovers_person_groups(assigned_wb_path):
    """read_capacity_sheet finds all persons."""
    import openpyxl
    wb = openpyxl.load_workbook(assigned_wb_path)
    layout = discover_layout(wb["Plan"])
    config_wps = read_plan_config(wb["Plan"], layout)
    cap = read_capacity_sheet(wb["Capacity"], config_wps)

    expected_names = {m["name"] for m in TEAM}
    assert set(cap["person_groups"].keys()) == expected_names
    wb.close()


def test_read_capacity_classifies_grant_vs_external_rows(assigned_wb_path):
    """First person's programme WP rows match their assignments."""
    import openpyxl
    wb = openpyxl.load_workbook(assigned_wb_path)
    layout = discover_layout(wb["Plan"])
    config_wps = read_plan_config(wb["Plan"], layout)
    cap = read_capacity_sheet(wb["Capacity"], config_wps)

    member = TEAM[0]
    group = cap["person_groups"][member["name"]]
    # Person has only their assigned WPs
    expected_wps = {a["wp"] for a in DEMO_CAPACITY_ASSIGNMENTS
                    if a["person"] == member["name"]}
    assert set(group["wp_values"].keys()) == expected_wps
    expected_ext = {
        e["project"] for e in EXTERNAL_COMMITMENTS
        if e["person"] == member["name"]
    }
    assert set(group["external_values"].keys()) == expected_ext
    wb.close()


def test_read_capacity_assigned_values_are_distributed(assigned_wb_path):
    """Assigned WP monthly values have auto-distributed non-None values."""
    import openpyxl
    wb = openpyxl.load_workbook(assigned_wb_path)
    layout = discover_layout(wb["Plan"])
    config_wps = read_plan_config(wb["Plan"], layout)
    cap = read_capacity_sheet(wb["Capacity"], config_wps)

    # Pick first person's first assignment
    asgn = DEMO_CAPACITY_ASSIGNMENTS[0]
    values = cap["person_groups"][asgn["person"]]["wp_values"][asgn["wp"]]
    assert len(values) == PROJECT_MONTHS
    # Active months should have non-None values
    active_values = [values[m - 1] for m in range(asgn["start"], asgn["end"] + 1)]
    assert all(v is not None for v in active_values), (
        "Expected distributed values, got None at some active months"
    )
    # Total should match total_fte
    assert sum(v for v in active_values if v is not None) == pytest.approx(
        asgn["total_fte"], rel=0.01
    )
    wb.close()


def test_read_capacity_external_commitments_round_trip(assigned_wb_path):
    """read_capacity_sheet extracts external_commitments matching config."""
    import openpyxl
    wb = openpyxl.load_workbook(assigned_wb_path)
    layout = discover_layout(wb["Plan"])
    config_wps = read_plan_config(wb["Plan"], layout)
    cap = read_capacity_sheet(wb["Capacity"], config_wps)

    ext = cap["external_commitments"]
    assert len(ext) == len(EXTERNAL_COMMITMENTS)
    for got, expected in zip(ext, EXTERNAL_COMMITMENTS):
        assert got["person"] == expected["person"]
        assert got["project"] == expected["project"]
        assert got["fte"] == pytest.approx(expected["fte"])
        assert got["start"] == expected["start"]
        assert got["end"] == expected["end"]
    wb.close()


# ── read_assignments_from_config ──


def test_read_assignments_from_config_round_trip(assigned_wb_path):
    """Assignments section on Config round-trips correctly."""
    import openpyxl
    wb = openpyxl.load_workbook(assigned_wb_path)
    assignments = read_assignments_from_config(wb["Config"])

    assert len(assignments) == len(DEMO_CAPACITY_ASSIGNMENTS)
    for got, expected in zip(assignments, DEMO_CAPACITY_ASSIGNMENTS):
        assert got["person"] == expected["person"]
        assert got["wp"] == expected["wp"]
        assert got["total_fte"] == pytest.approx(expected["total_fte"])
        assert got["start"] == expected["start"]
        assert got["end"] == expected["end"]
    wb.close()


def test_read_assignments_from_config_empty_for_plain_init(fresh_wb_path):
    """Plain init workbook has no assignments in the Config section."""
    import openpyxl
    wb = openpyxl.load_workbook(fresh_wb_path)
    assignments = read_assignments_from_config(wb["Config"])
    assert assignments == []
    wb.close()


# ── Init workbook (no assignments) ──


def test_init_capacity_no_wp_rows(fresh_wb_path):
    """Init workbook without assignments has no WP rows per person."""
    import openpyxl
    wb = openpyxl.load_workbook(fresh_wb_path)
    layout = discover_layout(wb["Plan"])
    config_wps = read_plan_config(wb["Plan"], layout)
    cap = read_capacity_sheet(wb["Capacity"], config_wps)

    assert len(cap["person_groups"]) > 0, "Should discover at least one person"

    for name, group in cap["person_groups"].items():
        # No WP rows in plain init (no assignments)
        assert len(group["wp_values"]) == 0, (
            f"{name}: expected no WP rows in plain init, "
            f"got {set(group['wp_values'].keys())}")
    wb.close()


# ── discover_layout validation (malformed Plan sheets) ──


class TestDiscoverLayoutValidation:
    """discover_layout rejects malformed Plan sheets with clear errors."""

    def test_missing_planned_header(self):
        """Raises ValueError when PLANNED block header is missing."""
        wb = Workbook()
        ws = wb.active
        assert ws is not None
        ws.title = "Plan"
        ws.cell(1, 1, "Title:")
        ws.cell(1, 2, "Test Project")
        with pytest.raises(ValueError, match="PLANNED FTE"):
            discover_layout(ws)
        wb.close()

    def test_missing_wp_data_rows_in_planned_block(self):
        """Raises ValueError when PLANNED block has no WP data rows."""
        wb = Workbook()
        ws = wb.active
        assert ws is not None
        ws.title = "Plan"
        # Compact header
        ws.cell(1, 1, "Title:")
        # PLANNED header present but no WP data follows
        ws.cell(3, 1, "PLANNED FTE (PERSON-MONTHS)")
        ws.cell(4, 1, "")  # year row (empty)
        ws.cell(5, 1, "WP ID")  # month header row
        # No WP formula rows follow
        with pytest.raises(ValueError, match="Cannot find first WP data row"):
            discover_layout(ws)
        wb.close()

    def test_no_wp_budget_section_raises(self):
        """Plan has literal WP IDs (no =Config!A{n} formula) AND Config is
        missing the WP BUDGET section → discover_layout must raise instead of
        silently falling back to row 8, which would write WP data to the wrong
        Config rows on rebuild (T3)."""
        wb = Workbook()
        plan_ws = wb.active
        assert plan_ws is not None
        plan_ws.title = "Plan"
        plan_ws.cell(1, 1, "Title:")
        plan_ws.cell(3, 1, "PLANNED FTE (PERSON-MONTHS)")
        plan_ws.cell(4, 1, "")
        plan_ws.cell(5, 1, "WP ID")
        # Literal WP ID (not a =Config!A{n} formula) — exercises the fallback
        plan_ws.cell(6, 1, "WP1")
        plan_ws.cell(6, 2, "Package 1")
        # Config sheet exists but has no WP BUDGET section header
        wb.create_sheet("Config")
        with pytest.raises(ValueError, match="WP Budget section"):
            discover_layout(plan_ws)
        wb.close()

    def test_discover_layout_finds_wp_formula_rows(self, fresh_wb_path):
        """discover_layout correctly finds WP data rows via formula scan."""
        import openpyxl
        wb = openpyxl.load_workbook(fresh_wb_path)
        layout = discover_layout(wb["Plan"])
        num_wps = len(WORK_PACKAGES)
        # Verify the layout is internally consistent
        assert layout["planned_last_wp_row"] == layout["planned_first_wp_row"] + num_wps - 1
        assert layout["planned_total_row"] == layout["planned_last_wp_row"] + 1
        # Verify first WP row has a formula in col A
        first_val = wb["Plan"].cell(layout["planned_first_wp_row"], 1).value
        assert str(first_val).startswith("="), (
            f"Expected formula at planned first WP row, got {first_val!r}"
        )
        wb.close()


# ── T1: Formula rate round-trip ──


def _find_row_by_label(ws, label_prefix):
    """Return the 1-based row index of the first col-A cell whose value
    starts with ``label_prefix`` (case-insensitive, substring)."""
    for r in range(1, ws.max_row + 1):
        label = str(ws.cell(r, 1).value or "")
        if label_prefix.lower() in label.lower():
            return r
    raise AssertionError(f"Label '{label_prefix}' not found in column A")


def _set_cell(ws, row: int, col: int, value) -> None:
    """Set a cell value with pyright-safe type narrowing (skips MergedCell)."""
    from openpyxl.cell import Cell
    cell = ws.cell(row, col)
    assert isinstance(cell, Cell), (
        f"Expected regular Cell at ({row}, {col}); got {type(cell).__name__}"
    )
    cell.value = value


class TestFormulaRateRoundTrip:
    """reader.py preserves formula strings for WP rates and FTE Rate."""

    def test_formula_rate_roundtrip_fte_rate(self, tmp_path):
        """Setting FTE Rate cell to a formula string survives read."""
        path = str(tmp_path / "fte_rate_formula.xlsx")
        init_workbook(output_path=path)

        # Overwrite FTE Rate cell (col B on the row labelled "FTE Rate (EUR)")
        # with a simple formula and re-read.
        wb = load_workbook(path)
        cfg = wb["Config"]
        rate_row = _find_row_by_label(cfg, "FTE Rate")
        _set_cell(cfg, rate_row, 2, "=AVERAGE(B9:B11)")
        wb.save(path)
        wb.close()

        wb = load_workbook(path)
        project_params, _, _ = read_config_sheet(wb["Config"])
        wb.close()

        assert project_params["rate"] == "=AVERAGE(B9:B11)", (
            f"Formula lost on read; got {project_params['rate']!r}"
        )

    def test_formula_fte_rate_survives_full_rebuild(self, tmp_path):
        """A user-authored FTE Rate formula must survive a full rebuild —
        the Config builder must not clobber it with the default SUMPRODUCT."""
        path = str(tmp_path / "fte_rate_rebuild.xlsx")
        init_workbook(output_path=path)

        # User replaces the FTE Rate cell with a custom formula
        wb = load_workbook(path)
        cfg = wb["Config"]
        rate_row = _find_row_by_label(cfg, "FTE Rate")
        custom_formula = "=AVERAGE(B9:B11)"
        _set_cell(cfg, rate_row, 2, custom_formula)
        wb.save(path)
        wb.close()

        # Rebuild the workbook — the formula must be preserved end-to-end
        from budget_planner.workbook import rebuild_workbook
        rebuild_workbook(input_path=path)

        wb = load_workbook(path)
        rebuilt_val = wb["Config"].cell(rate_row, 2).value
        wb.close()
        assert rebuilt_val == custom_formula, (
            f"Rebuild clobbered user formula; got {rebuilt_val!r}"
        )

    def test_formula_rate_roundtrip_wp_rate(self, tmp_path):
        """Formula in per-WP Rate cell survives read AND is rewritten on save."""
        path = str(tmp_path / "wp_rate_formula.xlsx")
        init_workbook(output_path=path)

        # Overwrite the first WP's rate cell with a formula
        wb = load_workbook(path)
        cfg = wb["Config"]
        _, original_wps, _ = read_config_sheet(cfg)
        # Find the WP Budget first data row by locating "WP BUDGET" header
        wp_header_row = _find_row_by_label(cfg, "WP BUDGET")
        first_wp_row = wp_header_row + 2  # header + column-headers row
        formula = "=Config!$B$4*1.1"
        _set_cell(cfg, first_wp_row, COL_WP_RATE, formula)
        wb.save(path)
        wb.close()

        # Re-read: rate_raw should preserve formula, rate should fall back to default
        wb = load_workbook(path)
        _, wps_after, _ = read_config_sheet(wb["Config"])
        wb.close()

        assert len(wps_after) == len(original_wps)
        assert wps_after[0]["rate_raw"] == formula, (
            f"Expected rate_raw to preserve formula; got {wps_after[0]['rate_raw']!r}"
        )
        assert isinstance(wps_after[0]["rate"], (int, float)), (
            "numeric `rate` should remain numeric for downstream consumers"
        )

        # Now rewrite via the config-sheet builder and confirm the formula
        # propagates back to the cell.
        from budget_planner.config_sheet import build_config_sheet
        out_wb = Workbook()
        out_ws = out_wb.active
        assert out_ws is not None
        out_ws.title = "Config"
        params_sample = {
            "title": "Rebuild Test",
            "start": __import__("datetime").datetime(2026, 1, 1),
            "months": 60,
            "rate": 7500,
            "eu_ceiling_pm": 12.0,
        }
        build_config_sheet(out_ws, out_wb, params_sample, wps_after)
        # Find the rate cell on the rebuilt sheet
        rebuilt_wp_header = _find_row_by_label(out_ws, "WP BUDGET")
        rebuilt_first_wp = rebuilt_wp_header + 2
        rebuilt_rate_val = out_ws.cell(rebuilt_first_wp, COL_WP_RATE).value
        assert rebuilt_rate_val == formula, (
            f"Builder lost formula on rewrite; got {rebuilt_rate_val!r}"
        )

    def test_numeric_rate_still_works(self, tmp_path):
        """Baseline: a numeric per-WP rate round-trips as a float."""
        path = str(tmp_path / "wp_rate_numeric.xlsx")
        init_workbook(output_path=path)

        wb = load_workbook(path)
        cfg = wb["Config"]
        wp_header_row = _find_row_by_label(cfg, "WP BUDGET")
        first_wp_row = wp_header_row + 2
        _set_cell(cfg, first_wp_row, COL_WP_RATE, 7500)
        wb.save(path)
        wb.close()

        wb = load_workbook(path)
        _, wps_after, _ = read_config_sheet(wb["Config"])
        wb.close()

        assert wps_after[0]["rate"] == pytest.approx(7500.0)
        # rate_raw captures the literal cell value (here: numeric)
        assert wps_after[0]["rate_raw"] == 7500


# ── T2: Anchored-label section header contract ──


class TestExactSectionHeaderMatching:
    """reader._find_header_row is exact-match, not substring — and
    layout drift (renamed sections, inserted blank rows) is detected loudly."""

    def test_exact_header_match_required(self, tmp_path):
        """Renaming the 'WP BUDGET' cell to a substring-containing label
        no longer matches; read_config_sheet raises ValueError with the
        canonical label in the message."""
        path = str(tmp_path / "exact_match.xlsx")
        init_workbook(output_path=path)

        wb = load_workbook(path)
        cfg = wb["Config"]
        wp_budget_row = _find_row_by_label(cfg, "WP BUDGET")
        # Rename to something that still contains "WP BUDGET" as a substring
        # (under the old substring matcher this would have matched).
        _set_cell(cfg, wp_budget_row, 1, "Legacy WP Budget (archived)")
        wb.save(path)
        wb.close()

        wb = load_workbook(path)
        with pytest.raises(ValueError, match="WP BUDGET"):
            read_config_sheet(wb["Config"])
        wb.close()

    def test_blank_row_between_header_and_data_detected(self, tmp_path):
        """Inserting a blank row between the 'WP BUDGET' title row and the
        data rows must NOT silently succeed.

        Under the old substring-match + hardcoded +2 offset, the reader would
        start scanning at title_row+2 — which after a blank insert lands on
        the original column-header row. It would then treat 'WP ID' as a WP
        identifier and call ``float('FTE Budget')`` — a surprising ValueError
        whose message points at parsing, not at layout drift. This test
        documents that a ValueError IS raised (so drift is not silently
        swallowed), which is the current, cleanly-observable outcome."""
        path = str(tmp_path / "blank_row_drift.xlsx")
        init_workbook(output_path=path)

        wb = load_workbook(path)
        cfg = wb["Config"]
        wp_budget_row = _find_row_by_label(cfg, "WP BUDGET")
        # Insert a blank row right after the section title row. This pushes
        # the column-header row (WP ID, WP Name, …) down by one, so the
        # reader's data-scan now starts on the column-header row.
        cfg.insert_rows(wp_budget_row + 1)
        wb.save(path)
        wb.close()

        wb = load_workbook(path)
        # Reader should raise — either because float('FTE Budget') fails, or
        # (better, once the reader is layout-aware) with an explicit drift
        # message. Either way: NOT a silent success.
        with pytest.raises((ValueError, TypeError)):
            read_config_sheet(wb["Config"])
        wb.close()

    def test_section_header_labels_covers_all_sections(self):
        """Every label in SECTION_HEADER_LABELS must appear as a literal
        string in some builder module, preventing drift between the
        canonical label dict and what the builders actually emit."""
        builder_dir = pathlib.Path(__file__).resolve().parent.parent / "budget_planner"
        builder_sources = {
            p.name: p.read_text()
            for p in builder_dir.glob("*.py")
            if p.name not in {"reader.py", "__init__.py", "__main__.py"}
        }
        for key, label in SECTION_HEADER_LABELS.items():
            # The literal label string must appear in at least one non-reader
            # module (builders, layout helper, or workbook orchestrator).
            quoted_variants = (f'"{label}"', f"'{label}'")
            found_in = [
                name for name, src in builder_sources.items()
                if any(q in src for q in quoted_variants)
                or f'SECTION_HEADER_LABELS["{key}"]' in src
            ]
            assert found_in, (
                f"Label {key!r}={label!r} is declared in SECTION_HEADER_LABELS "
                f"but no builder module references it (searched "
                f"{sorted(builder_sources)})"
            )
