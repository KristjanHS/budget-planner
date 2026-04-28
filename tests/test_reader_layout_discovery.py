"""Reader tests for discover_layout (Plan-sheet layout discovery).

Split from tests/test_reader_capacity.py (Phase 2 mirror-split).

Covers:
* TestDiscoverLayoutValidation — discover_layout rejects malformed Plan sheets.
* _iter_named_range_rows — internal helper for named-range row spans.

Out-of-scope and lives in sibling files:
* ``tests/test_reader_capacity.py`` — capacity-reader tests.
* ``tests/test_reader_config.py`` — Config-section readers (team, assignments,
  externals, reporting periods, grade rates).
* ``tests/test_reader_actuals.py`` — Actuals reader + formula-rate round-trip
  + exact section-header matching.
* ``tests/test_reader.py`` — legacy ``GRANT_ACCEPT_CEILING_MIGRATION``
  env-guarded migration paths.
* ``tests/test_reader_team_priority.py`` — TEAM ROSTER Priority field
  round-trip + clamping.
"""
import pytest
from openpyxl import Workbook

from budget_planner.config import WORK_PACKAGES
from budget_planner.reader import _iter_named_range_rows, discover_layout


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


# ── _iter_named_range_rows helper ──


def test_iter_named_range_rows_returns_live_row_span_after_middle_insert():
    """Helper must report the post-insertion row span via Excel's defined_names."""
    from openpyxl.workbook.defined_name import DefinedName

    wb = Workbook()
    ws = wb.active
    assert ws is not None
    ws.title = "Config"
    for i, val in enumerate(["A", "B", "C"], start=5):
        ws.cell(i, 1, value=val)
    wb.defined_names["roster_ftes"] = DefinedName(
        name="roster_ftes", attr_text="Config!$B$5:$B$7"
    )

    # Excel-style middle insert: shift rows 6+ down by 1.
    ws.insert_rows(6, amount=1)
    ws.cell(6, 1, value="MID")
    # openpyxl does NOT update DefinedName targets on insert — simulate Excel
    # by explicitly rewriting the range. The helper under test treats the
    # defined_names entry as the source of truth.
    wb.defined_names["roster_ftes"].attr_text = "Config!$B$5:$B$8"

    assert _iter_named_range_rows(wb, "roster_ftes") == (5, 8)


def test_iter_named_range_rows_handles_single_cell_ref():
    """Single-cell absolute ref (e.g. $B$5) must return (5, 5)."""
    from openpyxl.workbook.defined_name import DefinedName

    wb = Workbook()
    ws = wb.active
    assert ws is not None
    ws.title = "Config"
    wb.defined_names["project_start"] = DefinedName(
        name="project_start", attr_text="Config!$B$5"
    )
    assert _iter_named_range_rows(wb, "project_start") == (5, 5)


def test_iter_named_range_rows_returns_none_for_unknown_name():
    wb = Workbook()
    assert _iter_named_range_rows(wb, "does_not_exist") is None
