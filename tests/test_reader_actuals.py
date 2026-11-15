"""Reader tests for the Actuals sheet.

Split from tests/test_reader_capacity.py (Phase 2 mirror-split).

Covers:
* TestActualsReaderValidation — read_actuals_values rejects malformed sheets.
* read_actuals_values behavioral contract — formula-ref resolution, blank-row
  preservation, literal WP IDs, TOTAL row stop.
* TestFormulaRateRoundTrip — per-WP rate formula survives read AND rewrite.
* TestExactSectionHeaderMatching — anchored-label section-header contract.

Out-of-scope and lives in sibling files:
* ``tests/test_reader_capacity.py`` — capacity-reader tests.
* ``tests/test_reader_config.py`` — Config-section readers.
* ``tests/test_reader_layout_discovery.py`` — discover_layout validation.
* ``tests/test_reader.py`` — legacy ``GRANT_ACCEPT_CEILING_MIGRATION``
  env-guarded migration paths.
* ``tests/test_reader_team_priority.py`` — TEAM ROSTER Priority field
  round-trip + clamping.
"""

import pathlib

import pytest
from openpyxl import Workbook, load_workbook

from budget_planner.config import PROJECT_START
from budget_planner.config_sheet import build_config_sheet
from budget_planner.layout import (
    COL_WP_RATE,
    SECTION_HEADER_LABELS,
)
from budget_planner.reader import read_config_sheet
from budget_planner.workbook import init_workbook


# ── Actuals reader validation (malformed workbooks) ──


class TestActualsReaderValidation:
    """read_actuals_values rejects malformed Actuals sheets."""

    def test_missing_wp_id_header(self):
        """Raises ValueError when Actuals sheet has no 'WP ID' header."""
        from budget_planner.reader import read_actuals_values

        wb = Workbook()
        plan_ws = wb.active
        assert plan_ws is not None
        plan_ws.title = "Plan"

        actuals_ws = wb.create_sheet("Actuals")
        actuals_ws.cell(1, 1, "ACTUAL EUR")
        actuals_ws.cell(2, 1, "")
        # Row 3 is missing "WP ID" — has something else instead
        actuals_ws.cell(3, 1, "Identifier")
        actuals_ws.cell(3, 2, "Name")

        with pytest.raises(ValueError, match="Cannot find 'WP ID' header row"):
            read_actuals_values(actuals_ws, project_months=3)
        wb.close()

    def test_wp_id_header_at_nonstandard_row(self):
        """Actuals reader finds 'WP ID' header even if not at row 3."""
        from budget_planner.reader import read_actuals_values

        wb = Workbook()
        plan_ws = wb.active
        assert plan_ws is not None
        plan_ws.title = "Plan"
        plan_ws.cell(6, 1, "WP1")

        actuals_ws = wb.create_sheet("Actuals")
        actuals_ws.cell(1, 1, "ACTUAL EUR")
        actuals_ws.cell(2, 1, "")
        actuals_ws.cell(3, 1, "")  # extra blank row
        actuals_ws.cell(4, 1, "")  # another blank row
        # WP ID header shifted to row 5 (non-standard)
        actuals_ws.cell(5, 1, "WP ID")
        actuals_ws.cell(5, 2, "WP Name")
        # WP data at row 6
        actuals_ws.cell(6, 1, "=Plan!A6")
        actuals_ws.cell(6, 2, "=Plan!B6")
        actuals_ws.cell(6, 3, 42.0)
        # TOTAL row
        actuals_ws.cell(7, 1, "")
        actuals_ws.cell(7, 2, "TOTAL")

        result = read_actuals_values(actuals_ws, project_months=1)
        assert "WP1" in result
        assert result["WP1"] == [42.0]
        wb.close()


# ── read_actuals_values behavioral contract (restored 2026-04-26 after
#    Phase 2 over-pruning; code-review found T8 blank-row preservation,
#    =Plan!A{n} formula resolution, literal-WP-ID handling, and zero-row
#    preservation each had no surviving test) ──


def _make_actuals_wb():
    """Minimal workbook with Plan + Actuals sheets for read_actuals_values."""
    wb = Workbook()
    plan_ws = wb.active
    assert plan_ws is not None
    plan_ws.title = "Plan"
    plan_ws.cell(6, 1, "WP1")
    plan_ws.cell(7, 1, "WP2")

    actuals_ws = wb.create_sheet("Actuals")
    actuals_ws.cell(1, 1, "ACTUAL EUR")
    actuals_ws.cell(2, 1, "")
    actuals_ws.cell(3, 1, "WP ID")
    actuals_ws.cell(3, 2, "WP Name")
    actuals_ws.cell(4, 1, "=Plan!A6")
    actuals_ws.cell(4, 2, "=Plan!B6")
    actuals_ws.cell(4, 3, 100.0)
    actuals_ws.cell(4, 4, 200.0)
    actuals_ws.cell(4, 5, 0)
    actuals_ws.cell(5, 1, "=Plan!A7")
    actuals_ws.cell(5, 2, "=Plan!B7")
    actuals_ws.cell(5, 3, 0)
    actuals_ws.cell(5, 4, 0)
    actuals_ws.cell(5, 5, 0)
    actuals_ws.cell(6, 1, "")
    actuals_ws.cell(6, 2, "TOTAL")
    return wb


def test_read_actuals_values_resolves_formula_refs():
    """read_actuals_values resolves =Plan!A{n} formulas to WP IDs.

    Distinct codepath from the Config!A{n} refs the prod builder writes —
    integration tests in test_rebuild.py exercise the Config!-ref branch,
    so this is the only guard for the Plan!-ref branch in
    readers/layout_discovery.py::_resolve_plan_ref (called from
    readers/actuals.py::read_actuals_values).
    """
    from budget_planner.reader import read_actuals_values

    wb = _make_actuals_wb()
    result = read_actuals_values(wb["Actuals"], project_months=3)
    assert "WP1" in result
    assert result["WP1"] == [100.0, 200.0, 0.0]
    wb.close()


def test_read_actuals_values_keeps_all_zero_rows():
    """Rows with explicit numeric zero (not blank) survive the round-trip.

    Distinct from test_read_actuals_values_preserves_all_blank_row, which
    covers the all-None case introduced by T8.
    """
    from budget_planner.reader import read_actuals_values

    wb = _make_actuals_wb()
    result = read_actuals_values(wb["Actuals"], project_months=3)
    assert "WP2" in result
    assert result["WP2"] == [0.0, 0.0, 0.0]
    wb.close()


def test_read_actuals_values_stops_at_total():
    """read_actuals_values stops scanning at the TOTAL row."""
    from budget_planner.reader import read_actuals_values

    wb = _make_actuals_wb()
    result = read_actuals_values(wb["Actuals"], project_months=3)
    assert len(result) == 2
    wb.close()


def test_read_actuals_values_preserves_all_blank_row():
    """T8: WP row with all-None monthly cells is still preserved.

    Previously the reader dropped rows whose monthly values were every None
    ("never entered"), conflating them with untouched WPs. The round-trip
    contract is now: every Actuals WP row is preserved, blank cells included.
    """
    from budget_planner.reader import read_actuals_values

    wb = Workbook()
    plan_ws = wb.active
    assert plan_ws is not None
    plan_ws.title = "Plan"
    plan_ws.cell(6, 1, "WP1")
    plan_ws.cell(7, 1, "WP2")

    actuals_ws = wb.create_sheet("Actuals")
    actuals_ws.cell(1, 1, "ACTUAL EUR")
    actuals_ws.cell(3, 1, "WP ID")
    actuals_ws.cell(3, 2, "WP Name")
    actuals_ws.cell(4, 1, "=Plan!A6")
    actuals_ws.cell(4, 2, "=Plan!B6")
    actuals_ws.cell(4, 3, 100.0)
    actuals_ws.cell(4, 4, 200.0)
    actuals_ws.cell(4, 5, 300.0)
    actuals_ws.cell(5, 1, "=Plan!A7")
    actuals_ws.cell(5, 2, "=Plan!B7")
    actuals_ws.cell(6, 1, "")
    actuals_ws.cell(6, 2, "TOTAL")

    result = read_actuals_values(actuals_ws, project_months=3)
    assert "WP2" in result, (
        "WP2 row was dropped even though it appears in the Actuals table. "
        "Zero-spend / deliberately-cleared rows must survive the round-trip."
    )
    assert result["WP2"] == [None, None, None]
    wb.close()


def test_read_actuals_values_literal_wp_ids():
    """read_actuals_values handles literal (non-formula) WP IDs.

    Covers the early-return branch in readers/layout_discovery.py::
    _resolve_plan_ref for cells whose value does not start with '='. The prod
    builder always writes formula refs, so this branch has no integration
    coverage.
    """
    from budget_planner.reader import read_actuals_values

    wb = Workbook()
    plan_ws = wb.active
    assert plan_ws is not None
    plan_ws.title = "Plan"

    actuals_ws = wb.create_sheet("Actuals")
    actuals_ws.cell(1, 1, "ACTUAL EUR")
    actuals_ws.cell(2, 1, "")
    actuals_ws.cell(3, 1, "WP ID")
    actuals_ws.cell(3, 2, "WP Name")
    actuals_ws.cell(4, 1, "WP1")
    actuals_ws.cell(4, 2, "Mgmt")
    actuals_ws.cell(4, 3, 500.0)
    actuals_ws.cell(4, 4, 600.0)
    actuals_ws.cell(5, 1, "")
    actuals_ws.cell(5, 2, "TOTAL")

    result = read_actuals_values(actuals_ws, project_months=2)
    assert "WP1" in result
    assert result["WP1"] == [500.0, 600.0]
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
    assert isinstance(cell, Cell), f"Expected regular Cell at ({row}, {col}); got {type(cell).__name__}"
    cell.value = value


class TestFormulaRateRoundTrip:
    """readers/config.py preserves formula strings for WP rates.

    FTE Rate (B4) is a dashboard readout — the builder always regenerates
    it from the live grade/roster layout, so custom B4 formulas are
    intentionally not preserved. See config_sheet.build_config_sheet.
    """

    def test_formula_rate_roundtrip_wp_rate(self, tmp_path):
        """Formula in per-WP Rate cell survives read AND is rewritten on save."""
        path = str(tmp_path / "wp_rate_formula.xlsx")
        init_workbook(output_path=path)

        # Overwrite the first WP's rate cell with a formula
        wb = load_workbook(path)
        cfg = wb["Config"]
        _, original_wps = read_config_sheet(cfg)
        # Find the WP Budget first data row by locating "WP BUDGET" header
        wp_header_row = _find_row_by_label(cfg, "WP BUDGET")
        first_wp_row = wp_header_row + 2  # header + column-headers row
        formula = "=Config!$B$4*1.1"
        _set_cell(cfg, first_wp_row, COL_WP_RATE, formula)
        wb.save(path)
        wb.close()

        # Re-read: rate_raw should preserve formula, rate should fall back to default
        wb = load_workbook(path)
        _, wps_after = read_config_sheet(wb["Config"])
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
        out_wb = Workbook()
        out_ws = out_wb.active
        assert out_ws is not None
        out_ws.title = "Config"
        params_sample = {
            "title": "Rebuild Test",
            "start": PROJECT_START,
            "months": 60,
            "rate": 10000,
            "eu_ceiling_pm": 12.0,
        }
        build_config_sheet(out_ws, out_wb, params_sample, wps_after)
        # Find the rate cell on the rebuilt sheet
        rebuilt_wp_header = _find_row_by_label(out_ws, "WP BUDGET")
        rebuilt_first_wp = rebuilt_wp_header + 2
        rebuilt_rate_val = out_ws.cell(rebuilt_first_wp, COL_WP_RATE).value
        assert rebuilt_rate_val == formula, f"Builder lost formula on rewrite; got {rebuilt_rate_val!r}"

    def test_formula_wp_rate_roundtrip(self, tmp_path):
        """End-to-end: per-WP Rate formula survives init → edit → rebuild."""
        path = str(tmp_path / "wp_rate_roundtrip.xlsx")
        init_workbook(output_path=path)

        wb = load_workbook(path)
        cfg = wb["Config"]
        wp_header_row = _find_row_by_label(cfg, "WP BUDGET")
        first_wp_row = wp_header_row + 2
        formula = "=B2*1.1"
        _set_cell(cfg, first_wp_row, COL_WP_RATE, formula)
        wb.save(path)
        wb.close()

        # Full rebuild must preserve the per-WP formula in its Config cell
        from budget_planner.workbook import rebuild_workbook

        rebuild_workbook(input_path=path)

        wb = load_workbook(path)
        rebuilt_val = wb["Config"].cell(first_wp_row, COL_WP_RATE).value
        wb.close()
        assert rebuilt_val == formula, f"Rebuild clobbered user WP Rate formula; got {rebuilt_val!r}"

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
        _, wps_after = read_config_sheet(wb["Config"])
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
        # Top-level *.py only — readers/ is the parser side, never emits headers.
        # The non-recursive glob already skips readers/; the explicit exclusion
        # also skips reader.py (back-compat re-export shim, no literals) and
        # __init__/__main__ (package wiring, no literals).
        builder_sources = {
            p.name: p.read_text()
            for p in builder_dir.glob("*.py")
            if p.name not in {"reader.py", "__init__.py", "__main__.py"}
        }
        for key, label in SECTION_HEADER_LABELS.items():
            # The literal label string must appear in at least one builder
            # module (builders, layout helper, or workbook orchestrator).
            quoted_variants = (f'"{label}"', f"'{label}'")
            found_in = [
                name
                for name, src in builder_sources.items()
                if any(q in src for q in quoted_variants) or f'SECTION_HEADER_LABELS["{key}"]' in src
            ]
            assert found_in, (
                f"Label {key!r}={label!r} is declared in SECTION_HEADER_LABELS "
                f"but no builder module references it (searched "
                f"{sorted(builder_sources)})"
            )
