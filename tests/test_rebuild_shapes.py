"""Parameterized rebuild round-trip tests varying one Config-table shape at a time.

Mirrors the user's Excel workflow: generate a workbook, insert one row mid-section
(via openpyxl, with named-range auto-extend emulated where Excel would do so),
rebuild, and assert the inserted data survives.

Variants covered here:
- extra_wp: new row in WP BUDGET
- extra_external: new row in EXTERNAL COMMITMENTS
- extra_grade: new row in RATE BY GRADE

(``extra_person`` is covered by ``test_rebuild_shape_drift.py``;
``default`` — init → rebuild with no edits — is a no-op implicitly
exercised by every other rebuild test.)
"""
from pathlib import Path

import openpyxl
import pytest

from budget_planner.reader import (
    read_assignments_from_config,
    read_config_sheet,
    read_externals_from_config,
    read_grade_rates,
    read_team_from_config,
)
from budget_planner.workbook import init_workbook, rebuild_workbook
from tests.test_rebuild_shape_drift import _simulate_excel_insert_in_named_range


SECTION_LABELS = {
    "wp_budget": "WP BUDGET",
    "externals": "EXTERNAL COMMITMENTS",
    "rate_by_grade": "RATE BY GRADE",
}


def _first_data_row_after(ws, section_label: str) -> int:
    """Return the first data row for a section (header row + 2)."""
    for r in range(1, ws.max_row + 1):
        if str(ws.cell(r, 1).value or "").strip().upper() == section_label:
            return r + 2
    raise AssertionError(f"Section {section_label!r} not found")


@pytest.mark.parametrize(
    "variant",
    ["extra_wp", "extra_external", "extra_grade"],
)
def test_rebuild_roundtrip_shape_variant(tmp_path: Path, variant: str):
    path = tmp_path / f"wb_{variant}.xlsx"
    init_workbook(str(path))

    wb = openpyxl.load_workbook(path)
    config_ws = wb["Config"]

    if variant == "extra_wp":
        data_start = _first_data_row_after(config_ws, SECTION_LABELS["wp_budget"])
        insert_at = data_start + 1
        config_ws.insert_rows(insert_at, amount=1)
        config_ws.cell(insert_at, 1, "WPX")
        config_ws.cell(insert_at, 2, "Extra WP")
        config_ws.cell(insert_at, 3, 1.0)
        config_ws.cell(insert_at, 4, 6500)
    elif variant == "extra_external":
        data_start = _first_data_row_after(config_ws, SECTION_LABELS["externals"])
        insert_at = data_start + 1
        config_ws.insert_rows(insert_at, amount=1)
        config_ws.cell(insert_at, 1, "Alice")
        config_ws.cell(insert_at, 2, "Other Grant")
        config_ws.cell(insert_at, 3, 0.2)
        config_ws.cell(insert_at, 4, 1)
        config_ws.cell(insert_at, 5, 24)
        config_ws.cell(insert_at, 6, "Yes")
        for nr in ("ext_persons", "ext_ftes"):
            _simulate_excel_insert_in_named_range(wb, nr, insert_at, 1)
    elif variant == "extra_grade":
        data_start = _first_data_row_after(config_ws, SECTION_LABELS["rate_by_grade"])
        insert_at = data_start + 1
        config_ws.insert_rows(insert_at, amount=1)
        config_ws.cell(insert_at, 1, "Principal")
        config_ws.cell(insert_at, 2, 9500)
        _simulate_excel_insert_in_named_range(wb, "tbl_GradeRates", insert_at, 1)

    wb.save(path)
    wb.close()

    rebuild_workbook(str(path))

    wb = openpyxl.load_workbook(path)
    config_ws = wb["Config"]

    if variant == "extra_wp":
        _, wps, _ = read_config_sheet(config_ws)
        ids = [w["id"] for w in wps]
        assert "WPX" in ids, f"Rebuild dropped inserted WP. Got {ids}"
    elif variant == "extra_external":
        projects = [e["project"] for e in read_externals_from_config(config_ws)]
        assert "Other Grant" in projects, (
            f"Rebuild dropped inserted external commitment. Got {projects}"
        )
    elif variant == "extra_grade":
        grades = read_grade_rates(config_ws)
        assert "Principal" in grades, (
            f"Rebuild dropped inserted grade. Got {list(grades)}"
        )

    # Universal invariant: every assignment points at a live roster name.
    roster_names = {p["name"] for p in read_team_from_config(config_ws)}
    orphaned = {
        a["person"] for a in read_assignments_from_config(config_ws)
    } - roster_names
    assert not orphaned, f"Orphaned assignments after rebuild: {orphaned}"
    wb.close()
