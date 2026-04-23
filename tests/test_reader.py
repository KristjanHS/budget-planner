"""Tests for reader behaviour — indirect-view Ceiling→Budget migration."""
from datetime import datetime

import pytest
from openpyxl import Workbook

from budget_planner.config_sheet import build_config_sheet
from budget_planner.reader import read_config_sheet


def _sample_params():
    return {
        "title": "Reader Test",
        "start": datetime(2026, 7, 1),
        "months": 60,
        "rate": 7500,
        "eu_ceiling_pm": 12.0,
    }


def _sample_wps():
    return [
        {"id": "WP1", "name": "Mgmt", "fte": 12.0, "start": 1, "end": 60},
        {"id": "WP2", "name": "Design", "fte": 8.0, "start": 1, "end": 18},
    ]


def _build_fresh():
    wb = Workbook()
    ws = wb.active
    assert ws is not None
    ws.title = "Config"
    build_config_sheet(ws, wb, _sample_params(), _sample_wps())
    return wb, ws


def _seed_legacy_config(ws, wps):
    """Lay out a pre-restructure Config sheet: params at rows 1-9, legacy
    ``Ceiling (€)`` at col 5, Start/End/Type at cols 8/9/10."""
    param_rows = (
        ("Project Title", "Legacy Project"),
        ("Project Start", datetime(2026, 7, 1)),
        ("Project Months", 60),
        ("FTE Rate (EUR)", 7500),
        ("EU Annual Ceiling (PM/FTE/year)", 12.0),
        ("Hours per FTE Month", 143),
        ("Vacation Month", "Jul"),
        ("Num Vacation Months", 5),
        ("Version Stamp", "v0-legacy"),
    )
    for r, (label, value) in enumerate(param_rows, start=1):
        ws.cell(row=r, column=1, value=label)
        ws.cell(row=r, column=2, value=value)

    section_row = 11
    header_row = 12
    first_data_row = 13
    ws.cell(row=section_row, column=1, value="WP BUDGET")
    headers = ["WP ID", "WP Name", "FTE Budget", "Rate (€/PM)", "Ceiling (€)",
               "Assigned PM", "Variance", "Start", "End", "Type"]
    for col, hdr in enumerate(headers, start=1):
        ws.cell(row=header_row, column=col, value=hdr)

    for i, (wp_id, name, fte, rate, start, end, wp_type) in enumerate(wps):
        r = first_data_row + i
        ws.cell(row=r, column=1, value=wp_id)
        ws.cell(row=r, column=2, value=name)
        ws.cell(row=r, column=3, value=fte)
        ws.cell(row=r, column=4, value=rate)
        ws.cell(row=r, column=5, value=fte * rate)  # legacy Ceiling = FTE × Rate
        ws.cell(row=r, column=8, value=start)
        ws.cell(row=r, column=9, value=end)
        ws.cell(row=r, column=10, value=wp_type)


def test_post_restructure_reader_reads_budget_plus_i_directly():
    """A freshly-built workbook has "Budget +i" header and blue-input values;
    reader parses ``budget_plus_i`` from col F without triggering migration."""
    _wb, ws = _build_fresh()
    _params, wps = read_config_sheet(ws)
    # Freshly built blue-input value seeded from fte × rate × 1.25.
    by_id = {wp["id"]: wp for wp in wps}
    assert by_id["WP1"]["budget_plus_i"] == 112_500
    assert by_id["WP2"]["budget_plus_i"] == 75_000


def test_legacy_ceiling_layout_migrates_to_budget_plus_i():
    """Pre-restructure workbook with header "Ceiling (€)" at col 5 and no col 6
    Budget triggers migration: Budget +i = FTE × Rate × 1.25."""
    wb = Workbook()
    ws = wb.active
    assert ws is not None
    ws.title = "Config"
    _seed_legacy_config(
        ws,
        [
            ("WP1", "Mgmt",   12.0, 7500, 1, 60, "flat"),
            ("WP2", "Design",  8.0, 7500, 1, 18, "flat"),
        ],
    )
    with pytest.warns(RuntimeWarning, match="Ceiling value"):
        _params, parsed = read_config_sheet(ws)
    by_id = {wp["id"]: wp for wp in parsed}
    assert by_id["WP1"]["budget_plus_i"] == 112_500  # 12 × 7500 × 1.25
    assert by_id["WP2"]["budget_plus_i"] == 75_000   #  8 × 7500 × 1.25


def test_legacy_migration_sanity_guard_blocks_oversized_values(monkeypatch):
    """If legacy migration would seed any Budget +i > €3M, the reader refuses
    without the GRANT_ACCEPT_CEILING_MIGRATION env flag."""
    monkeypatch.delenv("GRANT_ACCEPT_CEILING_MIGRATION", raising=False)
    wb = Workbook()
    ws = wb.active
    assert ws is not None
    ws.title = "Config"
    # 350 × 7500 × 1.25 = 3,281,250 → over €3M cap
    _seed_legacy_config(ws, [("WP1", "Huge", 350.0, 7500, 1, 60, "flat")])
    with pytest.raises(RuntimeError, match="GRANT_ACCEPT_CEILING_MIGRATION"):
        read_config_sheet(ws)


def test_legacy_migration_sanity_guard_bypassed_with_env_flag(monkeypatch):
    """Oversized Budget +i migration proceeds when GRANT_ACCEPT_CEILING_MIGRATION is set."""
    monkeypatch.setenv("GRANT_ACCEPT_CEILING_MIGRATION", "1")
    wb = Workbook()
    ws = wb.active
    assert ws is not None
    ws.title = "Config"
    _seed_legacy_config(ws, [("WP1", "Huge", 350.0, 7500, 1, 60, "flat")])
    with pytest.warns(RuntimeWarning, match="Ceiling value"):
        _params, parsed = read_config_sheet(ws)
    assert parsed[0]["budget_plus_i"] == 3_281_250
