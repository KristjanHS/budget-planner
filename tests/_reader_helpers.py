"""Shared synthetic-Config builders for tests/test_reader_*.py."""

from openpyxl import Workbook

from budget_planner.config import PROJECT_START
from budget_planner.config_sheet import build_config_sheet


def _make_config_wb():
    """Create a workbook with a Config sheet for testing."""
    wb = Workbook()
    ws = wb.active
    assert ws is not None
    ws.title = "Config"
    params = {
        "title": "Test",
        "start": PROJECT_START,
        "months": 24,
        "rate": 10000,
    }
    wps = [
        {"id": "WP1", "name": "Mgmt", "fte": 12.0, "start": 1, "end": 24, "type": "flat"},
        {"id": "WP2", "name": "Dev", "fte": 8.0, "start": 1, "end": 12, "type": "flat"},
    ]
    build_config_sheet(ws, wb, params, wps)
    return wb


def _sample_assignments():
    return [
        {"person": "Alice", "wp": "WP1", "total_fte": 4.5, "start": 1, "end": 24},
        {"person": "Bob", "wp": "WP2", "total_fte": 3.0, "start": 1, "end": 12},
    ]


def _sample_externals():
    return [
        {
            "person": "Alice",
            "project": "ProjectX",
            "fte": 2.0,
            "start": 1,
            "end": 24,
            "is_eu": False,
        },
        {
            "person": "Bob",
            "project": "HorizonEU",
            "fte": 4.0,
            "start": 1,
            "end": 24,
            "is_eu": True,
        },
    ]


def _sample_team():
    return [
        {"name": "Alice", "max_fte": 1.0, "grade": "Senior"},
        {"name": "Bob", "max_fte": 1.0, "grade": "Mid"},
    ]


def _sample_grade_rates():
    return {"Senior": 12500, "Mid": 10000}


def _make_config_wb_with_assignments(assignments=None, externals=None):
    """Create a workbook with a Config sheet that has assignments/externals."""
    wb = Workbook()
    ws = wb.active
    assert ws is not None
    ws.title = "Config"
    params = {
        "title": "Test",
        "start": PROJECT_START,
        "months": 60,
        "rate": 10000,
        "eu_ceiling_pm": 12.0,
    }
    wps = [
        {"id": "WP1", "name": "Mgmt", "fte": 12.0, "start": 1, "end": 60, "type": "flat"},
        {"id": "WP2", "name": "Dev", "fte": 8.0, "start": 1, "end": 60, "type": "flat"},
    ]
    build_config_sheet(
        ws,
        wb,
        params,
        wps,
        team=_sample_team(),
        grade_rates=_sample_grade_rates(),
        assignments=assignments or [],
        external_commitments=externals or [],
    )
    return wb
