"""Overlay loader: the overlay file in BUDGET_PLANNER_OVERLAY_DIR overrides config.py placeholders."""

import json
import os
import subprocess
import sys
from pathlib import Path

import pytest

from budget_planner.overlay_settings import OVERLAY_DIR_ENV, load_overlay, overlay_dir_from_env

_DUMP = (
    "import json, budget_planner.config as c;"
    "print(json.dumps([c.PROJECT_START.isoformat(), c.PROJECT_MONTHS, c.GRADE_RATES['Mid'],"
    " c.WORK_PACKAGES[0]['name'], c.MONTH_COL_END - c.MONTH_COL_START + 1]))"
)


def _config_in_subprocess(overlay_dir):
    env = {**os.environ, OVERLAY_DIR_ENV: str(overlay_dir)}
    out = subprocess.run([sys.executable, "-c", _DUMP], env=env, capture_output=True, text=True, check=True)
    return json.loads(out.stdout)


def test_settings_env_overlays_config_constants(tmp_path):
    (tmp_path / "settings.env").write_text(
        "# private\nPROJECT_START=2024-03-01\nPROJECT_MONTHS=30\nGRADE_RATE_MID='1234'\nWP1_NAME=Real WP one\n"
    )
    start, months, mid, wp1, month_cols = _config_in_subprocess(tmp_path)
    assert (start, months, mid, wp1) == ("2024-03-01T00:00:00", 30, 1234, "Real WP one")
    assert month_cols == 30  # derived layout follows the overlaid length


def test_absent_companion_keeps_placeholders(tmp_path):
    from budget_planner import config

    start, months, mid, wp1, _ = _config_in_subprocess(tmp_path / "missing")
    assert (start, months, mid, wp1) == (
        config.PROJECT_START.isoformat(),
        config.PROJECT_MONTHS,
        config.GRADE_RATES["Mid"],
        config.WORK_PACKAGES[0]["name"],
    )


def test_unknown_key_is_rejected(tmp_path):
    (tmp_path / "settings.env").write_text("GRADE_RATE_MDI=1\n")
    with pytest.raises(ValueError, match="GRADE_RATE_MDI"):
        load_overlay(tmp_path)


def test_overlay_dir_env_resolution():
    assert overlay_dir_from_env({OVERLAY_DIR_ENV: ""}) is None
    assert overlay_dir_from_env({OVERLAY_DIR_ENV: "/x/y"}) == Path("/x/y")
    default = overlay_dir_from_env({})
    assert default is not None and default.name == "budget-planner-overlay"


def test_malformed_line_without_equals_is_rejected(tmp_path):
    (tmp_path / "settings.env").write_text("# header\nPROJECT_TITLE\n")
    with pytest.raises(ValueError, match=r"settings\.env:2: unknown or malformed setting 'PROJECT_TITLE'"):
        load_overlay(tmp_path)


@pytest.mark.parametrize(
    "line",
    [
        "PROJECT_MONTHS=0",
        "PROJECT_MONTHS=121",
        "PROJECT_MONTHS=abc",
        "RP_LENGTH=0",
        "DEFAULT_RATE=0",
        "GRADE_RATE_MID=-5",
        "PROJECT_START=2026-11-15",
        "PROJECT_START=Nov 2026",
    ],
)
def test_invalid_value_is_rejected_with_location(tmp_path, line):
    (tmp_path / "settings.env").write_text(f"# header\n{line}\n")
    key = line.partition("=")[0]
    with pytest.raises(ValueError, match=rf"settings\.env:2: invalid {key}: "):
        load_overlay(tmp_path)


def test_values_parse_with_bom_and_single_quote_pair(tmp_path):
    (tmp_path / "settings.env").write_text(
        "PROJECT_TITLE=\"'Quoted'\"\nWP1_NAME='unbalanced\"\nRP_LENGTH=4\nPROJECT_MONTHS=120\n",
        encoding="utf-8-sig",  # Windows editors prepend a BOM to the first key
    )
    assert load_overlay(tmp_path) == {
        "PROJECT_TITLE": "'Quoted'",
        "WP1_NAME": "'unbalanced\"",
        "RP_LENGTH": 4,
        "PROJECT_MONTHS": 120,
    }


_WINDOWS_DUMP = (
    "import json, budget_planner.config as c;"
    "rows = c.WORK_PACKAGES + c.REPORTING_PERIODS + c.DEMO_CAPACITY_ASSIGNMENTS + c.EXTERNAL_COMMITMENTS;"
    "print(json.dumps([c.PROJECT_MONTHS, c.DEMO_CURRENT_MONTH, [[r['start'], r['end']] for r in rows]]))"
)


@pytest.mark.parametrize("n", [1, 6, 14, 48, 60])
def test_fraction_windows_fit_any_overlaid_length(tmp_path, n):
    """Every seeded window satisfies 1 <= start <= end <= N whatever PROJECT_MONTHS the overlay sets."""
    (tmp_path / "settings.env").write_text(f"PROJECT_MONTHS={n}\n")
    env = {**os.environ, OVERLAY_DIR_ENV: str(tmp_path)}
    out = subprocess.run([sys.executable, "-c", _WINDOWS_DUMP], env=env, capture_output=True, text=True, check=True)
    months, demo_month, windows = json.loads(out.stdout)
    assert months == n  # vacuity guard: the overlay reached config
    assert 1 <= demo_month <= n
    bad = [(s, e) for s, e in windows if not 1 <= s <= e <= n]
    assert not bad, f"N={n}: windows outside 1..{n}: {bad}"
    assert max(e for _, e in windows) == n  # windows scale with N, not a fixed length


def test_rp_length_overlay_drives_reporting_periods(tmp_path):
    (tmp_path / "settings.env").write_text("PROJECT_MONTHS=14\nRP_LENGTH=4\n")
    env = {**os.environ, OVERLAY_DIR_ENV: str(tmp_path)}
    dump = (
        "import json, budget_planner.config as c; print(json.dumps([[r['start'], r['end']] for r in c.REPORTING_PERIODS]))"
    )
    out = subprocess.run([sys.executable, "-c", dump], env=env, capture_output=True, text=True, check=True)
    assert json.loads(out.stdout) == [[1, 4], [5, 8], [9, 12], [13, 14]]
