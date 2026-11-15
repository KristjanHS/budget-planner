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
