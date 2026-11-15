"""Optional seed overrides: a `KEY=VALUE` overlay file in the dir named by BUDGET_PLANNER_OVERLAY_DIR.

Unset → `../budget-planner-overlay` beside the repo; empty string → disabled (tests pin this).
"""

import os
from datetime import datetime
from pathlib import Path

OVERLAY_DIR_ENV = "BUDGET_PLANNER_OVERLAY_DIR"
_DEFAULT_OVERLAY_DIR = Path(__file__).resolve().parents[2] / "budget-planner-overlay"

_PARSERS = {
    "PROJECT_TITLE": str,
    "PROJECT_START": lambda v: datetime.strptime(v, "%Y-%m-%d"),
    "PROJECT_MONTHS": int,
    "DEFAULT_RATE": int,
    "GRADE_RATE_SENIOR": int,
    "GRADE_RATE_MID": int,
    "GRADE_RATE_JUNIOR": int,
    **{f"WP{n}_NAME": str for n in range(1, 10)},
}


def overlay_dir_from_env(environ=os.environ):
    """Resolve the companion dir; None when the overlay is disabled."""
    value = environ.get(OVERLAY_DIR_ENV)
    if value is None:
        return _DEFAULT_OVERLAY_DIR
    return Path(value) if value else None


def load_overlay(overlay_dir):
    """Typed overrides from the overlay file in `overlay_dir`; {} when the dir or file is absent."""
    if overlay_dir is None:
        return {}
    path = Path(overlay_dir) / "settings.env"
    if not path.is_file():
        return {}
    settings = {}
    for lineno, raw in enumerate(path.read_text(encoding="utf-8").splitlines(), 1):
        line = raw.strip()
        if not line or line.startswith("#"):
            continue
        key, sep, value = line.partition("=")
        key, value = key.strip(), value.strip().strip("\"'")
        if not sep or key not in _PARSERS:
            raise ValueError(f"{path}:{lineno}: unknown or malformed setting {key!r}")
        settings[key] = _PARSERS[key](value)
    return settings
