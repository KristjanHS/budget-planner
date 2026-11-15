"""Optional seed overrides: a `KEY=VALUE` overlay file in the dir named by BUDGET_PLANNER_OVERLAY_DIR.

Unset → `../budget-planner-overlay` beside the repo; empty string → disabled (tests pin this).
"""

import os
from datetime import datetime
from pathlib import Path

OVERLAY_DIR_ENV = "BUDGET_PLANNER_OVERLAY_DIR"
_DEFAULT_OVERLAY_DIR = Path(__file__).resolve().parents[2] / "budget-planner-overlay"


def _month_count(value):
    n = int(value)
    if not 1 <= n <= 120:
        raise ValueError(f"must be 1-120 months, got {n}")
    return n


def _rate(value):
    n = int(value)
    if n <= 0:
        raise ValueError(f"must be > 0, got {n}")
    return n


def _first_of_month(value):
    d = datetime.strptime(value, "%Y-%m-%d")
    if d.day != 1:
        raise ValueError(f"must fall on day 1 of a month, got {value}")
    return d


_PARSERS = {
    "PROJECT_TITLE": str,
    "PROJECT_START": _first_of_month,
    "PROJECT_MONTHS": _month_count,
    "RP_LENGTH": _month_count,
    "DEFAULT_RATE": _rate,
    "GRADE_RATE_SENIOR": _rate,
    "GRADE_RATE_MID": _rate,
    "GRADE_RATE_JUNIOR": _rate,
    **{f"WP{n}_NAME": str for n in range(1, 10)},
}


def _unquote(value):
    """Strip one matching pair of surrounding quotes, nothing more."""
    if len(value) >= 2 and value[0] == value[-1] and value[0] in "\"'":
        return value[1:-1]
    return value


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
    for lineno, raw in enumerate(path.read_text(encoding="utf-8-sig").splitlines(), 1):
        line = raw.strip()
        if not line or line.startswith("#"):
            continue
        key, sep, value = line.partition("=")
        key, value = key.strip(), _unquote(value.strip())
        if not sep or key not in _PARSERS:
            raise ValueError(f"{path}:{lineno}: unknown or malformed setting {key!r}")
        try:
            settings[key] = _PARSERS[key](value)
        except ValueError as exc:
            raise ValueError(f"{path}:{lineno}: invalid {key}: {exc}") from exc
    return settings
