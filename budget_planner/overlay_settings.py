"""Optional seed overrides: a `KEY=VALUE` overlay file in the dir named by BUDGET_PLANNER_OVERLAY_DIR.

Env set → that dir; set to empty → disabled (tests pin this); unset → the same key read from the
dotfiles layer `${DOTFILES_DIR:-<main checkout>/../dotfiles}/shared-settings/` file (parsed, never
sourced); no key or no file → no overlay, the public defaults in config.py stand.
"""

import os
from datetime import datetime
from pathlib import Path

OVERLAY_DIR_ENV = "BUDGET_PLANNER_OVERLAY_DIR"
DOTFILES_DIR_ENV = "DOTFILES_DIR"
SETTINGS_FILE = "settings.env"  # overlay file name, and the dotfiles layer's file name
DOTFILES_LAYER = Path("shared-settings") / SETTINGS_FILE
_REPO_ROOT = Path(__file__).resolve().parents[1]


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


def _key_values(path):
    """Yield (lineno, key, has_equals, value) per non-blank, non-comment `KEY=value` line."""
    for lineno, raw in enumerate(path.read_text(encoding="utf-8-sig").splitlines(), 1):
        line = raw.strip()
        if not line or line.startswith("#"):
            continue
        key, sep, value = line.partition("=")
        yield lineno, key.strip(), bool(sep), _unquote(value.strip())


def _main_checkout(repo):
    """The primary checkout: a linked worktree's `.git` file points at `<main>/.git/worktrees/<name>`."""
    dotgit = repo / ".git"
    if dotgit.is_file():
        gitdir = dotgit.read_text(encoding="utf-8").partition("gitdir:")[2].strip()
        target = repo / gitdir  # an absolute gitdir replaces repo
        if gitdir and target.parent.name == "worktrees":
            return target.parents[2].resolve()
    return repo


def _dotfiles_pointer(environ, repo):
    """The overlay dir named in the dotfiles layer; None when the file or the key is absent."""
    dotfiles = environ.get(DOTFILES_DIR_ENV) or _main_checkout(repo).parent / "dotfiles"
    layer = Path(dotfiles) / DOTFILES_LAYER
    if not layer.is_file():
        return None
    found = None
    for _, key, has_equals, value in _key_values(layer):
        if has_equals and key == OVERLAY_DIR_ENV:
            found = value  # last assignment wins, as in a shell
    return Path(found).expanduser() if found else None


def overlay_dir_from_env(environ=os.environ, repo=_REPO_ROOT):
    """Resolve the overlay dir; None when the overlay is disabled or not configured."""
    value = environ.get(OVERLAY_DIR_ENV)
    if value is None:
        return _dotfiles_pointer(environ, repo)
    return Path(value).expanduser() if value else None


def load_overlay(overlay_dir):
    """Typed overrides from the overlay file in `overlay_dir`; {} when the dir or file is absent."""
    if overlay_dir is None:
        return {}
    path = Path(overlay_dir) / SETTINGS_FILE
    if not path.is_file():
        return {}
    settings = {}
    for lineno, key, has_equals, value in _key_values(path):
        if not has_equals or key not in _PARSERS:
            raise ValueError(f"{path}:{lineno}: unknown or malformed setting {key!r}")
        try:
            settings[key] = _PARSERS[key](value)
        except ValueError as exc:
            raise ValueError(f"{path}:{lineno}: invalid {key}: {exc}") from exc
    return settings
