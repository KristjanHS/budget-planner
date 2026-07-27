"""Direct unit tests for reader._parse_config_float.

Covers the locale-aware numeric-cell parser that underpins roster / assignments
/ externals readers. Integration tests in test_reader_capacity.py exercise the
"European 0,35 in a roster" path; these tests pin the edge-case contract so
future tweaks to separator logic, NBSP handling, or the None/0.0 sentinels
don't regress silently.
"""

from decimal import Decimal

import pytest

from budget_planner.reader import _parse_config_float


# ── Numeric inputs ──


@pytest.mark.parametrize(
    "value,expected",
    [
        (0, 0.0),
        (1, 1.0),
        (-3, -3.0),
        (0.5, 0.5),
        (1.25, 1.25),
        (Decimal("0.35"), 0.35),
        (Decimal("-2.5"), -2.5),
    ],
)
def test_numeric_inputs_are_coerced(value, expected):
    assert _parse_config_float(value) == pytest.approx(expected)


# ── Sentinel inputs (None vs "") ──


def test_none_returns_zero():
    """None (empty cell on a row the caller already validated) → 0.0."""
    assert _parse_config_float(None) == 0.0


@pytest.mark.parametrize("value", ["", "   ", "\u00a0", " \u00a0 "])
def test_empty_or_whitespace_returns_none(value):
    """Blank strings are a "non-numeric sentinel" — callers break on None."""
    assert _parse_config_float(value) is None


# ── Bool rejection ──


@pytest.mark.parametrize("value", [True, False])
def test_bool_is_rejected(value):
    """bool is an int subclass; reject explicitly so True/False don't coerce to 1.0/0.0."""
    assert _parse_config_float(value) is None


# ── String locale variants ──


@pytest.mark.parametrize(
    "value,expected",
    [
        ("0,35", 0.35),  # EU comma decimal
        ("1,5", 1.5),
        ("1.5", 1.5),  # US dot decimal
        ("1\u00a0234,56", 1234.56),  # EU with NBSP thousands + comma decimal
        ("1.234,56", 1234.56),  # EU dot-thousands + comma decimal
        ("1,234.56", 1234.56),  # US comma-thousands + dot decimal
        ("1 234,56", 1234.56),  # regular-space thousands
        ("-0,5", -0.5),
        ("1e3", 1000.0),  # scientific notation
    ],
)
def test_string_locale_variants(value, expected):
    assert _parse_config_float(value) == pytest.approx(expected)


# ── Non-numeric strings (end-of-section sentinels) ──


@pytest.mark.parametrize(
    "value",
    [
        "=100",  # formula string (data_only=False reads)
        "=SUM(A1:A5)",
        "WP BUDGET",  # next-section header
        "Total FTE",
        "abc",
        "1.2.3",
    ],
)
def test_non_numeric_strings_return_none(value):
    assert _parse_config_float(value) is None
