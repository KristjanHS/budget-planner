"""Per-sheet workbook readers.

Public API mirrors the legacy ``budget_planner.reader`` module so callers can opt
into ``from budget_planner.readers import X`` directly. The legacy module
``budget_planner.reader`` remains as a thin re-export shim.
"""

from .actuals import read_actuals_override, read_actuals_values  # noqa: F401
from .actuals_hours import read_actuals_hours_sheet  # noqa: F401
from .capacity import read_capacity_assignments, read_capacity_sheet  # noqa: F401
from .config import (  # noqa: F401
    read_assignments_from_config,
    read_config_sheet,
    read_externals_from_config,
    read_grade_rates,
    read_plan_config,
    read_reporting_periods,
    read_team_from_config,
)
from .layout_discovery import (  # noqa: F401
    _data_row_after,
    _find_header_row,
    _iter_named_range_rows,
    _parse_config_float,
    discover_layout,
)
from .timelog import read_timelog_plan  # noqa: F401

__all__ = [
    "read_config_sheet",
    "read_plan_config",
    "read_grade_rates",
    "read_team_from_config",
    "read_reporting_periods",
    "read_assignments_from_config",
    "read_externals_from_config",
    "read_capacity_sheet",
    "read_capacity_assignments",
    "read_actuals_override",
    "read_actuals_values",
    "read_actuals_hours_sheet",
    "read_timelog_plan",
    "discover_layout",
]
