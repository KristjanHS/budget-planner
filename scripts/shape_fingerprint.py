"""Shape fingerprint — value-free structural snapshot of a workbook.

Privacy contract: the returned dict contains ONLY counts, row/column numbers,
sheet names, and named-range names. No cell values ever. If you extend this
module, do not add fields that store string or numeric data read from cells.
A canary test guards this contract — see tests/test_shape_fingerprint.py.
"""
from __future__ import annotations

import json
import re
import sys
from pathlib import Path

import openpyxl

from budget_planner.reader import (
    read_team_from_config,
    read_assignments_from_config,
    read_externals_from_config,
    read_reporting_periods,
    read_grade_rates,
    read_config_sheet,
)

SAFE_NAMED_RANGES = [
    "roster_ftes", "roster_grades",
    "assign_persons", "assign_wps", "assign_ftes",
    "ext_persons", "ext_ftes",
    "sched_wp_ids", "sched_starts", "sched_ends",
    "tbl_GradeRates", "project_start", "fte_rate",
    "eu_annual_ceiling_pm", "partner_ceiling", "current_month",
    # v2 hours-timelog named ranges: the through-month trust boundary.
    # eur_/hours_actuals_through_month point to per-sheet D1 effective cells
    # on Actuals / Actuals-Hours after the 2026-04-18 Config-through-month
    # cleanup; derived helpers (months_elapsed_*, reconciled_through_month)
    # are inlined at callsites and no longer exist as named ranges.
    "hours_per_fte_month",
    "eur_actuals_through_month", "hours_actuals_through_month",
    # Post-v2 version-stamp audit cell on Config. Single-cell named range;
    # value written fresh by init/rebuild so build_from_fingerprint emits
    # "unknown" (the default when no git context is passed).
    "version_stamp",
]

_REF_RE = re.compile(r"\$?([A-Z]+)\$?(\d+)(?::\$?([A-Z]+)\$?(\d+))?")


def _named_range_entry(wb, name):
    try:
        dn = wb.defined_names[name]
    except (KeyError, TypeError):
        return None
    try:
        dests = list(dn.destinations)
    except Exception:
        return None
    if not dests:
        return None
    sheet, ref = dests[0]
    m = _REF_RE.search(ref)
    if not m:
        return None
    col1, row1, col2, row2 = m.group(1), int(m.group(2)), m.group(3), m.group(4)
    first_row = row1
    last_row = int(row2) if row2 else row1
    cols = sorted({col1, col2}) if col2 else [col1]
    return {"sheet": sheet, "first_row": first_row, "last_row": last_row, "cols": cols}


def extract_fingerprint(path: str) -> dict:
    wb = openpyxl.load_workbook(path, data_only=False)
    config_ws = wb["Config"]

    named_ranges = {}
    for nm in SAFE_NAMED_RANGES:
        entry = _named_range_entry(wb, nm)
        if entry is not None:
            named_ranges[nm] = entry

    _, wps, sched = read_config_sheet(config_ws)
    sections = {
        "work_packages":        {"row_count": len(wps)},
        "wp_schedule":          {"row_count": len(sched)},
        "team_roster":          {"row_count": len(read_team_from_config(config_ws))},
        "assignments":          {"row_count": len(read_assignments_from_config(config_ws))},
        "external_commitments": {"row_count": len(read_externals_from_config(config_ws))},
        "reporting_periods":    {"row_count": len(read_reporting_periods(config_ws))},
        "grade_rates":          {"row_count": len(read_grade_rates(config_ws))},
    }

    cross = {}
    for sheet_name in ["Capacity", "Plan", "Actuals", "Summary", "Trends"]:
        if sheet_name not in wb.sheetnames:
            continue
        ws = wb[sheet_name]
        cells: set[tuple[str, int]] = set()
        total = 0
        for row in ws.iter_rows():
            for cell in row:
                v = cell.value
                if isinstance(v, str) and v.startswith("=") and "Config!" in v:
                    for m in _REF_RE.finditer(v):
                        total += 1
                        cells.add((m.group(1), int(m.group(2))))
        cross[sheet_name] = {"config_refs": total, "distinct_config_cells": len(cells)}

    return {
        "schema_version": 1,
        "sheet_names": list(wb.sheetnames),
        "named_ranges": named_ranges,
        "config_sections": sections,
        "cross_sheet_formula_targets": cross,
    }


DEFAULT_OUT_PATH = Path("tmp/fp.json")


def main(argv: list[str]) -> int:
    if len(argv) < 2 or argv[1] in {"-h", "--help"}:
        print(
            "usage: python -m scripts.shape_fingerprint <workbook.xlsx> [-o <path>] [--stdout]\n"
            f"       default output: {DEFAULT_OUT_PATH} (overwritten each run)",
            file=sys.stderr,
        )
        return 2
    in_path = argv[1]
    out_path: Path | None = DEFAULT_OUT_PATH
    if "-o" in argv:
        out_path = Path(argv[argv.index("-o") + 1])
    if "--stdout" in argv:
        out_path = None
    fp = extract_fingerprint(in_path)
    js = json.dumps(fp, indent=2, sort_keys=True)
    if out_path is None:
        print(js)
    else:
        out_path.parent.mkdir(parents=True, exist_ok=True)
        out_path.write_text(js, encoding="utf-8")
        print(f"Fingerprint written to {out_path}. Safe to share — value-free by construction.")
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv))
