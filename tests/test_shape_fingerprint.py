"""Fingerprint must be value-free by construction."""
from pathlib import Path

from scripts.shape_fingerprint import extract_fingerprint


def test_fingerprint_contains_only_structural_fields(tmp_path: Path):
    from budget_planner.workbook import init_workbook
    path = tmp_path / "wb.xlsx"
    init_workbook(str(path))

    fp = extract_fingerprint(str(path))

    # Top-level keys are the whitelist — anything else fails.
    assert set(fp.keys()) == {
        "schema_version", "sheet_names", "named_ranges",
        "config_sections", "cross_sheet_formula_targets",
    }
    # Named ranges: only structural fields.
    for nr in fp["named_ranges"].values():
        assert set(nr.keys()) <= {"sheet", "first_row", "last_row", "cols"}
    # Config sections: only counts.
    for sec in fp["config_sections"].values():
        assert set(sec.keys()) == {"row_count"}
