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


def test_generator_produces_workbook_matching_fingerprint(tmp_path: Path):
    """Round trip: fingerprint -> synthetic -> re-fingerprint = original fingerprint."""
    from budget_planner.workbook import init_workbook
    from scripts.build_from_shape import build_from_fingerprint

    src = tmp_path / "src.xlsx"
    init_workbook(str(src))
    fp = extract_fingerprint(str(src))

    synth = tmp_path / "synth.xlsx"
    build_from_fingerprint(fp, str(synth))

    fp2 = extract_fingerprint(str(synth))
    assert fp2["config_sections"] == fp["config_sections"], (
        "Synthetic must preserve row counts per Config section"
    )
    assert set(fp2["named_ranges"].keys()) == set(fp["named_ranges"].keys())
