"""Fingerprint must be value-free by construction."""
from pathlib import Path

from scripts.shape_fingerprint import extract_fingerprint


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

    # v2 hours-timelog named ranges must actually be captured in the
    # fingerprint — otherwise widening SAFE_NAMED_RANGES silently adds
    # ghost entries that neither init nor synth emits, and the set-equality
    # round-trip above passes on a trivially empty intersection.
    v2_names = {
        "hours_per_fte_month",
        "eur_actuals_through_month", "hours_actuals_through_month",
        "months_elapsed_hours", "months_elapsed_eur",
        "reconciled_through_month",
    }
    missing = v2_names - fp["named_ranges"].keys()
    assert not missing, (
        f"Fresh-init fingerprint missing v2 Config named ranges: "
        f"{sorted(missing)}"
    )
