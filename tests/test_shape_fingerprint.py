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
    assert fp2["config_sections"] == fp["config_sections"], "Synthetic must preserve row counts per Config section"
    assert set(fp2["named_ranges"].keys()) == set(fp["named_ranges"].keys())

    # Surviving scalar-cell named ranges (post-2026-04-18 through-month
    # cleanup) must actually be captured in the fingerprint — otherwise
    # widening SAFE_NAMED_RANGES silently adds ghost entries that neither
    # init nor synth emits, and the set-equality round-trip above passes
    # on a trivially empty intersection.
    surviving_names = {
        "hours_per_fte_month",
        "eur_actuals_through_month",
        "hours_actuals_through_month",
    }
    missing = surviving_names - fp["named_ranges"].keys()
    assert not missing, f"Fresh-init fingerprint missing surviving scalar named ranges: {sorted(missing)}"
    # Synthetic side must emit them too — otherwise the set-equality
    # round-trip above passes while ``build_from_fingerprint`` silently
    # drops them, leaving the forward guard on ``fp`` alone.
    missing_synth = surviving_names - fp2["named_ranges"].keys()
    assert not missing_synth, (
        f"Synthetic-rebuild fingerprint missing surviving scalar named ranges: {sorted(missing_synth)}"
    )
