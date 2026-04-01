"""Tests for TEAM and EXTERNAL_COMMITMENTS seed data in config."""

from budget_planner.config import TEAM, EXTERNAL_COMMITMENTS, WORK_PACKAGES


def _wp_ids() -> set[str]:
    return {wp["id"] for wp in WORK_PACKAGES}


def test_team_data_has_required_fields():
    """Each entry in TEAM has: name (str), max_fte (float > 0), wps (dict str->float)."""
    for member in TEAM:
        assert isinstance(member["name"], str)
        assert isinstance(member["max_fte"], (int, float))
        assert member["max_fte"] > 0
        assert isinstance(member["wps"], dict)
        for wp_id, share in member["wps"].items():
            assert isinstance(wp_id, str)
            assert isinstance(share, (int, float))


def test_team_wp_shares_reference_valid_wps():
    """Every WP ID in TEAM[*].wps exists in WORK_PACKAGES[*].id."""
    valid_ids = _wp_ids()
    for member in TEAM:
        for wp_id in member["wps"]:
            assert wp_id in valid_ids, (
                f"{member['name']} references unknown WP '{wp_id}'"
            )


def test_external_commitments_has_required_fields():
    """Each entry has: person (str), project (str), fte (float>0), start/end (int, start<=end)."""
    for entry in EXTERNAL_COMMITMENTS:
        assert isinstance(entry["person"], str)
        assert isinstance(entry["project"], str)
        assert isinstance(entry["fte"], (int, float))
        assert entry["fte"] > 0
        assert isinstance(entry["start"], int)
        assert isinstance(entry["end"], int)
        assert entry["start"] <= entry["end"], (
            f"{entry['person']}'s commitment has start ({entry['start']}) > end ({entry['end']})"
        )


def test_external_commitments_persons_in_team():
    """Every person in EXTERNAL_COMMITMENTS exists in TEAM[*].name."""
    team_names = {m["name"] for m in TEAM}
    for entry in EXTERNAL_COMMITMENTS:
        assert entry["person"] in team_names, (
            f"External commitment references unknown person '{entry['person']}'"
        )
