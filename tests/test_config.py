"""Tests for TEAM and EXTERNAL_COMMITMENTS seed data in config."""

from budget_planner.config import TEAM, EXTERNAL_COMMITMENTS, WORK_PACKAGES


def _wp_ids() -> set[str]:
    return {wp["id"] for wp in WORK_PACKAGES}


def test_team_wp_shares_reference_valid_wps():
    """Every WP ID in TEAM[*].wps exists in WORK_PACKAGES[*].id."""
    valid_ids = _wp_ids()
    for member in TEAM:
        for wp_id in member["wps"]:
            assert wp_id in valid_ids, (
                f"{member['name']} references unknown WP '{wp_id}'"
            )


def test_external_commitments_persons_in_team():
    """Every person in EXTERNAL_COMMITMENTS exists in TEAM[*].name."""
    team_names = {m["name"] for m in TEAM}
    for entry in EXTERNAL_COMMITMENTS:
        assert entry["person"] in team_names, (
            f"External commitment references unknown person '{entry['person']}'"
        )
