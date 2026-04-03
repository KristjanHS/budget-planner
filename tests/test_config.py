"""Tests for TEAM and EXTERNAL_COMMITMENTS seed data in config."""

from budget_planner.config import TEAM, EXTERNAL_COMMITMENTS


def test_external_commitments_persons_in_team():
    """Every person in EXTERNAL_COMMITMENTS exists in TEAM[*].name."""
    team_names = {m["name"] for m in TEAM}
    for entry in EXTERNAL_COMMITMENTS:
        assert entry["person"] in team_names, (
            f"External commitment references unknown person '{entry['person']}'"
        )
