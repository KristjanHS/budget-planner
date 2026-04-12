"""Shared test invariants for rebuilt workbooks.

Leading underscore keeps pytest from collecting this module as a test file.
"""
from budget_planner.reader import read_team_from_config


def assert_meta_matches_live(wb) -> None:
    """_META text must reflect the live team count after the last build.

    _META is a documentation blob regenerated on every build. This is a coarse
    sanity check that it was actually rewritten and not carried over stale
    (which would mask shape drift). Counts are printed as bare integers, so we
    look for the live team count somewhere in the blob.
    """
    meta_ws = wb["_META"]
    meta_text = meta_ws["A1"].value or ""
    live_team = len(read_team_from_config(wb["Config"]))
    # meta.py emits the team count as "{num_team} people" (lines 185, 219).
    # Matching the phrase — not bare digits — makes the invariant specific
    # enough to catch stale _META blobs where the count didn't update.
    expected = f"{live_team} people"
    assert expected in meta_text, (
        f"_META does not mention {expected!r}. "
        "Was _META rewritten on last build, or has meta.py's phrasing changed?"
    )
