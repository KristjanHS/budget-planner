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
    assert str(live_team) in meta_text, (
        f"_META does not mention live team count ({live_team}). "
        "Was _META rewritten on last build?"
    )
