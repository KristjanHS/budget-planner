"""Privacy canary: confirm no real cell value ever leaves the real file.

Seeds a workbook with an unforgeable marker in every input cell, then runs the
reproducer pipeline end-to-end. The marker must not appear in:
  - the extracted fingerprint JSON
  - any cell value of the generated synthetic workbook

If either assertion fails, a cell value is leaking — investigate immediately.
"""
from __future__ import annotations

import json
from pathlib import Path

import openpyxl
from openpyxl.cell.cell import Cell

from budget_planner.workbook import init_workbook
from scripts.make_reproducer import make_reproducer
from scripts.shape_fingerprint import extract_fingerprint


MARKER = "CONFIDENTIAL-CANARY-9f2a1c8e"
MAGIC_NUMERIC = 987654321.0


def _seed_with_marker(path: Path) -> None:
    """Seed Config's *data* cells with marker-tagged values.

    The reader relies on structural scaffolding in column A (param labels
    like "Project Title", section headers like "TEAM ROSTER", column
    headers) to locate data regions. Destroying those breaks extraction
    before the privacy contract can even be evaluated.

    So: we seed every numeric cell (all user data is numeric or in col B+)
    and every string cell from column B onwards. Column A strings are left
    intact — they are structural labels the reader needs. The privacy
    claim being tested is: *user data values* (col B+ strings, all
    numerics) must never appear in the fingerprint or synthetic.
    """
    init_workbook(str(path))
    wb = openpyxl.load_workbook(path)
    ws = wb["Config"]
    for row in ws.iter_rows():
        for cell in row:
            # MergedCell.value is read-only; skip them.
            if not isinstance(cell, Cell):
                continue
            v = cell.value
            if isinstance(v, str) and not v.startswith("="):
                # Preserve col A (structural labels) so the reader still
                # finds its section anchors. Seed col B+ strings.
                if cell.column == 1:
                    continue
                cell.value = f"{MARKER}-{cell.coordinate}"
            elif isinstance(v, (int, float)) and not isinstance(v, bool):
                # Every numeric cell becomes the magic number — this is
                # where user FTE/rate/schedule data lives.
                cell.value = MAGIC_NUMERIC
    wb.save(path)


def test_marker_never_leaks_through_reproducer_pipeline(tmp_path: Path) -> None:
    real = tmp_path / "real_with_marker.xlsx"
    _seed_with_marker(real)

    # Fingerprint must be marker-free.
    fp = extract_fingerprint(str(real))
    fp_json = json.dumps(fp)
    assert MARKER not in fp_json, "Fingerprint leaked a cell string value"
    assert "987654321" not in fp_json, "Fingerprint leaked the magic numeric"

    # Synthetic must be marker-free.
    fp_path, synth_path = make_reproducer(str(real), out_dir=tmp_path / "out")

    # Double-check the on-disk fingerprint JSON too (not just the in-memory dict).
    fp_disk_text = fp_path.read_text(encoding="utf-8")
    assert MARKER not in fp_disk_text, "On-disk fingerprint JSON leaked a cell string value"
    assert "987654321" not in fp_disk_text, "On-disk fingerprint JSON leaked the magic numeric"

    wb = openpyxl.load_workbook(synth_path)
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                v = cell.value
                if isinstance(v, str):
                    assert MARKER not in v, (
                        f"Marker leaked into {ws.title}!{cell.coordinate}"
                    )
                elif isinstance(v, (int, float)) and not isinstance(v, bool):
                    assert v != MAGIC_NUMERIC, (
                        f"Magic number leaked into {ws.title}!{cell.coordinate}"
                    )
