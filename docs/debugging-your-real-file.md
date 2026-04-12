# Debugging shape bugs in your confidential workbook

**When to use this:** You hit a bug in `rebuild` (or any other code path) on
your real `output/<your_confidential>.xlsx`. You want Claude's help reproducing
and fixing it, without sharing the real file.

## One-shot workflow

```bash
python scripts/make_reproducer.py \
    output/<your_confidential>.xlsx \
    tmp/reproducer.xlsx
```

From Windows:

```bat
scripts\make_reproducer.bat output\<your_confidential>.xlsx tmp\reproducer.xlsx
```

This produces `tmp/reproducer.xlsx` — a synthetic workbook with the same table
shapes (row counts per Config section, named-range spans, cross-sheet formula
targets) as your real file, populated with deterministic fake data
(`Person_01`, `WP_A`, fixed dates and rates). Safe to share.

## What to send Claude

1. A description of the bug ("After rebuild, Alice is missing from Config").
2. The path to the reproducer (`tmp/reproducer.xlsx`).
3. Optionally, the fingerprint JSON for extra context:

```bash
python scripts/make_reproducer.py \
    output/<your_confidential>.xlsx \
    tmp/reproducer.xlsx \
    --keep-fingerprint tmp/fp.json
```

Do NOT send the real workbook. Claude will not open it.

## Privacy guarantees

- `scripts/shape_fingerprint.py` reads only row counts, named-range spans, and
  formula coordinates. No cell values ever leave the real file.
- `scripts/build_from_shape.py` synthesises all names and numbers
  deterministically from the fingerprint. It never receives cell values from
  your real file.
- The pipeline is guarded by `tests/test_privacy_canary.py` — seeds every
  Config cell with a marker string, runs the full pipeline, and asserts the
  marker never appears in the fingerprint JSON or the synthetic workbook.
  If any cell value ever leaks, this test fails loudly.

## Re-running after a fix

Rerun `make_reproducer.py` whenever you edit your real file — the reproducer
is cheap to regenerate and always reflects the current shape.
