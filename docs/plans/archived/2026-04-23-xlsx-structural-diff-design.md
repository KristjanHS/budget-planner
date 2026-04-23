# 2026-04-23 — xlsx structural diff (`scripts/xlsx_structural_diff.py`)

**Status:** DESIGN approved 2026-04-23 (brainstorm session). Pending implementation. Single-session scope (~300–400 LOC script + ~15 tests).

## Motivation

The user iterates on the generated workbook (`output/budget_plan.xlsx`) directly in Excel between regenerations — fixing formatting, tweaking formulas, adjusting layout. These manual edits are the implicit spec for the next round of generator changes. To feed them back into the AI agent that wrote v1 of `budget_planner/**`, we need a diff that:

1. Captures **structure and logic** (formulas, named ranges, merged regions, styles, data validation) so the agent can replicate the intent in code.
2. Emits **no confidential data** (rates, hours, team-member names, budget numbers) so the diff can be shared into an agent context without leaking the project's sensitive inputs.

The workbook itself conflates the two (a formula cell's displayed number is a data value; the formula driving it is logic). An AI-focused diff tool has to split them deliberately — existing xlsx diff tools (openpyxl-compare, sheetdiff) emit values by default and are not an option.

## Goal

Emit `diff.json` — agent-readable, deterministic key ordering, structured by sheet — capturing everything about the workbook's shape and logic while **never emitting literal cell values** except where explicitly whitelisted as structural labels (column A headers, sheet tab names, row-1 banners).

Consumer: the AI agent that wrote the v1 generator. Output must be self-describing enough that the agent can act on it without reading this design doc.

## Scope

### In scope (v1)
- Workbook-level: sheet add/remove, defined names add/remove/refers-to, workbook protection flag
- Per-sheet structure: dimensions, merged ranges, frozen panes, sheet protection, tab color, hidden state
- Per-cell classification: formula / literal / empty
- Formulas: full old/new diff of formula string
- Literals: presence + `data_type` only (never value), with safe-literal allowlist (D2)
- `number_format` always diffed
- Style field-level diff (bold, italic, font color, fill color, border, alignment, wrap)
- Data validation: type + formula + hashed list contents
- **Row-shift and column-shift hints** (D6) — heuristic detection of structural row/column insertion/deletion
- Stable key ordering + top-level `summary` counts block

### Out of scope for v1 (listed in Deferred section)
- Conditional formatting
- Tables / ListObjects
- Charts (beyond a stub note)
- Images / embedded objects
- Row heights / column widths (noise generators unless deliberately changed)
- Print area / print titles
- Sheet-rename auto-detection — emit added + removed, let the consumer infer
- Comments / threaded comments (PII risk)
- Author, lastModifiedBy, timestamps

## Decisions log

### D1 — Formula-literal asymmetry (confidentiality core)

- **Formulas** are emitted in full (old and new strings). A formula string is logic. The only confidential content a formula can carry is a literal numeric constant inside the expression — the consumer needs those to understand the logic anyway.
- **Literals** emit only `data_type` and `presence`, never the value. Exception: safe-literal allowlist (D2).
- **Why:** matches the project's confidentiality rule that budget numbers, rates, and named individuals never leave the machine unredacted. Formulas = shareable logic; values = not shareable data.

### D2 — Safe-literal allowlist for structural labels

Without this, a diff entry `E5: type_change n→s` is unreadable — the consumer can't tell `E5` is a header row vs a numeric cell that got converted to text.

- Emit literal value when **ALL** of:
  - String type (`data_type == "s"`)
  - Length ≤ 40 chars
  - Does **not** match the sensitive-value regex (see Implementation notes)
  - Cell is inside a configured "label zone": column A of every sheet, row 1 of every sheet, and (by default) the first 3 rows of every sheet
- Never emit literal value for numeric / date / bool types, regardless of location.
- **Why:** headers and row labels are load-bearing structure the consumer must see. Numbers and dates are data.

### D3 — Merged-range and shared-formula normalization

Before diffing, on each workbook side:
- For every merged range, map every coord inside the range to the top-left anchor. Diff anchor-to-anchor.
- For shared/array formulas, resolve to the anchor cell; ignore the phantom refs openpyxl exposes on non-anchor member cells.

**Why:** without this, an unrelated merge/unmerge or a shared-formula anchor shift creates dozens of phantom "cell disappeared" entries with no semantic meaning. This is specifically called out in the user's original design sketch.

### D4 — Style diff: field-level, not hashed

Per changed cell, emit only the style fields that changed:
```json
{"kind": "style", "bold": "false→true", "fill_color": "FFFFFF→E2EFDA"}
```
Skip unchanged fields entirely.

**Why:** the consumer needs to act on the change (replicate the user's formatting choice in the v2 generator). A style hash saying "something changed, old 0xabc new 0xdef" forces a round-trip read of both workbooks. Field-level diff costs maybe 2× the bytes on changed cells only, which is a small fraction of the whole diff.

### D5 — Named ranges promoted to top-level

Output schema places `defined_names` as a sibling of `summary`, immediately before `workbook` and per-sheet sections — not buried inside `workbook`.

**Why:** a `project_start` or `fte_rate` re-target is a structural earthquake that explains many downstream formula diffs. If the consumer reads per-sheet formula diffs first, it will infer "dozens of formulas rewritten" when the truth is "one named range repointed." Order of emit matters for consumer reasoning.

### D6 — Row-shift and column-shift hint heuristic

When a row or column is inserted or deleted in Excel, every formula on that sheet with references crossing the insertion point shifts its row/column refs in lockstep. Without a hint, the consumer sees 50+ formula diffs and infers a massive refactor.

**Detection algorithm (per sheet, run once per direction):**

1. Collect the set `CHANGED = {coord : (old_formula, new_formula)}` for cells where both sides have a formula and they differ.
2. For each `coord` in `CHANGED`, parse both formula strings with the A1-ref regex `(\$?)([A-Z]{1,3})(\$?)(\d+)`.
3. Pair refs by positional order (old formula's k-th ref ↔ new formula's k-th ref). Skip coords where old and new have a different ref count (not a pure shift — real edit).
4. For each paired ref, compute `(row_delta, col_delta) = (new_row - old_row, new_col - old_col)`.
5. Collect all `(row_delta, col_delta)` tuples across all eligible changed formulas on the sheet into a multiset.

**Row-shift hint is emitted when ALL:**
- At least 5 changed formulas on the sheet contributed to the multiset (noise floor — fewer means the pattern is coincidental).
- At least 80% of tuples in the multiset have `col_delta == 0` AND the same non-zero `row_delta == N`.
- The sheet's `max_row` changed in a direction consistent with `N` (`max_row` increased if `N > 0`, decreased if `N < 0`).

Hint payload:
```json
{
  "kind": "row_shift",
  "delta": 3,
  "confidence": "high",
  "inferred_start_row": 7,
  "sample_coords": ["C8", "C9", "D8"]
}
```
`inferred_start_row` is the smallest old-side row number referenced among the shifted formulas minus the delta; heuristic only.

**Column-shift hint is emitted symmetrically** with `row_delta == 0` and same non-zero `col_delta == N`, guarded by `max_col` direction match.

**Mixed shift (both row and column deltas dominate):** emit both hints with `"confidence": "medium"` and include a `"note": "mixed row+col shift — verify"` field. The 80% threshold applies per direction.

**False-positive guardrails (no hint emitted if):**
- Fewer than 5 changed formulas on the sheet.
- The dominant delta fails the 80% threshold (mixed deltas → let raw formula diffs speak).
- The dominant delta direction (e.g. `row_delta = +3`) doesn't match the `dimensions` change direction.

**Why 80%, not 100%:** a user who inserted a row and also manually edited 2–3 unrelated cells on the same sheet will have mostly-shifted formulas plus a few noise entries. 100% would miss the real case. 80% balances false-negative risk against false-positive risk.

**Why per-sheet, not workbook-wide:** Excel row/column insertions are per-sheet operations. Workbook-wide aggregation would blur distinct events on different sheets.

### D7 — Config sheet special-cased

The Config sheet has a known structure: column A is structural keys (`fte_rate`, `project_start`, WP dates); column B is confidential values (rates in EUR/hour, hours, dates).

- Column A: full safe-literal emit per D2 (keys are structural and need to be visible).
- Column B: **presence + type only**, overriding D2 — even when the value would pass the sensitive-value regex.
- **Why:** belt-and-suspenders for the project's most sensitive sheet. An auditor or accidental leak of the diff must not permit reconstruction of `Config!B`. The regex could miss an edge case (e.g. a round-number rate like `"80"` that doesn't trip any pattern); hard-coding the sheet-column suppression closes the gap.

### D8 — Stable ordering + summary header

- Sort all keys alphabetically in the emitted JSON: sheets, cell coords (natural A1 order: A1, B1, A2, B2), defined names.
- Emit a top-level `summary` block with counts **before** any detailed section:
```json
"summary": {
  "sheets_added": 1,
  "sheets_removed": 0,
  "defined_names_changed": 2,
  "formula_changes": 47,
  "cells_added": 3,
  "cells_removed": 1,
  "style_changes": 12,
  "hints_emitted": 2
}
```

**Why:** (a) diff-of-diff (consumer running this script twice across time) must be readable — unstable key order breaks textual diff. (b) The summary lets the consumer triage before reading a 200-line body. If `formula_changes == 0` and `hints_emitted == 0`, the consumer can skip per-sheet bodies entirely.

## Output schema

```json
{
  "meta": {
    "file_a": "path/to/v1.xlsx",
    "file_b": "path/to/current.xlsx",
    "generated_at": "2026-04-23T10:00:00Z",
    "schema_version": 1
  },
  "summary": { ... },
  "defined_names": {
    "added":   {"fte_rate_2027": "Config!$B$5"},
    "removed": {},
    "changed": {"project_start": {"old": "Config!$B$2", "new": "Config!$B$3"}}
  },
  "workbook": {
    "sheets_added":   ["Timelog Plan"],
    "sheets_removed": [],
    "protection_changed": false
  },
  "sheets": {
    "Plan": {
      "structure": {
        "dimensions":     {"old": {"max_row": 48, "max_col": 18}, "new": {"max_row": 49, "max_col": 18}},
        "merged_added":   ["A1:D1"],
        "merged_removed": [],
        "frozen_panes":   {"old": "A2", "new": "B3"},
        "tab_color":      {"old": null,  "new": "FFC0CB"},
        "hidden":         null
      },
      "cells": {
        "C5": {"kind": "formula", "old": "=SUM(A1:A4)", "new": "=SUM(A1:A5)"},
        "D5": {"kind": "formula", "added": "=C5*1.2"},
        "E5": {"kind": "literal", "type_change": {"old": "n", "new": "s"}, "value": "Subtotal"},
        "F5": {"kind": "style",   "bold": "false→true", "fill_color": "FFFFFF→E2EFDA"},
        "G5": {"kind": "format",  "number_format": {"old": "0.00", "new": "0.00%"}}
      },
      "validation_changes": [
        {
          "range": "B2:B10",
          "type_change":     {"old": "list",   "new": "whole"},
          "list_hash_old":   "sha256:abc123...",
          "list_hash_new":   null
        }
      ],
      "hints": [
        {
          "kind": "row_shift",
          "delta": 1,
          "confidence": "high",
          "inferred_start_row": 7,
          "sample_coords": ["C8", "C9", "D8", "E8", "F8"]
        }
      ]
    }
  }
}
```

## Implementation notes

### Dependencies
openpyxl only (already a project dep). No new packages.

### Core pass structure
1. `openpyxl.load_workbook(path, data_only=False)` for both files. Hold both `Workbook` objects for the full pass.
2. Normalize merged ranges and shared-formula anchors on each side (D3).
3. For each sheet present in either side, build `{coord: cell_signature}` dicts where `cell_signature` captures `kind`, `formula` (if any), `data_type` (if literal), `number_format`, style fields.
4. Symmetric-difference coord sets for added/removed; intersect and compare signatures for changed.
5. Build per-sheet `cells` dict from the diff.
6. Run row-shift heuristic, then column-shift heuristic, on the changed-formula subset (D6).
7. Diff defined names (workbook-level).
8. Assemble the final JSON with `sort_keys=True` and write to `--out` path (default `diff.json`).

### Safe-literal sensitive-value regex
```python
SENSITIVE_LITERAL_RE = re.compile(
    r"[\d][\d,.\s]*[€%]"             # currency/percent-ish numbers
    r"|@"                             # email markers
    r"|\b\d{4}-\d{2}-\d{2}\b"        # ISO dates
    r"|\b\d+[,.]\d+\b"               # decimals
    r"|\b\d{4,}\b"                    # 4+ digit numbers (IDs, years, round euros)
)
```
If a string literal matches, suppress the value even if other D2 conditions pass. Emit `{"kind": "literal", "type": "s", "value_suppressed": true}` so the consumer knows a value existed but was redacted.

### A1-ref parsing regex (for D6 row/col-shift detection)
```python
A1_REF_RE = re.compile(r"(\$?)([A-Z]{1,3})(\$?)(\d+)")
```
Column letters converted to 1-based index via standard `openpyxl.utils.column_index_from_string`.

### Style fingerprint — fields to extract
Extract explicitly; never serialize the openpyxl style object:
- `font.bold`, `font.italic`, `font.color.rgb` (guard against `color` being `None`)
- `fill.fgColor.rgb` only when `fill.patternType == "solid"`
- `border.{left, right, top, bottom}.style`
- `alignment.horizontal`, `alignment.vertical`, `alignment.wrap_text`

Missing/None values normalize to the string `"none"` for stable diffs.

### CLI
```
python scripts/xlsx_structural_diff.py FILE_A FILE_B [--out diff.json]
```
No other flags in v1. No `--verbose`; output is already filtered to structural signal.

### File placement
Single file: `scripts/xlsx_structural_diff.py`, ~300–400 LOC target. Lives alongside `scripts/recalc.py` and `scripts/restore.py`. **No new module under `budget_planner/`** — this is an analyst/dev utility, not a generator component. It must never be imported by `budget_planner/**`.

## Testing

One test file: `tests/test_xlsx_structural_diff.py`. Construct both input workbooks in-memory via `openpyxl.Workbook()`; no fixtures needed. Follows project `testing-project.md` rules (no hardcoded seeds; assert JSON shape, not hash signatures).

### Test cases (v1)
1. Formula change on same coord (old → new formula string emitted).
2. Formula added (cell empty in A, formula in B).
3. Formula removed.
4. Literal type flip (`n` → `s`) with value suppression.
5. Safe-literal allowlist: column-A string emitted in value field.
6. Safe-literal suppression: Config!B string never emitted even when short and non-sensitive-looking.
7. `number_format` change without formula change.
8. Style change without value or formula change.
9. Merged range added.
10. Defined name added / removed / refers-to changed.
11. Sheet added / sheet removed.
12. **Row-shift hint positive case:** insert a row at row 5, verify 5+ formulas on the sheet all gain `+1` on their row refs; assert hint with `delta=1, confidence="high"`.
13. **Row-shift hint dimension guard:** 5 changed formulas with same `+1` delta but `max_row` unchanged → no hint.
14. **Column-shift hint positive case:** symmetric to #12 in column direction.
15. **Mixed-shift case:** 50/50 row and column deltas → both hints with `"confidence": "medium"`.
16. **Sub-threshold case:** only 3 shifted formulas → no hint (noise floor).

### What the tests do **not** check
- Byte-exact JSON output (unstable across openpyxl versions in some edge fields). Check keys and field values instead.
- Style hash values (D4 — no hashing).

### Gate
Standard project gate: `pytest tests/ -q` passes before merge. Script itself doesn't run `recalc.py` — it reads two existing xlsx files as inputs and doesn't modify either.

## Acceptance criteria

- `python scripts/xlsx_structural_diff.py file_a.xlsx file_b.xlsx --out diff.json` produces valid, parseable JSON.
- `jq -r '.. | strings?' diff.json | grep -E '[0-9]+[,.][0-9]+'` returns no matches when both inputs are real workbooks (no decimal numbers leak).
- `jq -r '.. | strings?' diff.json | grep '@'` returns no matches (no emails leak).
- On two workbooks where only row 7 on sheet `Plan` was inserted (no other changes), `diff.json.sheets.Plan.hints` contains exactly one `row_shift` hint with `delta == 1` and `confidence == "high"`.
- On two identical workbooks, `diff.json.summary` reports all-zero counts and every `sheets.<name>.cells` is empty.
- All tests in `tests/test_xlsx_structural_diff.py` pass.

## Deferred to a future iteration (explicitly NOT v1)

- Conditional formatting rule diff
- Tables / ListObjects (name, range, header, totals formulas)
- Chart type + series ref diff (stub only — v1 ignores charts entirely)
- Per-sheet row-height / column-width diff
- Sheet-rename auto-detection (content-similarity heuristic)
- Configurable label-zone overrides per sheet (v1 uses hard-coded defaults; Config column A, every sheet row 1, every sheet first 3 rows)
- Human-readable text summary (in addition to JSON)
- Auto-redaction of suspicious safe-literal matches with a warning log

Add only when a concrete consumer need surfaces.
