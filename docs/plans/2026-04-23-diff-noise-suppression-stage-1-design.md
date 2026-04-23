# Diff Noise-Suppression — Stage 1 Design (D + A + C)

**Date**: 2026-04-23
**Status**: Ready for `/impag`
**Seed**: `docs/plans/2026-04-23-diff-noise-suppression-ideas.md`
**Script**: `scripts/xlsx_structural_diff.py` (809 lines)
**Tests**: `tests/test_xlsx_structural_diff.py` (585 lines, 25 tests)
**Callers**: `diff.bat` + test file only. No downstream Python importers.
**Stage 2 sibling**: `docs/plans/2026-04-23-diff-noise-suppression-stage-2-design.md` (B1/B2 structural-shift detection — skeleton only, full design deferred)

---

## 0 — Resolved brainstorm decisions

| # | Decision | Source |
|---|----------|--------|
| 1 | Scope = full overhaul (A + B + C + D) split across two stages | User, Q1 |
| 2 | Suppression model = drop-by-default + top-level `noise_filtered_summary`; `--keep-noise` retains records and tags them with `noise_class` | User, Q2 |
| 3 | Normalization lives in comparator (custom equality predicates), not reader. Raw values preserved end-to-end; `--strict-*` flags become one-line flag flips if ever needed | User, Q3 |
| 4 | Structural-shift detection: B1 (row) + B2 (col) in Stage 2; B3 (table rename) deferred — zero observed instances so far | User, Q4 |
| 5 | Staging = 2 files. Stage 1 = D + A + C (this doc). Stage 2 = B1/B2 (sibling doc) | User, Q5 |
| 6 | `schema_version` bumps to `2` in Stage 1. Stage 2 adds `structural_shifts` key additively — no further bump | Q5 rationale |

---

## 1 — Goals

1. **N4 bug fix.** `added`/`removed` branches preserve `value` for string-literal cells the same way `kind_change` does.
2. **N1 noise at source.** Number-format comparison canonicalizes case + backslash escapes before `==`.
3. **N2 noise at source.** ARGB color comparison normalizes the alpha byte (`FF` and `00` both mean "fully opaque"; Excel ignores alpha for tab colors).
4. **Transparent suppression.** Always emit top-level `noise_filtered_summary` counting silenced deltas per category.
5. **Reversible suppression.** `--keep-noise` retains records and stamps each with `noise_class` (`"format"` or `"color"` in Stage 1). Comparator still uses canonical predicate, so same cells flag as equal — the tag exposes *why* a record was silenced.
6. **Schema v1 → v2.** Additive only; document the delta in the module docstring.

## 2 — Non-goals (Stage 1)

- No structural-shift detection (Stage 2).
- No change to `--mode=logic` filter rules except N4.
- No `"structural"` value in `noise_class` yet — Stage 2 adds it.
- No B3 (table rename).
- No `--strict-format` / `--strict-color` flags. `--keep-noise` already exposes what was silenced; strict-mode ships only if asked.

---

## 3 — CLI surface

```
--keep-noise    Retain records that would otherwise be dropped by noise
                canonicalization; tag each with noise_class.
```

No other CLI changes. `diff.bat` invocation unchanged.

---

## 4 — Schema v2 shape

Additions only. No key renames, no key removals.

```json
{
  "schema_version": 2,
  "summary": {
    "formula_changes": 190,
    "style_changes": 0,
    "cells_added": 10,
    "cells_removed": 75
  },
  "noise_filtered_summary": {
    "number_format_normalized": 2251,
    "tab_color_alpha_normalized": 8
  },
  "sheets": {
    "Config": {
      "cells": {
        "A1": {
          "data_type": "s",
          "value": "Foo",
          "noise_class": "format"
        }
      }
    }
  }
}
```

- `noise_filtered_summary` — always emitted; zero counts stay as `0`. Makes "nothing was silenced" explicit.
- `noise_class` — only present when `--keep-noise` and the record was tagged. Values in Stage 1: `"format"`, `"color"`.
- `summary.formula_changes` / `summary.style_changes` — already post-normalization counts. A consumer comparing v1 and v2 output over the same xlsx pair should see these numbers drop by exactly `noise_filtered_summary` totals.

---

## 5 — Canonicalization rules

Both live as module-level helpers in `scripts/xlsx_structural_diff.py`. Unit-tested directly.

### 5.1 — `_canon_numfmt(s: str) -> str`
```
1. s = s.upper()
2. s = s.replace("\\€", "€").replace("\\$", "$").replace("\\%", "%")
3. while s and s[0] == "\\": s = s[1:]
4. return s
```

Equivalence examples:
- `mmm\ yyyy` ≡ `MMM YYYY`
- `\€#,##0` ≡ `€#,##0`
- `mmm\ yy` ≡ `MMM YY`

Non-equivalence preserved: `"0.00"` ≠ `"0.0000"`; `"mmm yyyy"` ≠ `"mm yyyy"`.

### 5.2 — `_canon_argb(s: str) -> str`
```
if isinstance(s, str) and len(s) == 8 and s[:2].upper() in ("FF", "00"):
    return "__" + s[2:].upper()
return s if not isinstance(s, str) else s.upper()
```

Equivalence examples:
- `"FF0000FF"` ≡ `"000000FF"` → both canonicalize to `"__0000FF"`
- `"FFC00000"` ≡ `"00C00000"` → both → `"__C00000"`

Non-equivalence preserved: `"FF0000FF"` ≠ `"FF000080"` (different non-alpha bytes); non-8-char values pass through unchanged.

### 5.3 — Where predicates are applied

The comparator calls these inside the equality check for the two affected fields. Pseudocode:
```python
def _numfmt_equal(a, b, counter):
    if a == b: return True
    if _canon_numfmt(a) == _canon_numfmt(b):
        counter["number_format_normalized"] += 1
        return True
    return False
```

`counter` is the `noise_filtered_summary` accumulator, threaded through the existing comparator call-tree. Same pattern for `_argb_equal`.

When `--keep-noise` is set, the predicate still returns `True` (cells still compare equal) but the caller emits the record with `noise_class` stamped. Implementation: predicate returns a `(equal: bool, noise_class: str | None)` tuple; comparator decides whether to drop or tag based on the flag.

---

## 6 — N4 bug fix (D1 + D2)

### 6.1 — Audit (D2)

Locate every `added`/`removed` emission path for string-literal cells. Confirm each either:
- (a) already serializes `value` — leave alone, or
- (b) drops `value` — fix to serialize when `data_type == "s"`.

The seed identifies `scripts/xlsx_structural_diff.py` "branch that emits `kind: 'literal'` without a `value` field" as the target. Audit confirms by grep for `kind.*literal` and `"added": True` / `"removed": True` emission sites.

### 6.2 — Fix

On `added` and `removed` branches: when `data_type == "s"` and the cell has a string value, include `"value": <str>` in the emitted record. Numeric cells (`data_type == "n"`) leave `value` off — numbers are formula results, not semantic intent.

### 6.3 — Regression fixture

Two in-memory workbooks built via `openpyxl.Workbook()`:
- `a`: `ws["A1"] = "Foo"`, save to tmp path
- `b`: (empty A1), save to tmp path

Run logic-mode diff. Assert:
- Record for `A1` exists under the `removed` branch (comparing a→b direction) or `added` branch (b→a direction).
- Record contains `{"data_type": "s", "value": "Foo"}` plus the `removed: true` / `added: true` flag.

Two tests — one for each direction. Use `pytest tmp_path` fixture.

---

## 7 — Task breakdown for `/impag`

| # | Task | Files | Tests |
|---|------|-------|-------|
| 1 | Add `_canon_numfmt` + `_canon_argb` helpers | `scripts/xlsx_structural_diff.py` | 2 unit tests (canon round-trips, non-equivalence preserved) |
| 2 | Wire canonicalizing predicates into number-format + tab-color comparators; accumulate counts into `noise_filtered_summary`; emit at top level | `scripts/xlsx_structural_diff.py` | 2 tests: format-noise suppressed by default, color-noise suppressed by default; `noise_filtered_summary` always present |
| 3 | Add `--keep-noise` CLI flag; when set, retain records and stamp `noise_class` | `scripts/xlsx_structural_diff.py` | 2 tests: `--keep-noise` retains format record with `noise_class: "format"`, same for color |
| 4 | Fix N4: serialize `value` on `added`/`removed` string-literal branches | `scripts/xlsx_structural_diff.py` | 2 tests: added string preserved, removed string preserved (per §6.3 fixture) |
| 5 | Bump `schema_version` to `2`; update module docstring with v1→v2 delta | `scripts/xlsx_structural_diff.py` | 1 test: top-level record contains `"schema_version": 2` |

Total: +~60 LOC in script, +~9 tests in test file (25 → 34).

All tasks touch `scripts/xlsx_structural_diff.py`. Per `/impag` same-file-serialize rule: run sequentially 1→2→3→4→5. Task 1 (helpers) must land before task 2 (wires them into comparator); tasks 3, 4, 5 are logically independent of 2 but serialized on file to avoid subagent merge conflicts.

---

## 8 — Risks & mitigations

| Risk | Likelihood | Impact | Mitigation |
|------|:---:|:---:|------------|
| Canonicalization over-matches (two genuinely different formats collapse) | Low | Med | Unit tests assert non-equivalence cases (`"0.00"` ≠ `"0.0000"`, `"mmm yyyy"` ≠ `"mm yyyy"`). Add counterexamples to test suite before wiring predicate. |
| `\` stripping misses an escape sequence we haven't seen (e.g. `\₽`) | Med | Low | Keep rule #2 explicit for known currency symbols only (€/$/%); rule #3 (leading-backslash strip) is general. Unknown-char cases fall through as-is. If tests expose a real case, extend rule #2. |
| ARGB alpha rule too aggressive — some ARGB field actually uses alpha | Low | Med | Scope is tab colors + (if reused) fill/font colors. Excel ignores alpha everywhere we currently compare. Document the assumption in the helper docstring; if a future CF uses alpha=0x80 semantically, exempt that comparator. |
| `--keep-noise` implementation path differs from default path (divergence) | Med | Med | One predicate returns `(equal, noise_class)`. Single call site in comparator decides drop-vs-tag via the flag. No duplicated comparison logic. |
| Schema v2 breaks a consumer we don't know about | Low | Low | Only caller is `diff.bat`; grep confirmed. Additive-only shape — old v1 consumers that read known keys continue to work. |
| N4 fix changes test fixtures for pre-existing tests | Low | Low | Existing tests assert specific record shapes. Adding `value` to `added`/`removed` string branches may flip passing tests. Run full suite after task 4; update expected shapes if any assert `value is None` / key-absent. |

---

## 9 — Test plan

**Unit tests** (in `tests/test_xlsx_structural_diff.py`):
- `test_canon_numfmt_case_and_escape` — `"mmm\\ yyyy"` and `"MMM YYYY"` canonicalize equal.
- `test_canon_numfmt_preserves_non_equivalence` — `"0.00"` ≠ `"0.0000"` after canon.
- `test_canon_argb_alpha_opacity` — `"FF0000FF"` and `"000000FF"` canonicalize equal.
- `test_canon_argb_preserves_non_alpha` — `"FF0000FF"` ≠ `"FF000080"`.

**Integration tests** (two-workbook fixtures):
- `test_format_noise_suppressed_by_default` — workbooks differ only in format casing; default diff emits no cell records, `noise_filtered_summary.number_format_normalized == N`.
- `test_format_noise_kept_with_flag` — same pair with `--keep-noise`; records present with `noise_class: "format"`.
- `test_color_noise_suppressed_by_default` — workbooks differ only in tab-color alpha byte; default diff drops; counter reflects.
- `test_color_noise_kept_with_flag` — same pair with `--keep-noise`; `noise_class: "color"`.
- `test_noise_filtered_summary_always_present` — workbooks with zero noise still emit `noise_filtered_summary: {number_format_normalized: 0, tab_color_alpha_normalized: 0}`.
- `test_n4_removed_string_preserves_value` — §6.3 direction a→b.
- `test_n4_added_string_preserves_value` — §6.3 direction b→a.
- `test_schema_version_is_2` — top-level output carries `"schema_version": 2`.

**Pre-existing suite:** `pytest tests/test_xlsx_structural_diff.py -q` must pass. If any pre-existing test asserts "value not present" on an `added`/`removed` string branch, update the assertion (that was the buggy shape).

---

## 10 — Out of scope (deferred to Stage 2 or future designs)

- B1 / B2 structural-shift detection (Stage 2, sibling doc).
- B3 table-rename detection (deferred — no observed instances).
- `--strict-format` / `--strict-color` opt-out flags (deferred — ship only if a real consumer asks).
- Normalization for fill/font/border ARGB (N3 — not investigated yet; Stage 2 opens a brainstorm if `--keep-noise` output shows this is still a noise source after Stage 1).
- Reader-side canonicalization (Q3 rejected comparator-side as the baseline).
- Schema v1 back-compat shim for old consumers (only caller is `diff.bat`; not needed).

---

## 11 — Acceptance criteria

1. `pytest tests/test_xlsx_structural_diff.py -q` → 34 passing (25 existing + 9 new).
2. `pytest tests/ -q` → full suite green.
3. `pyright scripts/xlsx_structural_diff.py tests/test_xlsx_structural_diff.py` → no errors.
4. Running the script on the Aiko-port fixture pair (`diff/budget_plan_04.22_Aiko.xlsx` vs `diff/budget_plan_reb.xlsx`) produces output where `noise_filtered_summary.number_format_normalized` ≥ 2000 and `noise_filtered_summary.tab_color_alpha_normalized` ≥ 8 (matches seed observations).
5. Same fixture pair with `--keep-noise` produces output whose record count ≈ pre-Stage-1 record count, each suppressed record carrying `noise_class`.
6. Module docstring documents v1 → v2 schema delta.

---

## 12 — Post-ship follow-ups (not gating)

- Archive this plan + the Stage 2 plan to `docs/plans/archived/` once both stages ship (per `plan-hygiene.md`).
- If Stage 1 output reveals N3 (fill/font/border style noise) is still dominant, open a Stage 3 brainstorm on that.
- Update seed doc `2026-04-23-diff-noise-suppression-ideas.md` with a "SUPERSEDED BY" pointer to this file before archiving.
