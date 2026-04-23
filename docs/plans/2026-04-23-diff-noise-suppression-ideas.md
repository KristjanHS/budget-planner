# Diff Script Noise-Suppression — Ideas Seed

**Status**: SEED for separate brainstorm — not yet designed. Captured during the Aiko-port session on 2026-04-23 to make sure the observations are not lost.
**Source script**: `scripts/xlsx_structural_diff.py` (+ `tests/test_xlsx_structural_diff.py`)
**Triggered by**: `diff/diff_result.json` (220 KB) where ≥99% of reported changes were noise that had to be filtered downstream.

---

## Why this matters

Current behavior: the diff JSON contains every byte-level difference between two `.xlsx` files. The consumer (human or LLM) has to filter noise *post-hoc* to find real changes. With Aiko's diff:

| Reported | Actual signal | Noise % |
|---|---|---|
| `formula_changes: 2444` | ~190 real | ~92% noise |
| `style_changes: 15546` | ~0 real (in this run) | ~100% noise |
| `cells_added: 10` + `cells_removed: 75` | all real (but see N4 — **lossy**) | 0% noise, but values dropped |

A 220 KB JSON collapses to ~20 KB of real signal. Suppressing noise at *generation* time is cheaper than every consumer reinventing the same filter.

**Interpretability is the other half of the story** (2026-04-23 session-2 learning). Even after the consumer filters noise down to ~13 semantic events, if each event shows up as ~20–45 per-cell deltas spread across a band, the reader has to re-infer the structural intent ("this is a 1-row shift, not 75 independent cell changes"). A missing literal value on a single `removed: true` cell can tip the re-inference the wrong way — see N4. **Suppression reduces volume; structural-shift detection + value-preservation fix interpretability. Both matter.**

---

## Observed noise sources (Aiko diff, 2026-04-23)

### N1 — Number-format casing/escaping normalization (2 251 cells)
Excel/LibreOffice and openpyxl serialize the same logical format with different ASCII:

| What openpyxl writes | What Excel/LO round-trips to |
|---|---|
| `mmm\ yyyy` | `MMM YYYY` |
| `\€#,##0` | `€#,##0` |
| `mmm\ yy` (likely) | `MMM YY` |

These are the same format — just different escape and case conventions. Currently flagged as `number_format: {old, new}` deltas on every cell that has any custom format. Since headers/months/€ cells are ubiquitous, this dominates the JSON.

### N2 — Tab-color ARGB alpha-byte normalization (8 sheets)
| `old` (Excel/LO side) | `new` (openpyxl side) |
|---|---|
| `FF0000FF` | `000000FF` |
| `FFC00000` | `00C00000` |
| `FF7030A0` | `007030A0` |

The non-alpha bytes are identical; only the alpha byte (`FF` vs `00`) differs. Excel ignores alpha for tab colors anyway.

### N3 — `style_changes: 15546` (not investigated this session)
Suspected to be a similar story — font/fill/border ARGB alpha differences, or default-vs-explicit style attributes that mean the same thing rendered. Needs spot-investigation in the next brainstorm before designing a suppression rule.

### N4 — Value-serialization gap on `added`/`removed` string cells (CONFIRMED 2026-04-23 session 2)

**Promoted from "suspected" to confirmed — this bug actively caused a misdiagnosis.**

Logic-mode (`--mode=logic`, now default) string-literal cells on the `added: true` / `removed: true` branches emit entries like:
```json
"A30": {"data_type": "s", "kind": "literal", "removed": true}
"B32": {"data_type": "s", "kind": "literal", "added": true}
"A31": {"added": true, "data_type": "s", "kind": "literal", "value": "PER-PERSON SUMMARY"}
```

The third form (with `value`) happens only sometimes — inconsistently. The first two forms lose the literal text entirely. For cells whose *diff intent* is the text content (section titles, headers, labels, legends), losing the value means the diff tells you "something at A30 was removed" without telling you what it said.

**Actual impact during Aiko-port brainstorm (session 1)**: `A30: {removed: true}` (Aiko-only) vs `A31: {added: true, value: "PER-PERSON SUMMARY"}` (rebuild-only) was misread as a **section-title rename** ("PER-WP SUMMARY" → "PER-PERSON SUMMARY"). The real intent was a pure 1-row shift of the same "PER-PERSON SUMMARY" title. This misread survived session 1 and drove a retracted decision cluster (Q-C in `docs/plans/archived/2026-04-23-aiko-port-decisions.md`) in session 2. Cost ~3 turns plus a Correction section to recover.

**Fix surface**: the `added`/`removed` branches need to serialize `value` for string-literal cells the same way the `kind_change` branch serializes `old`/`new`. A numeric cell's value is already lost under most suppression goals (numbers are usually formula results), but string cells carry semantic intent — losing them is a missing-signal bug, not noise.

**Regression fixture when this gets fixed**: create two xlsx files differing only in `A1 = "Foo"` (file_a) vs `A1` empty (file_b). Logic-mode diff must emit `{"removed": true, "data_type": "s", "value": "Foo"}`.

### N5 — Layout-shift cell explosion (NEW, observed 2026-04-23)

A single "row shifted up by 1" or "column band shifted right by 1" generates tens of per-cell diff entries:

| Refactor in current main | Per-cell diff count in `diff_result.json` |
|---|---|
| Person-summary block header row 32→33, data 33–40 → 34–41 (T7.4, 1-row up) | ~75 (cols A–I × rows 30–41) |
| WP summary band Budget H→I, yearly I–M → J–N (T7.2, 1-col right) | ~80 (rows 6–15 × cols B–N) |
| Row 30 monthly-cost row (Aiko additive, not a shift) | 45 (cols A–AS × row 30) |

**Total real intent in this diff**: 3 refactors + 2 additive features + 8 Config formula changes = ~13 semantic events. **Total cell entries emitted**: 277 (Timelog Plan) + 8 (Config) = 285 per-cell diffs. **Signal compression ratio**: ~22:1.

The B-group brainstorm ideas (B1 row-shift, B2 col-shift) aren't theoretical noise-reduction — they'd have collapsed this diff from 285 cell-entries to ~13 structural events, and the misdiagnosis in session 1 would have been impossible (a `block_shifted_up_by_1` event leaves no room for a "title rename" interpretation).

**Priority bump**: B1/B2 was a "nice-to-have reduce volume" idea pre-Aiko-port. Post-Aiko-port, it's an *interpretability* fix. The logic-mode diff is actively misleading when reading layout refactors.

---

## Improvement ideas (to brainstorm)

### Group A — Logical-equivalence normalizers (highest payoff)
- **A1**: Normalize number formats before comparison (`upper()`, strip backslash escapes for known-safe chars `€$%`, drop leading `\`). Both sides go through `_canonicalize_number_format()` → compare canonical strings.
- **A2**: Normalize ARGB color values (drop alpha byte if it's `FF` or `00` — both are "fully opaque", which is the only value Excel uses).
- **A3**: Provide an opt-out flag (`--strict-format`, `--strict-color`) for users who genuinely want the byte-level diff.

### Group B — Structural-shift detection (reduces volume by a different mechanism)
- **B1**: Detect "row inserted/removed" as a structural event, not 44 per-cell deltas. If a contiguous block of cells differ only by a `±N` row offset in their formula references, emit one `block_shifted: {anchor, by_rows}` record instead.
- **B2**: Same for column inserts/removes.
- **B3**: Detect "table renamed" (`old_table → new_table`) as one event when all `tbl_X[…]` refs change to `tbl_Y[…]`.

### Group C — Schema additions for clarity
- **C1**: Add `summary.real_changes` count alongside `formula_changes` so consumers see "190 real / 2254 noise" up-front.
- **C2**: Tag each cell record with `noise_class: "format" | "color" | "structural" | null` so a consumer can filter without re-deriving.
- **C3**: Top-level `noise_filtered_summary: {number_format: 2251, tab_color: 8, ...}` so the suppression is explicit, not silent.

### Group D — Bug fixes (confirmed, P0)
- **D1**: `added: true` AND `removed: true` records must preserve string-literal `value` (see N4 for confirmed real-world misdiagnosis). This is P0 — it caused a retracted decision cluster in the Aiko-port brainstorm. Test fixture: xlsx pair differing only in `A1 = "Foo"` vs `A1` empty; assert the diff entry carries the literal text.
- **D2**: Audit logic-mode filter rules. Whatever logic is deciding "this cell's value is a pure-style change, drop the `value`" is over-aggressive for `added`/`removed` branches. Inspect the branch in `scripts/xlsx_structural_diff.py` that emits `kind: "literal"` without a `value` field.

---

## Open questions for the brainstorm

1. **Suppression vs. tagging** — is the goal to *drop* noise from the JSON, or to *tag* it so consumers can opt in? Tagging is friendlier for debugging the diff tool itself; dropping is friendlier for everyday use. (Hybrid: tag-then-drop based on a CLI flag.)
2. **Backwards compatibility** — does any downstream consumer (CI, code review tooling, the `--mode=logic` callers) parse the current shape? `git log --all -- scripts/xlsx_structural_diff.py` to find callers before changing schema.
3. **Where does normalization belong** — in the reader (one-shot canonicalize on load) or in the comparator (compare with a custom equality predicate)? Reader is simpler; comparator preserves raw data for `--strict` mode.
4. **`schema_version` bump** — current version is `1`. If output shape changes (new keys, dropped keys), bump to `2` and document.
5. **Does this turn into a Stage gate?** — the Timelog Plan / Config refactor work is shipping in stages with code-review. Is the diff tool similarly gated, or can these improvements ship piecewise?

---

## How to start the next brainstorm

1. Read this file + `scripts/xlsx_structural_diff.py` (likely <800 lines, but check first; use `wc -l`).
2. Confirm or refute N3 by grepping the diff output for `"style_changes"` shape and sampling 5–10 records.
3. Pick a scope: A-only (fast win), A+B (medium), A+B+C+D (full overhaul). User probably wants the menu.
4. Then run `/brainstorm` (mybrain skill) on the chosen scope and produce a proper design doc.

---

## Inputs preserved

- `diff/diff_result.json` — the example noisy output. Sparse-read only.
- `diff/budget_plan_04.22_Aiko.xlsx`, `diff/budget_plan_reb.xlsx` — the two inputs the diff was generated from. Useful as a regression fixture once suppression rules land.
