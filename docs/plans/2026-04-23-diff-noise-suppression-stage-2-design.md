# Diff Noise-Suppression — Stage 2 Design (B1 + B2 structural-shift detection)

**Date**: 2026-04-23
**Status**: SKELETON — full design deferred to a follow-up brainstorm once Stage 1 ships
**Seed**: `docs/plans/2026-04-23-diff-noise-suppression-ideas.md` § N5 + Group B
**Stage 1 sibling**: shipped 2026-04-23, commits 6a3162f → 38f920c

---

## Why deferred

Stage 2 is the only algorithmically novel piece of the full overhaul. Resolving it properly (block-detection heuristics, minimum-size thresholds, partial-shift handling, interaction with N4-preserved `value` strings) deserves a dedicated brainstorm session after Stage 1 lands. Stage 1's shipped output will also inform Stage 2's fixture pool — post-Stage-1 logic-mode diffs on real refactors give concrete cases to design against.

---

## Goals (inherited from seed § N5)

1. **B1 — Row-shift detection.** A contiguous block of cells whose only difference is a `±N` row offset in formula refs + cell positions emits a single `block_shifted: {anchor, by_rows, cell_count}` record instead of N per-cell deltas.
2. **B2 — Column-shift detection.** Same, for columns.
3. **Interpretability, not just volume.** Prevent repeats of the 2026-04-23 misdiagnosis (§N4 of seed): `A30 removed` + `A31 added "PER-PERSON SUMMARY"` was mis-read as a title rename. A `block_shifted_up_by_1` event at the block level leaves no room for that misinterpretation.

## Non-goals (Stage 2)

- **B3 (table rename)** — different algorithm (string-pattern match on formula refs), zero observed cases. Deferred to a separate design if ever needed.
- No change to normalization rules shipped in Stage 1.
- No new CLI flags unless detection has a tunable threshold worth exposing.

## Schema changes (additive, still `schema_version: 2`)

- Add top-level `structural_shifts: [ {sheet, anchor, extent, by_rows, by_cols, cell_count} ]` array.
- Add `"structural"` value to `noise_class` taxonomy for per-cell records that were absorbed into a shift event. Under `--keep-noise`, the absorbed records still appear, tagged `noise_class: "structural"` and carrying a back-reference to the shift index.
- Extend `noise_filtered_summary` with `structural_shifts_absorbed: N` counting cells folded into shift records.

## Open design questions (to resolve in the Stage 2 brainstorm)

1. **Detection algorithm.** Contiguous-block scan? Hash-based row-fingerprint matching? Minimum block size to qualify (1 cell is not a "shift")? What if a "shifted" block has one-cell overrides in the middle?
2. **Shift direction ambiguity.** Row 32→33 can be described as "row 32 inserted above" or "rows 33+ shifted down by 1". Pick one canonical description and document.
3. **Interaction with N4 string values.** Does a shifted block still emit preserved string values for its top / bottom rows? Probably yes — the values need to survive the shift record for audit.
4. **Per-sheet vs cross-sheet.** Are shifts only detected within a sheet, or can a block be "moved" between sheets? Seed observations are all within-sheet; start there.
5. **False-positive risk.** If two unrelated rows happen to differ by the same content at an offset, the detector will wrongly merge them. What's the safeguard — minimum block size, content fingerprint, sheet-structure anchor?

## Expected deliverables (when this stage is designed)

- Full detection algorithm spec with pseudocode.
- Schema example with real shift records from the Aiko-port fixture.
- Task breakdown for `/impag`.
- Test plan with both detection-positive and detection-negative fixtures (adjacent changes that look shift-like but aren't).
- Risks table, especially for false-positive collapse.

## How to start the Stage 2 brainstorm

1. Confirm Stage 1 has shipped and the output is stable.
2. Collect 3–5 real refactor fixtures from recent `git log` (T7.2 col shift, T7.4 row shift, and 1–2 more).
3. Read this skeleton + Stage 1 design + seed §N5.
4. Run `/brainstorm` (mybrain skill) on the open-questions list above.
5. Write a full design file at this same path (replacing this skeleton) or append sections below.
