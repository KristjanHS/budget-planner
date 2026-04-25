# Doc-bloat root causes — brainstorm starter

**Date**: 2026-04-25
**Status**: STARTER — fresh-session brainstorm seed; no decisions yet
**Constraint from user**: avoid creating unnecessary new rules / skill instructions that themselves bloat context. Solutions that *delete* or *reframe* existing rules are preferred over solutions that *add* new ones.

---

## Context

Just completed a deduplication pass on `docs/*.md` (Bundles A + B, ~80 lines saved across 5 files). The dedup was needed because the same prose accreted across 3-4 docs over multiple plan-archive cycles. Bundle C was deferred for a fresh-session brainstorm — fixing the *root causes* so the duplication doesn't return in 2 months.

A follow-up brutal-honesty audit on the same 5 docs argued the structure itself is wrong (5 files with overlapping responsibilities and no precedence). That audit's three-level action ladder is at the bottom of this file as related context — but **this brainstorm is about root causes / process, not file restructuring**. Restructuring is downstream of fixing the bloat-generation mechanism.

---

## Symptoms observed

**S1.** Indirect-view scope paragraph copy-pasted verbatim across 3 docs (`grant-domain.md`, `design-invariants.md`, `internals.md`) until the dedup pass.

**S2.** `+i`/`-i` storage convention told 4 different ways across 2 files; tab colors duplicated between architecture.md prose and internals.md table; cross-sheet DAG rendered as both mermaid (architecture) and table (internals).

**S3.** `design-invariants.md`'s own preamble says "Entries expire. When refactoring removes an invariant, delete the entry. Git log carries the history; this file carries currently-load-bearing specs." Yet pre-dedup the file contained dated stage tags (`shipped 2026-04-23 as Aiko-port T2`, `(2026-04-19 v2 redesign)`, `Stage 4 (T14–T15) will harden…`). **The file's own author-time rule was ignored at every later edit.**

**S4.** Plan-archive cycles add prose. Per `.claude/rules/plan-hygiene.md`:
> "Before `git mv`ing a design doc to `archived/`, grep §Design content against `docs/internals.md`, `docs/design-invariants.md`, `docs/grant-domain.md`. Any load-bearing fact not present in a living doc is merge-debt — migrate first."

This is unilaterally additive. There's no counterpart "if the fact is *already* in living doc X, don't add a second copy to Y" rule.

**S5.** CLAUDE.md Reference Index has 8 trigger rows; **4 rows point at content that is duplicated across the docs they reference**. Multiple triggers can land on the same fact via different docs, encouraging writers to add the fact to all candidate destinations.

**S6.** The `condense` skill globs `**/CLAUDE.md` only — project docs (`docs/*.md`) have no periodic-dedup mechanism. They accrete forever unless manually invoked.

---

## Hypothesized root causes

### R1. No precedence rule between living docs
Each fact has multiple plausible homes (e.g., indirect-view rules belong to both `grant-domain.md` AND `design-invariants.md`). Without a precedence rule, the "thorough" answer is to put it everywhere. The "thorough" answer is wrong but isn't penalized.

### R2. Plan-archive migration is one-way
`plan-hygiene.md` mandates *moving* content into living docs before archive. There's no symmetric "before adding, dedupe" gate. Net flow: plans → docs (always positive). Living-doc shrinkage requires a separate explicit pass (this one).

### R3. Stage tags are sticky
`(2026-04-19 v2 redesign)`, `Aiko-port T2`, `Phase 3-4 totals rows` — these get embedded as parentheticals during a rollout because they're meaningful *during* the rollout. They survive into living docs forever because no rule says "strip stage tags after ship." `design-invariants.md` preamble *implies* the rule but doesn't enforce it on edit.

### R4. The "delete on absorb" iron rule covers plan→living, not living→living
Global rule: "Superseded docs: consume before archive — fully absorb the old doc's load-bearing content into the replacement *before* moving the old to archive." This is the correct shape but only fires on plan archival. When two living docs accrete overlapping content, no rule activates.

### R5. Reference Index design encourages spread
CLAUDE.md's Reference Index pattern is "trigger → which doc to read." With overlapping topics (indirect-view, vacation, timelog) the natural extension is "trigger → which doc**s**" — and from the writer's side, "if my topic spans triggers, write to all matching docs." The index design accidentally rewards spread.

### R6. The "always emit pointers" hygiene isn't a rule anywhere
After a doc grows large, the right pattern is `## Topic → see X` (one-line pointer). This pattern works (used in the dedup pass). But no rule tells writers to default to pointer-on-write; they default to inline-on-write because pointers feel less complete.

### R7. No "doc fits in head" budget
Files grow monotonically because there's no per-file or aggregate ceiling. `design-invariants.md` is the densest doc at 132L and stays useful; `grant-domain.md` at 332L is mostly waste. The difference is content discipline, not file size — but a soft budget would force the discipline conversation.

---

## Solution shapes — brainstorm prompts (for fresh session)

Constraint reminder: prefer solutions that *delete*, *reframe*, or *consolidate* existing rules over new rules. Each shape below is a starting point, not a recommendation.

### Shape 1: Eliminate the precedence problem by collapsing files
If there are only 2 docs (`spec.md` + `runbook.md`), there's no "which doc owns indirect-view?" question. Bundle C's R1/R5 disappear. R2/R4 still need addressing for plan-archive flow. Closest match to brutal-honesty audit's Merger A.

### Shape 2: Add a single precedence rule, no new file
Single line in CLAUDE.md: "Each topic has exactly one canonical living doc; others link via `→ see X §Y`. When triggers overlap, the deepest doc wins (design-invariants > internals > architecture > grant-domain)." Costs ~2 lines of CLAUDE.md context, addresses R1 + R5. Doesn't address R2/R3.

### Shape 3: Convert `plan-hygiene.md` migration rule from "additive" to "delta-aware"
Edit the existing migration paragraph: "If the fact is already present in living doc X, do not duplicate to doc Y; verify and link. If it's absent, add to the *single* canonical doc per the precedence rule, not all candidates." No new rule file. Addresses R2.

### Shape 4: Strip-on-archive checklist for stage tags
One line added to plan-archive flow in `plan-hygiene.md`: "Before merging plan content into living docs, strip dated stage tags (`(2026-04-23 as Aiko-port T2)`, `Stage N (TX-TY)`, version-suffix parentheticals). The git log carries those." Addresses R3.

### Shape 5: Make the dedup pass a recurring action, not a one-off
Either: extend the `condense` skill's Phase 1 globs to include `<project>/docs/*.md`; OR add a brief checklist line to `/retro` skill ("if `docs/*.md` haven't been condensed in N sessions, run a dedup pass"). Addresses R6 by making periodic-dedup the default. Cost: ~3-line edit to one existing skill, no new skill.

### Shape 6: Per-file purpose declaration at top of each doc
Each living doc opens with a 1-sentence "this file owns X; if your fact isn't X, it doesn't belong here." Forces writer-time precedence check. Cost: 5 lines × N files. Addresses R1 at write-time. Requires Shape 1 or Shape 2 to define the precedence first.

### Shape 7: Delete `architecture.md` and `grant-domain.md` outright; the structure problem dissolves
Brutal-honesty audit's Merger A. ~600 lines deleted from agent context. Addresses R1/R2/R5/R7 by making them inapplicable. Heaviest action; surfaces "but where do diagrams live?" — answer: code docstrings or `docs/human/` (already `.claudeignore`'d).

### Shape 8: Audit the Reference Index for trigger-to-doc cardinality
Walk every trigger row in CLAUDE.md's Reference Index. If a trigger maps to ≥2 docs, that's the duplication signal. Fix by either (a) collapsing the docs, (b) splitting the trigger into more specific triggers, or (c) picking one canonical doc and deleting the duplicate. Diagnostic, not prescriptive — surfaces which R5 instances are real.

---

## Discarded directions (already considered, not promising)

- **D1. New rule file `.claude/rules/doc-hygiene.md`.** Violates user's "no new rules" constraint. Also ironic: a new rule about avoiding doc bloat would itself add context cost.
- **D2. Per-doc line-count CI gate.** Mechanical; doesn't catch low-value content under the cap. Penalizes density.
- **D3. Pre-commit hook that diffs `docs/*.md` and warns on growth.** Annoying, false-positive heavy, doesn't catch the cross-file duplication that's the actual problem.

---

## Pre-flight notes for fresh-session brainstormer

- All files in scope are in context size: 5 docs total ≤ 1028 lines after Bundles A+B. Read all 5 fully before brainstorming — partial reads will miss the cross-file overlap.
- The dedup pass already executed is in `git log` ~2026-04-25; check the diffs to see what *kind* of duplication was hit. Pattern matters more than instances.
- Bundle A + B writeup lives in conversation history of the prior session; this file is the carry-over.
- The brutal-honesty audit's three-level ladder (delete diagrams / cut grant-domain to 50L / collapse to 2 files) is *adjacent* to Bundle C but not the same question. Bundle C is "fix the generation mechanism." Restructuring is "fix the current state." Both are valid; sequence matters (fixing the mechanism first means restructuring doesn't have to be redone in 6 months).

---

## Decision questions for fresh session

1. Which root causes (R1-R7) are real and which are over-fitting to recent symptoms?
2. Of the 8 solution shapes, which 1-2 give the most leverage with the least added context cost?
3. Should restructuring (brutal-honesty Merger A) be sequenced *before* or *after* the root-cause fix? Argument for before: smaller surface to apply rules to. Argument for after: rules should drive the restructuring, not be retrofitted to it.
4. Is the `condense` skill the right home for periodic project-doc dedup, or does that responsibility belong to `/retro`?
