# Doc-bloat root causes — design

**Date**: 2026-04-25
**Status**: DESIGN — converged from /mybrain Deep mode session; ready for implementation plan
**Constraint from user**: solutions that *delete* or *reframe* existing rules preferred over solutions that *add* new ones. Net-new text allowed if leverage justifies (quality-weighted).
**Critical reframe**: all doc writes are performed by Claude under skills/rules — no human writers. The injection point is whichever skill/rule is mid-write.

---

## Goal & metric

**Goal:** Reduce stale/duplicated content auto-loaded into fresh sessions.
**Metric:** total tokens of `docs/spec.md` + `docs/runbook.md` + `docs/refactor-incidents.md` + `.claude/rules/*.md` content auto-loaded per fresh session, sampled monthly.
**Target:** monotonic non-increase except when a new invariant ships (override-logged).

---

## Mechanism map (replaces R1–R7 in starter)

The starter's R1–R7 was symptom-derived from the 5-doc world. After ideation (Inversion + Assumption Challenge + Perspective Multiplication), the live mechanisms are:

| Code | Mechanism | Source |
|------|-----------|--------|
| **M1** | Skills emit doc-prose by default (`mybrain`, `impag`, `senior-architect`, `kaizen:*`, `brutal-honesty-review`) | Inversion #6, surprise #1 |
| **M2** | `impag` dispatches multi-task plans with no cross-task spec.md grep | Inversion surprise #3 |
| **M3** | Stage tags (`(Aiko-port T2)`, `Stage N (TX-TY)`) authored at write-time, not archive-time | Assumption A3 |
| **M4** | Verb asymmetry: enforceable add-verbs, judgment-only subtract-verbs | Perspective Lens 1+2 cross-lens |
| **M5** | No write-time gate firing on `docs/*.md` regardless of skill | Assumption A1 |
| **M6** | Plan→living migration is one-way (R2 from starter, partly addressed by b5db2dd) | Starter R2 |
| **M7** | "2-doc baseline" claim is false; 3 docs + 3 trigger rows still exist | Assumption A5 |

R1 (precedence), R5 (Reference Index spread), R7 (fits-in-head budget) are subsumed by M4/M7 or non-independent.

---

## Chosen approach: Hook-first gate at the write choke point

**Rejected alternatives:**
- **B (Verb-flip rule edits)** — Critical Rule 7 already exists as a skill-output guard *and didn't prevent the recent bloat*. Rule-level gates only fire if the executing skill loads them. Doesn't match the reframed problem (writers are Claude under varying skills).
- **C (Restructure-then-stop)** — Reduces blast radius via M7 finish, but doesn't stop the *generator*. Single-doc bloat still possible.

**Why A:** A `PreToolUse` hook fires regardless of which skill is executing — only solution that matches the all-writers-are-Claude reframe. Constraint-satisfying: hook *replaces* prose rules (deletes them) rather than adding more.

**Coverage:** A hits 6/7 mechanisms (M1, M2, M3, M5, M6 directly; M4 dissolves because the gate IS the subtractive verb). M7 remains out of scope, tracked as follow-up.

---

## §1 — Gate scope

**Paths gated** (exact list, no recursive globs):
- `docs/spec.md`, `docs/runbook.md`, `docs/refactor-incidents.md`
- `.claude/rules/*.md` (project)
- `CLAUDE.md` (project, optional — toggle via env var `DOCS_BLOAT_GATE_CLAUDE_MD=1`)

**Tools gated** (PreToolUse):
- `Write`, `Edit` — proposed content visible directly
- `Bash` — command-regex matcher for `(sed|tee|cat|>|>>)` targeting any gated path

**Out of scope:** `docs/plans/**`, `docs/analysis/**`, `docs/human/**`, `docs/research/**`, all global `~/.claude/`.

## §2 — Gate logic

For each gated tool call:

1. Compute proposed-vs-current diff (line-count delta).
2. **Pass** if any of:
   - net delta ≤ 0
   - `→ see ` present in added text
   - new `^## ` heading introduced (carve-out §4.1)
   - override sentinel present (§3)
3. **Block** otherwise with stderr:
   > `docs-bloat-gate: net +N lines without pointer / removal / new section / override. Add a "→ see X §Y" pointer, pair with a removal, or add <!-- bloat-gate-override: <reason> -->.`

Gate does NOT inspect content quality; only structural shape. Latency target <50ms.

## §3 — Override mechanism + audit log

**Sentinel:** `<!-- bloat-gate-override: <reason ≥ 10 chars> -->` in the added content (Write/Edit) or as a comment in the file post-Bash.

**On override:** hook still runs but exits 0; appends one line to `docs/analysis/bloat-overrides.log`:

```
2026-04-25T14:30:00Z  spec.md  +47 lines  reason="indirect-view §added"  tool=Edit
```

`docs/analysis/` is `.claudeignore`'d, so the log doesn't auto-load. `/retro` skill gains a 1-line check: *"if `bloat-overrides.log` has ≥3 entries since last retro, surface for review."*

## §4 — Carve-outs

1. **New-heading carve-out:** if the diff introduces a brand-new `^## ` heading not present in the on-disk file, all lines under that heading (until next `^## ` or EOF) are exempt from net-delta accounting in the *introducing edit only*. Subsequent edits to that section get full gate treatment.
2. **Reference-Index rows:** trigger-table rows in CLAUDE.md (lines matching `^\| .* \| .*\.md`) count as 1 line regardless of content length — they're indexes, not prose.
3. **Generated content:** anything between `<!-- generated:start -->` and `<!-- generated:end -->` exempt entirely.

## §5 — Companion rule-text deletions (B-sliver)

Delete redundant prose now made obsolete by the hook:

- `plan-hygiene.md`: remove the bolded *"If the fact is already present in `docs/spec.md`, do not duplicate elsewhere…"* line — the hook enforces it.
- `plan-hygiene.md`: shorten the migration paragraph; replace with one line *"Migration adds to `docs/spec.md`; the docs-bloat-gate hook enforces net-non-positive delta unless overridden."*
- CLAUDE.md Critical Rule 7: shorten skill-emit-default paragraph; the hook enforces structural shape, leaving Rule 7 to state intent only.

**Net effect:** ~12–15 lines removed from rule files; ~80 lines added in `.claude/hooks/docs-bloat-gate.py`. Rule context shrinks; hook lives outside auto-loaded scope.

## §6 — Testing & rollout

**Test cases (manual, not pytest):**

1. Edit adds 5-line paragraph, no pointer, no removal → blocked.
2. Edit adds 5 lines + removes 6 → passes (net -1).
3. Edit adds 30 lines under new `## §Indirect-View Math` heading → passes via §4.1.
4. Edit adds 30 lines with `<!-- bloat-gate-override: new EU compliance invariant -->` → passes; appends to log.
5. `Bash` with `echo "..." >> docs/spec.md` → blocked by Bash matcher.
6. Edit to `docs/plans/foo.md` → not gated (out of scope).

**Rollout:** ship hook + companion rule deletions in **one commit** so the rules don't reference a non-existent hook or vice versa. Run for 1 week; review `bloat-overrides.log` to tune carve-outs.

## §7 — Out of scope + deferred follow-ups

**Out of scope this design:**
- **M7 finish** — collapse `runbook.md` + `refactor-incidents.md` into `spec.md` (separate restructuring task).
- `condense` skill scope-extension to `docs/*.md` (the hook makes this less urgent).
- `mybrain` / `impag` design-doc emission discipline (Critical Rule 7 still owns intent layer).

**Deferred follow-ups:**
- Run `reflect` on the `de-bloat` skill to absorb learnings from this session.
- **Audit prose-rules-that-should-be-hooks** — survey `.claude/rules/*.md`, global rules, and skill bodies for instructions whose enforcement depends on Claude reading and obeying them mid-write. Output: table of candidates ranked by leverage if converted to hook gates.

---

## Key design decisions (numbered, for audit trail)

1. **Hook over rule-text** because Critical Rule 7 demonstrates rule-level guards under-perform when the executing skill doesn't load the rule. (Finding 5 from ideation.)
2. **Bash gate included** because Claude has been observed to bypass blocked Write/Edit via `sed -i` / `cat >`. Without Bash coverage, the gate is theatrical. (Pre-mortem failure mode #4.)
3. **`.claude/rules/*.md` in scope from day one** because the deeper goal is auto-loaded-context tokens, not docs-only line count. (Abstraction laddering zoom-out.)
4. **Override sentinel logged, not silent** because invisible overrides degrade silently. The log is `.claudeignore`'d so it doesn't itself bloat context. (Pre-mortem #1.)
5. **New-heading carve-out** because legitimate new sections are 30+ net-positive lines and forcing override-overuse defeats the gate. (Pre-mortem #3, Hidden Assumption H2.)
6. **Single-commit rollout** of hook + rule deletions because partial rollout creates dangling references in either direction.
7. **No pytest gate change** because the hook is mechanical and external; existing `test_config_named_range_targets_match_spec_doc` still covers the spec.md table coupling.

---

## Pre-mortem failure modes (full table)

| # | Failure mode | Likelihood | Impact | Mitigation |
|---|---|:---:|:---:|---|
| 1 | Sentinel-override overuse | High | High | Log every override; surface count in `/retro` |
| 2 | Hook latency hurts workflow | Low | Med | Line-count + grep only; benchmark <50ms |
| 3 | False positives on new sections | Med | High | New-heading carve-out (§4.1) |
| 4 | Bypass via Bash (`sed -i`, `cat >`, `>>`) | Med | **Critical** | Bash PreToolUse matcher (§1) |
| 5 | Pointer regex out of sync with conventions | Low | Med | Liberal pattern: any `→ see ` + filename or `§` |
| 6 | Hook fires on archived plan moves | Low | Low | Path scope is exact-list, not `docs/**` |
| 7 | Bloat relocates to `.claude/rules/*.md` | Med | Med | Gate scope includes `.claude/rules/*.md` |
| 8 | Gate hardens M7's incomplete restructure | Med | Low | Acknowledged; M7 follow-up tracked |
