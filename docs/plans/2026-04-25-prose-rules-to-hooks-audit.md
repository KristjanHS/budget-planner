# Prose-rules-to-hooks audit — plan

**Date**: 2026-04-25
**Status**: AUDIT — survey complete; awaiting decision on Wave-1 scope
**Predecessor**: `docs/plans/2026-04-25-doc-bloat-root-causes-brainstorm.md` (docs-bloat-gate design)
**Lens (carry-over from predecessor)**: rule-level guards under-perform when the executing skill doesn't load the rule. Hook-level gates fire regardless of skill — the right shape for write-time-dependent enforcement.

---

## Goal

For each prose rule in `.claude/rules/*.md`, CLAUDE.md, and selected skill bodies, determine whether it would convert cleanly to a deterministic `PreToolUse` (or related) hook gate — and rank by leverage. Output is a prioritized hook backlog, not implementation.

---

## "Hook-convertible" definition (gate)

A rule converts cleanly iff ALL three hold:

1. **Deterministic check** — line-count delta, regex match, file existence, path glob, or similar mechanical predicate. Not a judgment call about "load-bearing-ness" or "appropriate density."
2. **Path-or-tool-scoped** — fires on a known set of (tool, path-pattern) combinations.
3. **Mid-write enforcement** — the rule's purpose is to gate something Claude is about to write, not to inform interpretation of code being read.

Non-convertible categories (caught early so they don't pollute the backlog):
- Read-time guidance (e.g., "scan whole file before claiming refactor done").
- Cross-file semantic mismatch (e.g., "rule examples must match shipped code").
- Test-writing *patterns* (e.g., "use bystander assertions in pruning tests").
- Meta-level Claude behavior (e.g., "batch CLAUDE.md edits to session boundaries").

---

## Top 5 hook candidates (Wave 1 backlog)

Ranked by leverage = (fire frequency) × (current under-enforcement evidence) × (deletable prose lines).

### W1.1 — Session-generated audits gated to `docs/analysis/`
**Source:** `plan-hygiene.md:26` — *"Session-generated audits, analyses, and reports go in `docs/analysis/` (already in `.claudeignore` — keeps them out of auto-loaded context)."*
**Gate:** PreToolUse `Write` on `docs/*.md` (root only); regex scan content for `audit|analysis|report|research` keywords + commit-message skill name (`mybrain`/`brutal-honesty-review`/`senior-architect`/`deep-research`/`kaizen:*`); reject unless path is `docs/analysis/` or `.claudeignore`'d.
**Determinism:** H. **Failure cost:** auto-loaded prose accretion.
**Why now:** **Overlaps with docs-bloat-gate design's §2 scope.** Either fold into that hook (preferred, single hook covers both surfaces) or ship as sibling. Decision in §Open questions below.

### W1.2 — Docs-only plan archive blocked unless pytest passes
**Source:** `plan-hygiene.md:31` — *"Docs-only plans still require a pytest gate. … Before `git mv`ing any plan whose diff is `docs/**`-only, run `pytest tests/ -q`."*
**Gate:** PreToolUse on `Bash` matching `git mv .*docs/plans/.*archived/`; require recent passing pytest exit-code stamp (or run pytest inline if stamp absent/stale).
**Determinism:** H. **Evidence of under-enforcement:** the rule exists *because* a docs-only plan once shipped that broke `test_config_named_range_targets_match_spec_doc`. The rule still relies on Claude reading `plan-hygiene.md` mid-archive; the test couples to spec.md but skills don't load the rule by default.
**Why now:** known prior incident; mechanical fix.

### W1.3 — No plan doc for ≤2-file / ≤1-session work
**Source:** `plan-hygiene.md:11-13` — *"No plan doc for features ≤ 2 files changed AND ≤ 1 session of work. Commit message + code diff is the record."*
**Gate:** PreToolUse `Write` on `docs/plans/*.md` (new file creation); count files touched in current session via `git diff --name-only HEAD` (or session-start ref); reject if ≤2 non-doc files with `<!-- plan-required: <reason> -->` override sentinel available.
**Determinism:** M (session boundary detection is fuzzy).
**Why now:** ~50% of recent plan files in `docs/plans/` would have failed this gate. Highest *prevented* bloat per fire.
**Risk:** session-boundary heuristic needs tuning; may need refinement after Wave 1 telemetry.

### W1.4 — External-process tests go in manual tier
**Source:** `~/.claude/rules/testing.md:32` — *"External-process tests (docker, kubectl, headless browser) go in manual tier from day one."*
**Gate:** PreToolUse `Write` on `tests/**/*.py` (excluding `tests/manual/`); regex scan body for `docker|kubectl|subprocess\.run|playwright|selenium|requests\.(get|post)`; reject unless path is `tests/manual/` or override sentinel present.
**Determinism:** H.
**Why now:** observed 3-5× annual false-green slow-suite tests in non-programme repos; cheap to ship; transferable to other projects.
**Cross-project:** this rule is global, not project-specific. Implementation should live in `~/.claude/hooks/` or a portable hook framework.

### W1.5 — Unused-import pre-declaration triggers F401 echo
**Source:** `~/.claude/rules/python-refactors.md:42-45` — *"Pre-declared imports must be used in same file or marked noqa."*
**Gate:** PreToolUse `Edit`/`Write` on `**/*.py`; AST or regex check for `from X import Y` where `Y` doesn't appear elsewhere in file content AND no `# noqa: F401` / `_ = Y` aliasing; reject.
**Determinism:** M (regex misses some patterns; AST is heavier).
**Why now:** PostToolUse F401-echo spam costs 5-10× message-noise per occurrence (~1-2× per 3-session refactor cycle). Shifts the catch from PostToolUse echo to PreToolUse block.

---

## Wave 2 (lower priority, but still hookable)

| # | Rule | Source | Determinism | Reason for Wave 2 |
|---|------|--------|:---:|---|
| W2.1 | One plan file per stage; no `-v2-` / `-enhancements-` siblings | plan-hygiene.md:23 | H | Low fire rate (1× observed) |
| W2.2 | LO `_xlfn.` prefix for MINIFS/MAXIFS/IFS/SWITCH/TEXTJOIN/CONCAT | openpyxl.md:26 | H | Project-scoped, ~2-3×/year |
| W2.3 | Named-range defs with hyphenated/spaced sheet names quoted | openpyxl-infra.md:12 | H | 1-2× per refactor; project-scoped |
| W2.4 | EUR money leaf formulas wrap math in `ROUND(..., 0)` | openpyxl.md:27 | H | Very narrow scope; low aggregate fire |
| W2.5 | Constant-rename shared-suffix aliasing detector | python-refactors.md:22-25 | M | Regex-fragile; warn-only, not block |

Wave 2 candidates ship as a single multi-pattern hook later, after Wave 1 telemetry confirms the gate-design pattern works.

---

## Rules that look hookable but aren't

For the audit record — these were considered and rejected:

- **Font-color convention** (`openpyxl.md:30`) — requires cross-file semantic comparison between `docs/spec.md` and builder code. No single tool gate.
- **Rule examples copy shipped form verbatim** (`rule-authoring.md:7`) — mid-read concern; needs human review.
- **Lazy-imports whole-file scan** (`python-refactors.md:11-13`) — reading guidance, not write gate. Cannot enforce a search strategy.
- **Test bystander assertion requirements** (`testing.md:42-46`) — semantic test-writing pattern; pattern match on test names is fragile.
- **Critical Rule 7 clause (a) "artifact survives past session"** — semantic. Clause (b) "no existing doc covers" is mechanical and IS captured by W1.1.

---

## Surprising findings (pattern-level)

1. **`plan-hygiene.md` is the hookable hotspot.** 5 of 8 distinct clauses (62%) qualify. The file has accumulated process rules over 6 months; almost all are mechanical-by-design but live as prose because we lacked a hook framework.
2. **Testing rules are fundamentally not hookable.** They're about test-writing *patterns* (intent, what's being asserted), not output shape. Better as a `/test-review` checklist skill than as gates.
3. **Python-refactor rules split mechanical/semantic.** Mechanical ones (import pre-declaration, shared-suffix renames) hookable; semantic ones (lazy-import strategy, optional-param fallback relocation) not.
4. **Global rules are mostly read-time or meta-level.** `claude-md-edits.md`, `hooks.md`, `reading-large-files.md` — they inform how Claude works, not what it outputs. Not hookable.
5. **Verb-asymmetry from prior session shows up here too.** Almost every hookable rule in the audit is *prohibitive* (block this output). Only one (W1.2 pytest gate) is *additive* (require this action). Hooks naturally enforce subtractive verbs better.

---

## Recommended Wave-1 scope

**Ship together** (single PR, mirroring docs-bloat-gate rollout pattern):

1. **W1.1 absorbed into the existing docs-bloat-gate design.** Edit predecessor design §2 (Gate scope) to add: *"reject `Write` to `docs/*.md` root if content contains audit/analysis keywords AND skill emitter is in markdown-emitting list, unless path is `docs/analysis/`."* No second hook; one mechanism, one config.
2. **W1.2 (docs-only plan pytest gate)** as a separate `PreToolUse Bash` matcher — different tool surface, different trigger, doesn't fit the docs-bloat-gate shape.
3. **W1.4 (external-process tests)** as a separate hook — global scope, lives in `~/.claude/hooks/`, not project-specific.

**Defer to Wave 2:** W1.3 (no-plan-doc gate) and W1.5 (F401 pre-block). Both have determinism caveats that would benefit from observing W1.1/W1.2/W1.4 behavior in production first.

Total Wave 1 effort estimate: ~3-4h (1 absorbed scope edit + 2 new hook scripts + settings.json wiring + tests).

---

## Open decision questions

1. **W1.1 absorption vs. sibling hook.** Absorption is cleaner (one hook, one config); but the docs-bloat-gate's net-delta semantic doesn't naturally cover the "audit-content keyword" check. Recommendation: absorb, with the audit-keyword rule as an explicit branch of the hook logic. Confirm?
2. **W1.4 (external-process tests) global vs. project-scoped first.** Shipping global means it affects all your projects from day one. Shipping project-scoped means slower iteration but less blast radius. Default: project-scoped first (this repo); promote to global after 2 weeks of clean operation.
3. **Hook implementation language.** Predecessor design assumed Python (`.claude/hooks/docs-bloat-gate.py`). Should W1.2 and W1.4 hooks share a common helper module (path-glob, JSON parsing) to avoid duplication, or stay independent? Recommendation: independent for Wave 1 (each hook ≤80 lines); refactor to shared module only after 3+ hooks ship.
4. **Override-sentinel naming convention.** Predecessor uses `<!-- bloat-gate-override: <reason> -->`. Should other hooks adopt the same sentinel pattern (`<!-- <hook-name>-override: <reason> -->`)? Recommendation: yes; document in `.claude/hooks/README.md` (out of auto-load scope) when ≥2 hooks ship.

---

## Out of scope this plan

- Implementation of the predecessor docs-bloat-gate (covered by predecessor doc).
- Wave 2 hooks (W2.1–W2.5).
- Reflect-on-de-bloat skill update (still tracked as separate follow-up from prior session).
- Refactoring rule-file *content* — this audit categorizes rules; it does not edit them.

---

## Appendix: full audit table (30 rules surveyed)

Lives in conversation history of this session's audit subagent run; not migrated here to keep this plan tight. Re-run audit if needed via subagent prompt: *"Survey rule files (project + global) for hook-convertible prose; output ranked table per ~/projects/budget-planner/docs/plans/2026-04-25-prose-rules-to-hooks-audit.md methodology."*
