---
paths:
  - "docs/**/*.md"
  - "**/CLAUDE.md"
  - "docs/plans/**/*.md"
  - ".claude/rules/**/*.md"
last_verified: 2026-04-24
---
# Plan / Doc Hygiene

## When to write a plan doc
- No plan doc for features ≤ 2 files changed AND ≤ 1 session of work. Commit message + code diff is the record.
- Single plan doc (not multi-stage split) for features ≥ 2 sessions or ≥ 3 files.
- Multi-stage split (one file per stage) only after ≥ 2 stages have already shipped and a handoff is actually needed. Do not split speculatively at design time.

## Plan archival
- Plans representing a single shipped session (≤ 1 day end-to-end, all commits reference the plan path or feature name): `git rm`, don't `git mv`. Git log is the audit trail; the archived md is a second copy. Archive only plans whose design rationale isn't fully captured in commit messages or in `docs/spec.md`.

Shipped plans left in `docs/plans/` root auto-load into future sessions and masquerade as active work. When `/impag` (or any plan-survey step) finds a plan fully shipped, `git mv` it into `docs/plans/archived/` the same session — don't just note shipped status and move on.
Before `git mv`: grep the plan path across `budget_planner/**`, `tests/**`, `docs/**`, `CLAUDE.md` and update references to `docs/plans/archived/<filename>.md`. Verify `.claudeignore` names `archived/` (this project's convention, not `archive/`). If a sibling plan references the archived one with conditional/status language (`"Doc A pending"`, `"if X ships first"`, `"either can ship first"`), the condition has resolved — resolve sibling status notes and post-archive diff-list section headers to state the realized order concretely, not as a hypothetical.
Before `git mv`ing a design doc to `archived/`, grep §Design content against `docs/spec.md`. Any load-bearing fact not present is merge-debt — migrate first. **If the fact is already present in `docs/spec.md`, do not duplicate elsewhere; verify and link via `→ see docs/spec.md § X`.** Do NOT archive plans marked "approved but not implemented" — those are still active even if dormant.
If the plan has working-tree amendments (e.g. a §0 resolution note added in the archiving session), `git add <plan>` the amendment BEFORE `git mv` — `git mv` stages the rename against HEAD content, silently dropping uncommitted working-tree edits. The archive commit then ships without the amendment and the audit-trail content is lost.
When a multi-stage split IS warranted (gate in `## When to write a plan doc`): one file per stage under `docs/plans/`, each sized to fit a 200k-token session context. Write `<date>-<feature>-stage-A.md`, `-stage-B.md`, … plus a short `<date>-<feature>-overview.md` index if useful; cross-reference between files. Single-stage plans stay in one file (Design + Plan + Corrections sections in place, as before). Archive rule (`git mv` to `archived/` after ship) applies per-stage file. No `-v2-` / `-enhancements-` version siblings — amend in place with a changelog entry inside the same stage file.

## Doc hygiene
Session-generated audits, analyses, and reports go in `docs/analysis/` (already in `.claudeignore` — keeps them out of auto-loaded context). The `docs/` root is for trigger-indexed reference docs listed in CLAUDE.md. Don't drop one-off session outputs into the reference tier.
Bucket-migration tasks (moving content out of a sheet/module): before committing, grep the old identifier across `docs/` (including `docs/plans/*.md` — non-archived plan files gate paused-work resumption and will mislead the next implementing agent), `CLAUDE.md`, and `.claude/rules/`. Additive migrations leave stale prose in downstream docs the migrator never opened.
vN → vN+1 design migrations: grep both the identifier AND the *human-readable feature names* that the old design used (e.g. "Proposed grid", "Assigned grid", "Window months remaining"). Identifier-level sweeps come up clean while stale wording survives in comments, docstrings, Mermaid string bodies, and architecture Notes.
When shifting documented row positions on a sheet (Config param rows, WP section, etc.), grep living docs (`budget_planner/layout.py` + any doc citing row numbers in `docs/spec.md`) and test docstrings citing specific rows/counts — assertion-level edits don't reach prose.
Col-xref flips on shared surfaces (`Config!$E` → `Config!$F`, `cell_ref(col=5)` → `col=6`): grep non-archived `docs/plans/*.md` + `docs/spec.md` + `.claude/rules/` for old col-letter AND old semantic name ("Budget -i" when flipping to +i); file Corrections to stale sibling plans in the same commit.
Docs-only plans still require a pytest gate. `tests/test_config_sheet.py::test_config_named_range_targets_match_spec_doc` parses the named-range table from `docs/spec.md` for `` | `name` | Config!$B$N | Named range | `` rows; restructuring that table breaks pytest even though no code changed. Before `git mv`ing any plan whose diff is `docs/**`-only, run `pytest tests/ -q`; the impag skill's "skip pytest for research-style plans" carve-out assumes no docs→test coupling, which this project violates.

## Re-work audit scope isn't gated by "phased plan"

Any design doc that lands a long ordered step list (≥10 steps, even a single-PR rollout) has the same cross-step re-write risk as a multi-phase plan. Walk each step's concrete edits and tag which file/function each touches. If the same file appears in two sequential steps where the first's change depends on state the second removes, merge them into one per-file pass. **Trigger is "same target, two steps," not "≥3 phases."** The mybrain skill's Step 7 item 5 phase-gate is a floor, not a ceiling.

## Grep user-facing prose when designing a feature removal

Removals compile fine but leave contradictory strings behind — `raise ValueError("use X for ...")` where X is being deleted, docstrings describing the removed shape, data-validation list literals naming the removed option. These survive the refactor and actively mislead users at runtime. Before finalizing a removal design, grep:

- `raise.*Error.*"`
- `"""`-led docstrings
- `\.value\s*=\s*"` string literals

…for the thing being removed; rewrite or delete each in the same PR.
