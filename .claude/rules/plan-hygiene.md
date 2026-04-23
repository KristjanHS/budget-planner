---
paths:
  - "docs/**/*.md"
  - "**/CLAUDE.md"
  - "docs/plans/**/*.md"
  - ".claude/rules/**/*.md"
last_verified: 2026-04-22
---
# Plan / Doc Hygiene

## Plan archival
Shipped plans left in `docs/plans/` root auto-load into future sessions and masquerade as active work. When `/impag` (or any plan-survey step) finds a plan fully shipped, `git mv` it into `docs/plans/archived/` the same session — don't just note shipped status and move on.
Before `git mv`: grep the plan path across `budget_planner/**`, `tests/**`, `docs/**`, `CLAUDE.md` and update references to `docs/plans/archived/<filename>.md`. Verify `.claudeignore` names `archived/` (this project's convention, not `archive/`). If a sibling plan references the archived one with conditional/status language (`"Doc A pending"`, `"if X ships first"`, `"either can ship first"`), the condition has resolved — resolve sibling status notes and post-archive diff-list section headers to state the realized order concretely, not as a hypothetical.
Before `git mv`ing a design doc to `archived/`, grep §Design content against `docs/internals.md`, `docs/design-invariants.md`, `docs/grant-domain.md`. Any load-bearing fact not present in a living doc is merge-debt — migrate first. Do NOT archive plans marked "approved but not implemented" — those are still active even if dormant.
If the plan has working-tree amendments (e.g. a §0 resolution note added in the archiving session), `git add <plan>` the amendment BEFORE `git mv` — `git mv` stages the rename against HEAD content, silently dropping uncommitted working-tree edits. The archive commit then ships without the amendment and the audit-trail content is lost. *Why this rule exists:* 2026-04-23 solver-freeze-point archive `dec2d98` shipped without the §0 code-review follow-up catalog I'd just written; needed a follow-up `cde1475` to restore it.
Multi-stage plans split by stage: one file per stage under `docs/plans/`, each sized to fit a 200k-token session context. **Why:** long single-file plans bloat every auto-load + every `/impag` subagent dispatch — a 500-line plan covering stages A/B/C wastes tokens in every stage-B session re-reading already-shipped stage-A content and stage-C spec. **How to apply:** when `mybrain` or a direct design session spans ≥2 stages, write `<date>-<feature>-stage-A.md`, `-stage-B.md`, … plus a short `<date>-<feature>-overview.md` index if useful; cross-reference between files. Single-stage plans stay in one file (Design + Plan + Corrections sections in place, as before). Archive rule (`git mv` to `archived/` after ship) applies per-stage file. No `-v2-` / `-enhancements-` version siblings — amend in place with a changelog entry inside the same stage file.

## Doc hygiene
Session-generated audits, analyses, and reports go in `docs/analysis/` (already in `.claudeignore` — keeps them out of auto-loaded context). The `docs/` root is for trigger-indexed reference docs listed in CLAUDE.md. Don't drop one-off session outputs into the reference tier.
Bucket-migration tasks (moving content out of a sheet/module): before committing, grep the old identifier across `docs/` (including `docs/plans/*.md` — non-archived plan files gate paused-work resumption and will mislead the next implementing agent), `CLAUDE.md`, and `.claude/rules/`. Additive migrations leave stale prose in downstream docs the migrator never opened.
vN → vN+1 design migrations: grep both the identifier AND the *human-readable feature names* that the old design used (e.g. "Proposed grid", "Assigned grid", "Window months remaining"). Identifier-level sweeps come up clean while stale wording survives in comments, docstrings, Mermaid string bodies, and architecture Notes.
When shifting documented row positions on a sheet (Config param rows, WP section, etc.), grep living docs (`budget_planner/layout.py` + any doc citing row numbers in `docs/design-invariants.md`) and test docstrings citing specific rows/counts — assertion-level edits don't reach prose.
Docs-only plans still require a pytest gate. `tests/test_config_sheet.py::test_config_named_range_targets_match_internals_doc` parses the Cross-Sheet Reference Map table in `docs/internals.md` for `` | `name` | Config!$B$N | Named range | `` rows — restructuring that table breaks pytest even though no code changed. Before `git mv`ing any plan whose diff is `docs/**`-only, run `pytest tests/ -q`; the impag skill's "skip pytest for research-style plans" carve-out assumes no docs→test coupling, which this project violates.

## Re-work audit scope isn't gated by "phased plan"

Any design doc that lands a long ordered step list (≥10 steps, even a single-PR rollout) has the same cross-step re-write risk as a multi-phase plan. Walk each step's concrete edits and tag which file/function each touches. If the same file appears in two sequential steps where the first's change depends on state the second removes, merge them into one per-file pass. **Trigger is "same target, two steps," not "≥3 phases."** The mybrain skill's Step 7 item 5 phase-gate is a floor, not a ceiling.

*Why this rule exists:* 2026-04-19 WP gap removal design §12 — initial draft split `timelog_plan.py` and `capacity.py` each across two steps where step N dropped the `wp_schedule` parameter and step N+2 rewrote the helper that iterated it. Either order left the tree non-compiling between steps until merged per-file.

## Grep user-facing prose when designing a feature removal

Removals compile fine but leave contradictory strings behind — `raise ValueError("use X for ...")` where X is being deleted, docstrings describing the removed shape, data-validation list literals naming the removed option. These survive the refactor and actively mislead users at runtime. Before finalizing a removal design, grep:

- `raise.*Error.*"`
- `"""`-led docstrings
- `\.value\s*=\s*"` string literals

…for the thing being removed; rewrite or delete each in the same PR.

*Why this rule exists:* 2026-04-19 WP gap removal — `workbook.py:157-166` validator directed users at WP SCHEDULE, the section being deleted. Caught during design only because the file was opened for an unrelated reason.
