---
paths:
  - "docs/**/*.md"
  - "**/CLAUDE.md"
  - "docs/plans/**/*.md"
  - ".claude/rules/**/*.md"
---
# Plan / Doc Hygiene

## Plan archival
Shipped plans left in `docs/plans/` root auto-load into future sessions and masquerade as active work. When `/impag` (or any plan-survey step) finds a plan fully shipped, `git mv` it into `docs/plans/archived/` the same session — don't just note shipped status and move on.
Before `git mv`: grep the plan path across `budget_planner/**`, `tests/**`, `docs/**`, `CLAUDE.md` and update references to `docs/plans/archived/<filename>.md`. Verify `.claudeignore` names `archived/` (this project's convention, not `archive/`). If a sibling plan references the archived one with conditional/status language (`"Doc A pending"`, `"if X ships first"`, `"either can ship first"`), the condition has resolved — resolve sibling status notes and post-archive diff-list section headers to state the realized order concretely, not as a hypothetical.
Before `git mv`ing a design doc to `archived/`, grep §Design content against `docs/internals.md`, `docs/design-invariants.md`, `docs/grant-domain.md`. Any load-bearing fact not present in a living doc is merge-debt — migrate first. Do NOT archive plans marked "approved but not implemented" — those are still active even if dormant.
Per-feature design + plan + later phases go in **one** dated file under `docs/plans/`, appended in sections: *Design*, *Plan*, *Corrections*, *Phase N*. No `-v2-` / `-enhancements-` siblings — amend in place with a changelog entry. `mybrain` + `impag` skills both write into this one file; the archive rule (`git mv` after ship) is unchanged.

## Doc hygiene
Session-generated audits, analyses, and reports go in `docs/analysis/` (already in `.claudeignore` — keeps them out of auto-loaded context). The `docs/` root is for trigger-indexed reference docs listed in CLAUDE.md. Don't drop one-off session outputs into the reference tier.
Bucket-migration tasks (moving content out of a sheet/module): before committing, grep the old identifier across `docs/` (including `docs/plans/*.md` — non-archived plan files gate paused-work resumption and will mislead the next implementing agent), `CLAUDE.md`, and `.claude/rules/`. Additive migrations leave stale prose in downstream docs the migrator never opened.
vN → vN+1 design migrations: grep both the identifier AND the *human-readable feature names* that the old design used (e.g. "Proposed grid", "Assigned grid", "Window months remaining"). Identifier-level sweeps come up clean while stale wording survives in comments, docstrings, Mermaid string bodies, and architecture Notes.
When shifting documented row positions on a sheet (Config param rows, WP section, etc.), grep living docs (`budget_planner/layout.py` + any doc citing row numbers in `docs/design-invariants.md`) and test docstrings citing specific rows/counts — assertion-level edits don't reach prose.
Docs-only plans still require a pytest gate. `tests/test_config_sheet.py::test_config_named_range_targets_match_internals_doc` parses the Cross-Sheet Reference Map table in `docs/internals.md` for `` | `name` | Config!$B$N | Named range | `` rows — restructuring that table breaks pytest even though no code changed. Before `git mv`ing any plan whose diff is `docs/**`-only, run `pytest tests/ -q`; the impag skill's "skip pytest for research-style plans" carve-out assumes no docs→test coupling, which this project violates.
