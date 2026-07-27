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

Shipped plans left in `docs/plans/` root auto-load into future sessions and masquerade as active work. When `/impag` (or any plan-survey step) finds a plan fully shipped, `git mv` it into `docs/plans/archive/` the same session — don't just note shipped status and move on.
Before `git mv`: grep the plan path across `budget_planner/**`, `tests/**`, `docs/**`, `CLAUDE.md` and update references to `docs/plans/archive/<filename>.md`. Verify `.claudeignore` names `archive/` — the global convention, adopted here 2026-07-27 (was `archived/`; old refs may survive in `docs/analysis/`). If a sibling plan references the archived one with conditional/status language (`"Doc A pending"`, `"if X ships first"`, `"either can ship first"`), the condition has resolved — resolve sibling status notes and post-archive diff-list section headers to state the realized order concretely, not as a hypothetical.
Before `git mv`ing a design doc to `archive/`, grep §Design content against `docs/spec.md`. Any load-bearing fact not present is merge-debt — migrate first. Walk each §section — vocabulary grep hits upstream §s but misses downstream rendering rationale. Do NOT archive plans marked "approved but not implemented" — those are still active even if dormant. Non-duplication / `→ see X §Y` pointer hygiene is unenforced — no hook checks it; `docs-bloat-gate.py` only emits advisory slop/density nudges and never inspected pointers.
If the plan has working-tree amendments (e.g. a §0 resolution note added in the archiving session), `git add <plan>` the amendment BEFORE `git mv` — `git mv` stages the rename against HEAD content, silently dropping uncommitted working-tree edits. The archive commit then ships without the amendment and the audit-trail content is lost.

**Pathspec BOTH halves of the rename.** `git mv` stages a delete plus an add; a commit pathspec naming only `docs/plans/archive/<stem>.md` commits the add and leaves the delete staged — the doc then exists at both paths and still auto-loads. List the old path too.

**Gate the `git mv` on `git merge-base --is-ancestor <sha> main`.** Archived ≠ landed: a worktree can archive commits that never reached `main`.

**Worktree ships, sync archives.** When the last stage ships from a linked worktree, "same session" is impossible — the commits reach `main` only at `make syncwt`. The shipping session records the plan as archive-pending in `project_state`; the landing session re-checks each landed stem (`git log --grep <stem>` all on `main`, only a USER-owed confirm left) and does the `git mv` there.

## Cite plan docs by stem, not path

Cite by filename stem — `2026-04-25-prose-rules-to-hooks-audit.md`, never the `docs/plans/`-prefixed path. Archiving then changes the path without dangling any ref. If a directory hint helps, point at the directory, and always at `docs/plans/`, never `docs/plans/archive/` — a repoint into the archive re-breaks on the next move. Canonical form: `<stem>.md (docs/plans/)`. Binds prose and code docstrings alike; a script that *reads* a plan file globs the basename across both `docs/plans/` and `docs/plans/archive/` and hardcodes neither.

## Deleting a SECTION HEADING dangles refs exactly like deleting a file

Nothing warns — the file still resolves. When a restructure merges or renames `##` headings, grep inbound section cites and repoint them **in the same commit**; run it repo-wide, `docs/plans/` included (active AND archived).

**Use this form verbatim — the wildcard is load-bearing:**

```
grep -rn "<file-stem>.\{0,3\}§" --include=*.md --include=*.py
```

`.\{0,3\}` must stay a **dot** (any char) and the stem must include `.md`. Real cites read `` `plan-hygiene.md` §Claims `` — a closing backtick sits between the stem and the `§`, so a "tightened" variant like `[^ ]\{0,3\}\s*§` matches nothing and reports a false all-clear.
When a multi-stage split IS warranted (gate in `## When to write a plan doc`): one file per stage under `docs/plans/`, each sized to fit a 200k-token session context. Write `<date>-<feature>-stage-A.md`, `-stage-B.md`, … plus a short `<date>-<feature>-overview.md` index if useful; cross-reference between files. Single-stage plans stay in one file (Design + Plan + Corrections sections in place, as before). Archive rule (`git mv` to `archive/` after ship) applies per-stage file. No `-v2-` / `-enhancements-` version siblings — amend in place with a changelog entry inside the same stage file.

## Doc hygiene
Session-generated audits, analyses, and reports go in `docs/analysis/` (already in `.claudeignore` — keeps them out of auto-loaded context). The `docs/` root is for trigger-indexed reference docs listed in CLAUDE.md. Don't drop one-off session outputs into the reference tier.
Bucket-migration tasks (moving content out of a sheet/module): before committing, grep the old identifier across `docs/` (including `docs/plans/*.md` — non-archived plan files gate paused-work resumption and will mislead the next implementing agent), `CLAUDE.md`, and `.claude/rules/`. Additive migrations leave stale prose in downstream docs the migrator never opened.
vN → vN+1 design migrations: grep both the identifier AND the *human-readable feature names* that the old design used (e.g. "Proposed grid", "Assigned grid", "Window months remaining"). Identifier-level sweeps come up clean while stale wording survives in comments, docstrings, Mermaid string bodies, and architecture Notes.
When shifting documented row positions on a sheet (Config param rows, WP section, etc.), grep living docs (`budget_planner/layout.py` + any doc citing row numbers in `docs/spec.md`) and test docstrings citing specific rows/counts — assertion-level edits don't reach prose.
Col-xref flips on shared surfaces (`Config!$E` → `Config!$F`, `cell_ref(col=5)` → `col=6`): grep non-archived `docs/plans/*.md` + `docs/spec.md` + `.claude/rules/` for old col-letter AND old semantic name ("Budget -i" when flipping to +i); file Corrections to stale sibling plans in the same commit.
Docs-only plans still require a pytest gate. `tests/test_config_sheet.py::test_config_named_range_targets_match_spec_doc` parses the named-range table from `docs/spec.md` for `` | `name` | Config!$B$N | Named range | `` rows; restructuring that table breaks pytest even though no code changed. Before `git mv`ing any plan whose diff is `docs/**`-only, run `pytest tests/ -q`; the impag skill's "skip pytest for research-style plans" carve-out assumes no docs→test coupling, which this project violates.

## Re-work audit scope isn't gated by "phased plan"

Any design doc that lands a long ordered step list (≥10 steps, even a single-PR rollout) has the same cross-step re-write risk as a multi-phase plan. Walk each step's concrete edits and tag which file/function each touches. If the same file appears in two sequential steps where the first's change depends on state the second removes, merge them into one per-file pass. **Trigger is "same target, two steps," not "≥3 phases."** The mybrain skill's Step 7 item 5 phase-gate is a floor, not a ceiling.

Test deletes go where the prod-code removal lands, not where the topic finishes — late-stage test deletes leave earlier stages pytest-broken.

Grep changed layout constants across `tests/` before splitting — siblings hardcoding old values fail the suite even when named test files pass.

## Feature-removal grep

Before finalizing a removal design, sweep `raise.*Error.*"`, `"""` docstrings, and `\.value\s*=\s*"` literals — see `~/.claude/references/pre-ship-sweeps.md` § "Grep user-facing prose when designing a feature removal".
