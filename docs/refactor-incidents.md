---
name: Refactor incident archive
description: Project-specific incidents that drove the global Python/test/hook rules — kept here so the global rules stay generic.
paths:
  - "**/*.py"
  - "tests/**"
  - "**/test_*.py"
  - "**/*_test.py"
  - "**/hooks/*.py"
last_verified: 2026-04-25
---

# Refactor Incident Archive

These incidents originally lived as "Why this rule exists" paragraphs in `~/.claude/rules/` (python-refactors, testing, hooks). Moved here 2026-04-25 because they only make sense inside this project.

The general rules still apply globally; consult those files first. This archive is for "which past incident drove that rule" audits.

## Lazy imports — `another-cli` (2026-03-20)

`another-cli` UV packaging refactor missed 5 lazy imports inside `main()` branches; tests passed because no test exercised those paths. Drove rule: scan whole file for imports, not just the top.

## Tightening optional → required params (2026-04-12, 2026-04-13)

- 2026-04-12: `build_summary_sheet` hardening broke `test_trends.py` via the dual-caller pattern; fix was moving the lazy-load up into `_build_all_sheets`.
- 2026-04-13: audit flagged `build_config_sheet`'s fallbacks as "redundant"; applying it broke ~15 fixtures in `test_config_sheet.py` / `test_team.py` / `test_reader_v2.py`.

Drove rule: dual-caller gotcha — orchestrator + bare test-fixture calls means the fallback is load-bearing for fixtures.

## Constant renames — shared-suffix aliasing (2026-04-13)

Programme Tier A sweep on `summary.py` used `Edit replace_all: true` on `TOTAL_COLS → COLS_PER_YEAR`, silently rewriting `RP_TOTAL_COLS → RP_COLS_PER_YEAR`. Reviewer caught it; took a follow-up commit to fix.

Drove rule: grep `[A-Z_]*OLD_NAME\b` for siblings before `replace_all`.

## Removing exported constants — `tests/**` import sweep (2026-04-19)

Timelog Plan v2 Phase 1b removed `SECTION_TITLE_ASSIGNED` / `SECTION_TITLE_PROPOSED`; plan's test matrix only named `test_timelog_plan.py` for rewriting, but `test_timelog_plan_roundtrip.py` (not listed) still imported both and failed collection on the first full-suite run.

Drove rule: `grep -rn "\bNAME\b" tests/` before removing exported constants.

## Mass kwarg drops — sed/replace_all over per-file Edit (2026-04-20)

Doc B (WP gap-removal) `wp_schedule=` sweep across 11 source files + 13 test files blew the session-stop budget mid-Step 8; the test-file sed pass rewrote 5 files in a single Bash call at the same cost as two Edits.

Drove rule: ≥4 callsites with identical kwarg drop → use sed or single `Edit(replace_all=true)`, not per-site Edits.

## Pre-declared imports — F401 hook echoes (2026-04-24)

Actuals-Hours T6 pre-declared `WITH_INDIRECT_HEADER_FILL` + `WITHOUT_INDIRECT_HEADER_FILL` in the second Edit; F401 fired on 6 subsequent unrelated Edits until the usage Edit cleared it — ~100 lines of duplicated hook context, ~2-3% of session budget burned before budget hook forced wrap-up.

Drove rule: batch new import with its first usage in the same Edit.

## Test bystander — multiset complement (2026-04-19)

Timelog solver `test_switch_reduction_respects_vacation` initially asserted `out[p][m] == pre[p][m]` on non-vacation months and failed because the pass legitimately swapped at months 1-2; replaced with `Counter(codes across non-vac months)` comparison.

Drove rule: rearrangement-preserving operations need set/multiset-equality on the bystander complement, not element-by-element equality.

## Hook bounded reads — `pre-compact.py` (2026-04-18)

`pre-compact.py` review found `find_active_todos` used `readlines()[-200:]`, which would push the 10s hook timeout at the cliff and silently null out `active_todos` in the sidecar precisely when recovery needs it most. Commit `230feb9`.

Drove rule: `collections.deque(fh, maxlen=N)` for tail reads in transcript-parsing hooks.

## Instruction-file discipline self-incident (2026-04-23)

Condense pass on `~/.claude/CLAUDE.md` (6.3k tokens) found ~1k tokens of removable content: two bullets restating `reading-large-files.md`, two bullets carrying project-specific 2026-04-19 WP gap removal incidents that belonged here, four bullets sharing a "pre-flight sweep" shape that condensed into one reference + pointer, and one bullet citing `docs/plans/already_implemented_*` (a convention this project doesn't use). Drove the pre-commit self-review checklist in `instruction-file-discipline.md`.
