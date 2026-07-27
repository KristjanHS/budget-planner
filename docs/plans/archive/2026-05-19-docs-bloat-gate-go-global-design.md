# docs-bloat-gate goes global — grant-side design

**Date:** 2026-05-19
**Status:** Design approved; implementation pending (dotfiles session next, then grant consumer cleanup).
**Companion:** `~/projects/dotfiles/<TBD by dotfiles session>` will hold the global-side design doc.
**Builds on:** `docs/plans/archive/2026-04-26-docs-bloat-gate-redesign.md` (v2 hook design — convention defaults match its hardcoded paths).

## Why

The docs-bloat-gate hook has lived as a project-local PreToolUse hook since the v2 redesign. The slop blocklist (S2) and lexical-density check (S3) are pure prose-quality signals with no grant-specific assumptions — they belong at the user level so every project benefits. The path-gated tier-cap mechanic (S1) and the `docs/*.md` root audit-keyword block (W1.1) are project policy and stay opt-in per repo.

## Decisions

1. **Always-on S2/S3, opt-in S1.** Global hook runs S2 and S3 on every `.md` write across all projects (and on writes outside any project — `~/.claude/references/`, `~/.claude/rules/`). S1 tier caps activate only when a project surfaces gated paths.
2. **Convention defaults with JSON override.** Inside a project (`CLAUDE_PROJECT_DIR` set), the hook auto-discovers:
   - `docs/spec.md` → gated if the file exists.
   - `.claude/rules/` → gated prefix if the directory exists.
   - W1.1 audit-keyword check on `docs/*.md` root → active if `docs/` exists.
   - Override log → `docs/analysis/bloat-overrides.log` if `docs/analysis/` exists; else logging is skipped (no auto-mkdir).
   Projects that diverge drop `.claude/docs-bloat-gate.json` with any subset of: `gated_docs`, `gated_rule_prefixes`, `audit_root_check`, `log_path`, `claude_md_gated`. The grant project matches every convention default, so no config file is needed here.
3. **Memory paths exempt.** Writes matching `~/.claude/projects/*/memory/*.md` skip both S2 and S3 (and S1, which would not fire there anyway). Memory saves stay frictionless; slop accumulation in memory is accepted as a trade-off.
4. **Engine in dotfiles, no per-project copy.** The hook script and the `docs_quality.py` library both move to dotfiles and are symlinked under `~/.claude/`. No project keeps its own copy.

## Scope — grant side

Changes that happen in this repo (after the dotfiles session ships the global hook):

| File | Action |
|---|---|
| `.claude/hooks/docs-bloat-gate.py` | Delete (now lives at `~/.claude/hooks/`). |
| `.claude/lib/docs_quality.py` | Delete. No other importer (verified by grep). |
| `.claude/lib/` | Remove if empty after the lib goes. |
| `.claude/hooks/README.md` | Drop the docs-bloat-gate row from the active-hooks table; add one line pointing at `~/.claude/hooks/README.md` for the troubleshooting playbook. |
| `.claude/settings.json` | Remove the `Write\|Edit\|Bash` PreToolUse entry for docs-bloat-gate. The remaining six hooks stay project-local. |
| `CLAUDE.md` | "7 active" → "6 active"; note docs-bloat-gate is now global. |
| `.claude/rules/plan-hygiene.md` | Update the cosmetic `.claude/hooks/docs-bloat-gate.py` path reference to `~/.claude/hooks/docs-bloat-gate.py`. |
| `tests/hooks/test_docs_bloat_gate_v2.py` | Adjust five legacy tests (log-mkdir + Bash-seed + non-gated rename) and add five new tests; existing 600+ lines otherwise unchanged (fixtures use `tmp_path` with the same `docs/spec.md` and `.claude/rules/` paths the convention defaults discover). |

## Out of scope

- Re-tuning the slop wordlist or density threshold.
- Adding new signals (S4+) beyond what v2 ships.
- Migrating the `docs-only-plan-gate.py` hook (stays project-local — its `git mv` archive logic is grant-specific).
- Adding S2/S3 to Bash heredocs (content too hard to parse out of redirect/heredoc shapes; matches current v2 behavior).

## Test plan

The bulk of the 600-line suite passes through the refactor unchanged because every fixture sets `CLAUDE_PROJECT_DIR=tmp_path` and creates files at `docs/spec.md` or `.claude/rules/a.md` — the exact paths the new convention defaults auto-discover. Five legacy tests need fixture or payload adjustments:

- Three log-reading tests (`test_first_use_passes_and_logs`, `test_l2_bypass_writes_log_entry`, `test_log_is_appended_jsonl`) mkdir `docs/analysis/` upfront — the refactored hook skips logging when that dir is absent (Decision 2, no auto-mkdir).
- `TestBashGate` gains an autouse fixture seeding `docs/spec.md` + `.claude/rules/foo.md` so convention discovery activates the Bash write-pattern regex (the old hook hardcoded those paths; the global hook discovers them at runtime).
- `test_non_gated_paths_pass_through` renamed to `test_non_gated_paths_pass_s1` and its payload swapped from a slop phrase to dense filler — under always-on S2/S3, slop on non-gated `.md` correctly blocks now.

Five additions cover new behavior:

1. **`test_s2_fires_on_ungated_md`** — slop phrase on a non-gated `.md` blocks. Locks always-on S2.
2. **`test_s3_fires_on_ungated_md`** — low-density addition on a non-gated `.md` blocks. Locks always-on S3.
3. **`test_json_config_overrides_convention`** — `.claude/docs-bloat-gate.json` with a custom `gated_docs` list wins over the convention default.
4. **`test_json_config_override_disables_audit_and_redirects_log`** — `audit_root_check: false` lets a `docs/research-*.md` write pass; `log_path: "tmp.log"` redirects sentinel-bypass JSONL.
5. **`test_memory_path_bypasses_all_signals`** — slop + low-density under `**/.claude/projects/*/memory/*.md` passes.

Pyright runs clean after the lib import path resolves through the symlink.

## Done when

- The global hook fires on a known S2 phrase from outside any project (verified via the isolation-test recipe in the new `~/.claude/hooks/README.md`).
- The global hook fires on an S1 tier-cap overrun against `docs/spec.md` from this project, without any project-side hook installed.
- A slop-laden memory-file write passes silently.
- `pytest tests/hooks/test_docs_bloat_gate_v2.py -v` is green, including the five new tests (`test_s2_fires_on_ungated_md`, `test_s3_fires_on_ungated_md`, `test_json_config_overrides_convention`, `test_json_config_override_disables_audit_and_redirects_log`, `test_memory_path_bypasses_all_signals`).
- `python scripts/recalc.py output/budget_plan.xlsx && pytest tests/ -v` is green (no regression in the broader suite from settings.json changes).
- The Active gates list in `CLAUDE.md` Critical Rule 7 reflects six project gates plus the global one.

## Handoff

Next session runs from `~/projects/dotfiles/` cwd. That session:
1. Writes the dotfiles-side companion design doc.
2. Moves `docs-bloat-gate.py` and `docs_quality.py` into the dotfiles tree (with `git mv` where the source lives in grant — actually a delete-and-add since the dotfiles repo is separate; history stays in grant's git log).
3. Refactors the hook per Decisions 1–3.
4. Creates `~/.claude/lib/` plus per-file symlinks for the hook and lib.
5. Edits `~/.claude/settings.json` (via the dotfiles path) to register the global PreToolUse entry.
6. Ships `~/.claude/hooks/README.md` (moved-and-reframed from the grant project copy).
7. Commits in dotfiles.

After dotfiles ships, return to grant cwd and execute the consumer-side delete/edit list above.
