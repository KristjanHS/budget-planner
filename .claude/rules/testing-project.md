---
paths:
  - "tests/**"
  - "**/test_*.py"
---
# Project Testing Rules

Tests must not hardcode config.py seed values (WP names, person names, FTE amounts, start/end months). Derive expected values from config imports or from the generated workbook.
When a refactor touches both the reader and a builder, the test must exercise the full round-trip: init → user-edit via openpyxl → `rebuild_workbook` → reopen and assert. A reader-only "does it parse?" assertion is a false green — the builder may still clobber the preserved value.
Docs-only stages don't warrant the /impag "full-suite before final commit" run. If `git diff --stat` for the stage's last task is all `.md`/`.txt` plus at most trivial docstring/comment/path-string edits in `.py`, skip the full suite — the PostToolUse pyright hook validates Python edits and markdown can't regress the suite. Keep the full run when any substantive Python changed earlier in the stage.
When tightening reader validation (strict enum / blank-cell / range enforcement), grep Config-cell-writing test fixtures for the columns being validated. Fixtures that populate by column index (e.g. `ws.cell(r, 10, ...)` or per-cell marker-seeding loops) bypass the builder's data validation, so values the old reader tolerated can surface as ValueError cross-test failures unrelated to the validation target. Grep target: `ws\.cell\(.*,\s*COL_<name>|ws\.cell\(.*,\s*<col_idx>` and enum-seeding loops in `tests/`.
Refactor-twin: when deleting a seed value, `grep -rn '\bOLD\b' tests/` first — stale string refs in docstrings can pass on coincidence.
Test cell writes: use `ws.cell(row=r,column=c,value=X)` or `_set()` helper — direct `.value=X` trips pyright `Cell|MergedCell`.
Named-range fixture writes: target the `DefinedName` cell (grep `workbook.py`) — sibling override strips fake it (date fallback masks).
