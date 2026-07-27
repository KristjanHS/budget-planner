WORKBOOK := output/budget_plan.xlsx

.PHONY: help check test recalc lint fmt

help:
	@echo "make check    - verification gate: recalc the workbook, then run the full pytest suite"
	@echo "make test     - pytest only (skips recalc; data_only reads can return None)"
	@echo "make recalc   - recalculate $(WORKBOOK) via LibreOffice"
	@echo "make lint     - pyright + ruff check + ruff format --check over the whole project"
	@echo "make fmt      - ruff format (rewrites files in place)"
	@echo "make syncwt|commit|commitwt - shared worktree targets (~/.config/make/worktree.mk)"

# The gate. CLAUDE.md Critical Rule 2: skipping recalc makes data_only=True reads
# return None, so tests pass silently on broken output. `syncwt` auto-detects this
# target as its verification gate.
check: recalc test

recalc:
	@if [ -f "$(WORKBOOK)" ]; then \
		python scripts/recalc.py "$(WORKBOOK)"; \
	else \
		echo "note: $(WORKBOOK) absent — nothing to recalc (generate it first)"; \
	fi

test:
	@pytest tests/ -q

# ruff is a host tool here (deliberately not in uv.lock), so it runs bare — the same
# binary the PostToolUse ruff hook falls back to, which keeps the two verdicts identical.
lint:
	@uv run pyright
	@ruff check --no-cache .
	@ruff format --check --no-cache .

# `ruff format` is NOT a one-pass fixed point on this repo: one pass over the builder
# modules leaves 7 of them still needing a reformat (the `style_cell(...).value = (...)`
# long-assignment shape). Reproduced on ruff 0.12.9 and 0.16.0 alike — so fmt runs to
# convergence and then asserts it, rather than trusting a single pass.
#
# --no-cache everywhere because a write pass records its own output as formatted: a
# cached `--check` right after a cached `ruff format` reported all 85 files clean while
# a cold check found the 7. The cache is not the bug, but it hides it, and it saves
# 0.02s on a repo this size.
fmt:
	@ruff format --no-cache . >/dev/null
	@ruff format --no-cache .
	@ruff format --check --no-cache .

# Shared worktree/commit targets (syncwt, commit, commitwt). At the bottom so the
# include can't steal the default goal (first target read wins → keep `help`).
-include $(HOME)/.config/make/worktree.mk
