# Python 3.14 Upgrade Plan

> **For Claude:** REQUIRED SUB-SKILL: Use superpowers:executing-plans to implement this plan task-by-task.

**Goal:** Upgrade the project from Python 3.12/3.13 to Python 3.14 (latest stable: 3.14.4) as the minimum supported version, and adopt relevant 3.14 features.

**Architecture:** The upgrade is low-risk — all 259 tests already pass on Python 3.14.0. The main work is updating version pins across config files, upgrading the dev environment, refreshing the lockfile, and selectively adopting Python 3.14 features (deferred annotations, template strings where appropriate).

**Tech Stack:** Python 3.14, uv, openpyxl, pytest, devcontainers

---

## Pre-flight Summary

| Item | Current | Target |
|------|---------|--------|
| Runtime | Python 3.12.12 | Python 3.14.4 |
| `requires-python` | `>=3.10` | `>=3.14` |
| Devcontainer | `uv python install 3.13` | `uv python install 3.14` |
| `uv.lock` | `requires-python = ">=3.10"` | Regenerated for `>=3.14` |
| Test suite | 259 passed on 3.14.0 | 259 passed on 3.14.4 |

---

### Task 1: Update `pyproject.toml` minimum Python version

**Files:**
- Modify: `pyproject.toml:5` (`requires-python` line)

**Step 1: Edit `requires-python`**

In `pyproject.toml`, change:

```python
requires-python = ">=3.10"
```

to:

```python
requires-python = ">=3.14"
```

**Step 2: Commit**

```bash
git add pyproject.toml
git commit -m "chore: bump requires-python to >=3.14"
```

---

### Task 2: Upgrade devcontainer to Python 3.14

**Files:**
- Modify: `.devcontainer/Dockerfile:82-83`

**Step 1: Update Python version in Dockerfile**

Change lines 82-83 in `.devcontainer/Dockerfile`:

```dockerfile
# Install Python 3.13 via uv (fast binary download, not source compilation)
RUN uv python install 3.13 --default
```

to:

```dockerfile
# Install Python 3.14 via uv (fast binary download, not source compilation)
RUN uv python install 3.14 --default
```

**Step 2: Commit**

```bash
git add .devcontainer/Dockerfile
git commit -m "chore: upgrade devcontainer Python to 3.14"
```

---

### Task 3: Install Python 3.14.4 and recreate project venv

**Step 1: Install latest 3.14 patch**

```bash
uv python install 3.14
```

**Step 2: Recreate the project virtual environment**

```bash
rm -rf .venv
uv venv --python 3.14
uv pip install -e ".[dev]"
```

**Step 3: Verify Python version**

```bash
.venv/bin/python --version
```

Expected: `Python 3.14.4` (or latest 3.14.x)

**Step 4: Regenerate lockfile**

```bash
uv lock
```

**Step 5: Run full test suite**

```bash
pytest tests/ -q
```

Expected: `259 passed`

**Step 6: Commit**

```bash
git add uv.lock
git commit -m "chore: regenerate uv.lock for Python 3.14"
```

---

### Task 4: Remove `from __future__ import annotations`

Python 3.14 has native deferred annotation evaluation (PEP 649/749), making `from __future__ import annotations` unnecessary.

**Files:**
- Modify: `tests/test_trends.py:2` (remove `from __future__ import annotations`)

**Step 1: Write a test to verify annotations still work**

Run the existing test suite — if `test_trends.py` tests pass without the import, we're good.

**Step 2: Remove the import**

In `tests/test_trends.py`, remove line 2:

```python
from __future__ import annotations
```

**Step 3: Run the affected tests**

```bash
pytest tests/test_trends.py -v
```

Expected: All tests pass (deferred evaluation is now built-in).

**Step 4: Commit**

```bash
git add tests/test_trends.py
git commit -m "chore: remove __future__ annotations (native in 3.14)"
```

---

### Task 5: Scan for any remaining old-style typing patterns

**Files:**
- Potentially modify: any `.py` files still using `from typing import` for types available as builtins

**Step 1: Search for modernizable typing imports**

```bash
grep -rn "from typing import" budget_planner/ tests/
```

Review results. In Python 3.14:
- `typing.Literal` and `typing.overload` are still needed (no builtin equivalent)
- `typing.TYPE_CHECKING` is still needed
- Any `Optional`, `Union`, `Dict`, `List`, `Tuple` → replace with `X | None`, `X | Y`, `dict`, `list`, `tuple`

**Step 2: Apply any needed modernizations**

Based on the codebase scan, only `Literal`, `overload`, and `TYPE_CHECKING` are imported from `typing` — these have no builtin replacement. **No changes expected.**

**Step 3: Run tests**

```bash
pytest tests/ -q
```

Expected: 259 passed

**Step 4: Commit (only if changes were made)**

```bash
git commit -m "chore: modernize typing imports for Python 3.14"
```

---

### Task 6: Update project documentation

**Files:**
- Modify: `CLAUDE.md` (if Python version is mentioned)
- Modify: `README.md` (if Python version is mentioned)

**Step 1: Search for version references**

```bash
grep -rn "3\.1[0-3]\|Python 3\." CLAUDE.md README.md docs/
```

**Step 2: Update any stale version references**

Update any mentions of Python 3.10/3.12/3.13 to reflect the new 3.14 minimum.

**Step 3: Commit**

```bash
git add CLAUDE.md README.md docs/
git commit -m "docs: update Python version references to 3.14"
```

---

### Task 7: Final verification

**Step 1: Clean test run**

```bash
pytest tests/ -v
```

Expected: 259 passed, 0 warnings

**Step 2: Verify workbook generation**

```bash
python -m budget_planner init --demo
python scripts/recalc.py output/budget_plan.xlsx
```

Expected: Workbook generates and recalculates without errors.

**Step 3: Verify Python version in venv**

```bash
python --version
```

Expected: `Python 3.14.x`

---

## Out of Scope (Future Consideration)

These Python 3.14 features are noted but **not included** in this upgrade plan:

| Feature | Relevance | Why deferred |
|---------|-----------|-------------|
| **Template strings (PEP 750)** | Could use for Excel formula generation | Low priority; current f-strings work fine |
| **Free-threaded Python** | Parallel sheet building | Experimental; no perf bottleneck exists |
| **`concurrent.interpreters`** | Multi-interpreter concurrency | Overkill for this project's scope |
| **Tail-call interpreter (3-5% speedup)** | Automatic; no code changes needed | Already active by default |
