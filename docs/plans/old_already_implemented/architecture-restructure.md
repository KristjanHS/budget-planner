# Architecture Restructure Plan

**Status:** APPROVED v3
**Date:** 2026-03-31

## Design Decisions

1. **Block Registry** in workbook.py for future cost category expansion
2. **Read-back support** — preserve user data (actuals + manual planned edits) during regeneration
3. **No premature abstraction** — no subpackages, no generic iteration, no layout contract module
4. **Flat package** — every module is top-level in budget_planner/
5. **3 test files** — grouped by fixture dependency, not by spreadsheet area

---

## Proposed Structure

```
budget-planner/
  pyproject.toml
  budget_planner/
    __init__.py
    __main__.py              # CLI: python -m budget_planner [--update path]
    config.py                # MOVED from ./config.py
    styles.py                # EXTRACTED — fonts, fills, borders, style_cell()
    plan.py                  # Plan sheet: globals + config table + both FTE blocks
    summary.py               # Summary sheet builder
    meta.py                  # _META sheet builder
    reader.py                # discover_layout() + read_workbook()
    workbook.py              # Orchestrator (create + update modes, block registry)
  tests/
    conftest.py              # Fixtures: fresh wb, wb-with-actuals, recalculated wb
    test_workbook.py         # All 13 structural checks (shares wb fixture)
    test_roundtrip.py        # Reader + update preserves actuals
    test_recalc.py           # LibreOffice recalc (auto-skip if unavailable)
  scripts/
    office/                  # UNCHANGED
    recalc.py                # STAYS HERE (imports from office/)
  output/                    # RENAMED from xls/
  .claude/
    rules/
      dev-workflow.md        # Adapted from another-cli
  docs/plans/
  CLAUDE.md, SPEC.md, CONVENTIONS.md, etc.
```

**Deleted after migration:**
- `./config.py` → `budget_planner/config.py`
- `./build_budget.py` → `budget_planner/plan.py` + `summary.py` + `meta.py` + `workbook.py`
- `./verify.py` → `tests/`
- `./xls/` → `output/`

---

## Module Design

### config.py

Moved from root. Unchanged content. Single source of truth for project data.
`build_budget.py`'s duplicate `EXAMPLE_WPS` and constants are deleted.

### styles.py

Extracted from `build_budget.py` lines 31-58:
- Font constants: `BLUE_INPUT`, `BLACK_FORMULA`, `GREEN_XREF`, `HEADER_FONT`, `TITLE_FONT`, `SECTION_FONT`
- Fill constants: `GREY_FILL`, `LIGHT_BLUE_FILL`, `LIGHT_GREEN_FILL`, `SECTION_FILL`, etc.
- `THIN_BORDER`, `CENTER`, `WRAP`
- `style_cell(ws, row, col, ...)` helper
- `col_letter(c)` wrapper

~50 lines, no logic.

### plan.py

Builds the entire Plan sheet. Contains:

1. `build_plan_sheet(ws, preserved_data={})` — top-level function called by orchestrator
   - Writes global parameters (rows 1-2, named ranges) — ~10 lines inline
   - Writes config table — ~30 lines inline
   - Calls `_build_month_block()` twice (planned + actual)
   - Returns layout dict with all row numbers

2. `_build_month_block(ws, start_row, title, is_actuals, preserved_data={})` — the existing
   `build_month_block` from `build_budget.py:275-392`, with one addition: if `preserved_data`
   has values for a WP, use them instead of even-distribution defaults (planned) or zeros (actual).

The layout dict returned:
```python
{
    "config_first_wp_row": 6,
    "config_last_wp_row": 13,
    "config_total_row": 14,
    "planned_header_row": 16,
    "planned_first_wp_row": 19,
    "planned_last_wp_row": 26,
    "planned_total_row": 27,
    "actual_header_row": 29,
    "actual_first_wp_row": 32,
    "actual_last_wp_row": 39,
    "actual_total_row": 40,
}
```

Plain keys, no namespacing — there are only two blocks today.

### summary.py

```python
def build_summary_sheet(wb, layout):
```
Takes workbook and layout dict. Builds Summary sheet. Knows about planned and actual
explicitly — no generic iteration. Current `build_budget.py:414-558` extracted as-is.

### meta.py

```python
def build_meta_sheet(wb, layout):
```
Takes workbook and layout dict. Writes _META documentation. The big format string
from `build_budget.py:73-205` extracted as-is. Uses layout dict values for row references.

### reader.py

Two functions:

```python
def discover_layout(ws) -> dict:
    """Scan existing Plan worksheet, reconstruct layout dict.
    
    Heuristics (same as old verify.py):
    - Scan column A rows 1-20 for "WP ID" → config_header_row
    - Scan for "PLANNED" / "ACTUAL" → block headers
    - Count WP rows between headers
    """

def read_workbook(filepath: str) -> dict:
    """Read existing workbook, extract user-entered cell values.
    
    Returns:
        {
            "planned": {
                0: [0.2, 0.2, None, ...],   # 60 values per WP index
                1: [0.44, 0.44, ...],
            },
            "actual": {
                0: [0.15, 0.2, 0.0, ...],
                1: [0.5, 0.0, ...],
            },
        }
    
    None = inactive/empty cell. Values preserved exactly as entered.
    """
```

### workbook.py

Orchestrator + block registry + CLI wiring.

```python
from budget_planner.plan import build_plan_sheet
from budget_planner.summary import build_summary_sheet
from budget_planner.meta import build_meta_sheet
from budget_planner.reader import read_workbook

def generate_workbook(output_path: str = "output/budget_plan.xlsx",
                      existing_path: str | None = None) -> str:
    """Generate workbook. If existing_path given, preserve user data."""
    preserved = read_workbook(existing_path) if existing_path else {}

    wb = Workbook()
    plan = wb.active
    plan.title = "_META"  # will be reordered

    plan_ws = wb.create_sheet("Plan")
    layout = build_plan_sheet(plan_ws, preserved_data=preserved)
    build_summary_sheet(wb, layout)
    build_meta_sheet(wb, layout)
    wb.save(output_path)
    return output_path
```

### __main__.py

```python
import argparse
from budget_planner.workbook import generate_workbook

parser = argparse.ArgumentParser(description="Generate programme budget workbook")
parser.add_argument("--update", metavar="PATH", help="Existing workbook to preserve data from")
parser.add_argument("-o", "--output", default="output/budget_plan.xlsx")
args = parser.parse_args()
generate_workbook(output_path=args.output, existing_path=args.update)
```

---

## Round-Trip Workflow

### Fresh generation
```bash
python -m budget_planner
python scripts/recalc.py output/budget_plan.xlsx
pytest tests/ -v
```

### Update (preserves actuals)
```bash
python -m budget_planner --update output/budget_plan.xlsx
python scripts/recalc.py output/budget_plan.xlsx
pytest tests/ -v
```

### What gets preserved vs regenerated

| Element | Preserved | Regenerated |
|---|---|---|
| Planned FTE monthly values (blue cells) | Yes | Only for new WPs |
| Actual FTE monthly values (blue cells) | Yes | Only for new WPs |
| Formulas (black cells) | No | Always |
| Styles, formatting, borders | No | Always |
| Config table | No | From config.py |
| Named ranges, _META, Summary | No | Always |

---

## Pytest Migration

### conftest.py

```python
import json, subprocess, pytest, openpyxl
from budget_planner.workbook import generate_workbook
from budget_planner.reader import discover_layout

@pytest.fixture(scope="session")
def fresh_wb_path(tmp_path_factory):
    out = tmp_path_factory.mktemp("wb") / "test_budget.xlsx"
    generate_workbook(str(out))
    return str(out)

@pytest.fixture(scope="session")
def recalc_result(fresh_wb_path):
    result = subprocess.run(
        ["python3", "scripts/recalc.py", fresh_wb_path],
        capture_output=True, text=True
    )
    if result.returncode != 0:
        pytest.skip("LibreOffice not available")
    return json.loads(result.stdout)

@pytest.fixture(scope="session")
def wb(fresh_wb_path, recalc_result):
    return openpyxl.load_workbook(fresh_wb_path, data_only=True)

@pytest.fixture
def wb_with_actuals(tmp_path):
    path = tmp_path / "with_actuals.xlsx"
    generate_workbook(str(path))
    wb = openpyxl.load_workbook(str(path))
    plan = wb["Plan"]
    layout = discover_layout(plan)
    row = layout["actual_first_wp_row"]
    plan.cell(row, 3, 0.15)   # WP1, M1
    plan.cell(row, 4, 0.20)   # WP1, M2
    wb.save(str(path))
    return str(path)
```

### test_workbook.py — All 13 structural checks

Maps directly from verify.py:

```python
# Checks 2-3: Sheet existence + global params
def test_sheets_exist(wb): ...
def test_project_start_date(wb): ...
def test_fte_rate(wb): ...

# Check 4: Named ranges
def test_named_ranges(wb): ...

# Checks 5-6: Config table (parametrized per WP)
def test_config_header_exists(wb): ...
@pytest.mark.parametrize("wp_idx", range(8))
def test_wp_config_valid(wb, wp_idx): ...

# Checks 7-8: Block headers
def test_planned_block_exists(wb): ...
def test_actual_block_exists(wb): ...

# Checks 9-10: FTE validation (parametrized per WP)
@pytest.mark.parametrize("wp_idx", range(8))
def test_planned_fte_variance(wb, wp_idx): ...
@pytest.mark.parametrize("wp_idx", range(8))
def test_actual_fte_zeros(wb, wp_idx): ...

# Check 11: Inactive months (parametrized per WP)
@pytest.mark.parametrize("wp_idx", range(8))
def test_inactive_months_empty(wb, wp_idx): ...

# Check 12: Summary
def test_summary_wp_count(wb): ...

# Check 13: _META
def test_meta_has_content(wb): ...
def test_meta_has_modification_guide(wb): ...
```

### test_roundtrip.py — Reader + update

```python
def test_discover_layout_finds_all_blocks(fresh_wb_path): ...
def test_read_workbook_extracts_planned_values(fresh_wb_path): ...
def test_read_workbook_extracts_actual_values(wb_with_actuals): ...
def test_update_preserves_actuals(wb_with_actuals): ...
def test_update_preserves_manual_planned_edits(wb_with_actuals): ...
def test_new_wp_gets_defaults_not_preserved(wb_with_actuals): ...
```

### test_recalc.py — LibreOffice

```python
def test_recalculation_zero_errors(recalc_result): ...
def test_formula_count_positive(recalc_result): ...
```

### pyproject.toml

```toml
[project]
name = "budget-planner"
version = "0.1.0"
requires-python = ">=3.10"
dependencies = [
    "openpyxl>=3.1.0",
    "python-dateutil",
]

[project.optional-dependencies]
dev = ["pytest>=7.0"]

[tool.pytest.ini_options]
testpaths = ["tests"]
```

---

## Dependency Graph

```
Step 1:  pyproject.toml + budget_planner/__init__.py + __main__.py   (standalone)
Step 2:  budget_planner/styles.py                                    (standalone)
Step 3:  budget_planner/config.py                                    (move, standalone)
Step 4:  budget_planner/plan.py                                      (depends: 2, 3)
Step 5:  budget_planner/summary.py                                   (depends: 2, 3)
Step 6:  budget_planner/meta.py                                      (depends: 2)
Step 7:  budget_planner/reader.py                                    (standalone)
Step 8:  budget_planner/workbook.py                                  (depends: 4, 5, 6, 7)
Step 9:  tests/conftest.py + test_workbook.py + test_roundtrip.py + test_recalc.py  (depends: 8)
Step 10: .claude/rules/dev-workflow.md                            (standalone)
Step 11: Rename xls/ → output/, delete old files, update CLAUDE.md (depends: all)
```

**Parallel groups:**
- A: Steps 1, 2, 3, 7, 10 (all standalone)
- B: Steps 4, 5, 6 in parallel (different files, shared deps from A)
- C: Step 8 (after B)
- D: Step 9 (after C)
- E: Step 11 (after all)

**File overlap check:** No parallel steps write the same file.

---

## Dev Workflow Rule

`.claude/rules/dev-workflow.md`:

```markdown
---
paths: "**/*.py"
---
# Dev Workflow

## Build + Test
After code changes: `pytest tests/ -v`
Recalc tests need LibreOffice — auto-skip if unavailable.
Run pytest once per step. Never re-run to "confirm".

## Update mode
When testing workbook updates, always inject realistic actuals
into the test fixture — don't just test fresh generation.

## Linting
Run pyright before committing. Fix all errors.
Subagents create dead code — verify after each task.

## Code Review
Suggest review before committing.
Move completed plans to docs/plans/old_already_implemented/.

## Invariants
- Excel formulas only — never hardcode calculated values
- Font color: Blue=input, Black=formula, Green=cross-sheet
- _META sheet must reflect current workbook structure
- Block builders must accept preserved_data for round-trip safety
```
