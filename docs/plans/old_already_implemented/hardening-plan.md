# Hardening Plan — Brittleness Analysis & Remediation

**Date:** 2026-04-03
**Status:** Draft — awaiting review

---

## Brittleness Analysis

### CRITICAL — Silent data corruption risk

#### C1. `month_col_start = 3` duplicated in 6 modules

**Files:** `plan.py:99`, `capacity.py:69`, `actuals.py:45`, `summary.py:40`, `trends.py:52`, `meta.py:50`

Every builder independently sets `month_col_start = 3`. Cross-sheet formulas assume this is consistent — if one module changes and another doesn't, formulas reference wrong columns.

The downstream derived values (`val_col_sum`, `val_col_budget`, `val_col_diff`, `val_col_cum`, `val_col_pct_ceil`, `val_col_breach`) are also re-derived independently in each module.

**Impact:** Adding a column before month data (e.g., "Category") requires synchronized changes to 6 modules + reader.py, with no compile-time or runtime check that they agree.

#### C2. `_wp_config_row()` default `config_first_wp_row=8`

**File:** `capacity.py:15` — `def _wp_config_row(wp_id, work_packages, config_first_wp_row=8)`
**Called at:** `capacity.py:188` — `config_row = _wp_config_row(wp_id, work_packages)` (no layout param!)

The caller doesn't pass the actual `config_first_wp_row` from the plan layout. It relies on the hardcoded default of 8. Today the Plan builder produces `config_first_wp_row=8`, but this is an emergent property of the header layout (4 global rows + 1 gap + 1 section header + 1 column header = row 8). If Plan ever adds/removes a header row, this default silently generates wrong `=Plan!A{n}` formulas.

**Impact:** Capacity sheet formulas reference wrong Plan config rows. Budget and variance formulas break silently.

#### C3. Config sheet reader assumes fixed row positions

**File:** `reader.py:14-23`
```python
project_params = {
    "title": ws.cell(1, 2).value,     # Row 1
    "start": ws.cell(2, 2).value,     # Row 2
    "months": int(ws.cell(3, 2).value or 60),  # Row 3
    "rate": float(ws.cell(4, 2).value or 7500), # Row 4
}
r = 8  # WP data starts here
```

If a user inserts a row in the Config sheet (e.g., adds a "Project ID" field), the reader silently reads wrong cells. No validation that row 1 actually contains "Project Title".

**Impact:** Rebuild reads garbage data and generates a corrupted workbook.

#### C4. Actuals reader hardcodes `first_wp_row = 4`

**File:** `reader.py:426`
```python
first_wp_row = 4
```

If the Actuals header structure changes (e.g., adding a sub-header row), the reader reads from the wrong rows.

---

### HIGH — Cross-sheet formula breakage

#### H1. Capacity → Plan formula coupling via `capacity_layout["person_groups"]`

**File:** `plan.py:50-52`
```python
for person_name, group in capacity_layout["person_groups"].items():
    if wp_id in group["wp_rows"]:
        cap_rows.append(group["wp_rows"][wp_id])
```

`_apply_planned_formulas` trusts that `capacity_layout` has the exact dict structure produced by `build_capacity_sheet`. There's no schema validation — a renamed key or structural change causes a silent KeyError or missing data.

#### H2. Cross-sheet formulas hardcode sheet names as strings

Throughout all builders:
- `f"=Plan!A{config_row}"` (capacity.py:192, actuals.py:132, summary.py:107, trends.py:101)
- `f"=Team!A{roster_row}"` (capacity.py:169, trends.py:275)
- `f"=Actuals!{cl}{act_total}"` (trends.py:373)
- `f"=Config!B1"` (plan.py:110)

Sheet name strings are scattered across 7 modules. Renaming a sheet requires a project-wide search-and-replace.

#### H3. Trends section rows hardcoded

**File:** `trends.py`
- `section1_row = 3` (line 65)
- `header_row = 4` (line 74)
- `first_wp_data_row = 5` (line 91)
- `person_section_row = 32` (line 245)
- `person_header_row = 33` (line 254)
- `person_first_row = 34` (line 263)
- `helper_row_planned = 40` (line 342)

The gap between section 1 (ends at row ~14) and section 2 (starts at row 32) accommodates the chart. If the number of WPs exceeds ~7, WP rows collide with the chart anchor at row 15. The person load section at row 32 has no dynamic spacing — more WPs push the total row into chart territory.

#### H4. Reader string-matching fragility

**File:** `reader.py`
- `val == "WP ID"` (line 364) — exact match, case-sensitive
- `ws.cell(r, 1).value == "Person"` (line 78) — exact match
- `"WP SHARE SEEDS" in str(val).upper()` (line 103) — case-insensitive substring
- `"EXTERNAL COMMITMENTS" in str(val).upper()` (line 179) — limited to first 20 rows
- `str(col_b).upper() in ("TOTAL", "REMAINING")` (line 305) — case-insensitive

Inconsistent matching strategies. A user editing "WP ID" to "WP Id" or "WP #" breaks discover_layout. The 20-row search limit for external commitments means adding config rows could push the section beyond detection range.

---

### HIGH — UX bugs

#### H5. `_changed` test output skips recalc (no auto-fit)

`recalc.py`'s LibreOffice macro auto-fits all columns except `_META` (lines 28-46). The main `budget_plan.xlsx` gets recalc'd, but `test_rebuild_e2e.py:_save_change()` copies the rebuilt workbook to `output/budget_plan_change.xlsx` **before** recalc runs. Result: the `_changed` file has default 8.43-char column widths.

`summary.py:248` also has a dead comment `# Column widths` with no code after it.

**Impact:** The `_changed` file meant for manual inspection looks broken on open.

#### H6. Sheet protection blocks manual column/row resizing

**File:** `workbook.py:197` — `ws.protection.set_password("grant")`

openpyxl's `SheetProtection.set_password()` defaults to `formatColumns=True` and `formatRows=True`. In the OOXML spec, `True` means the action is **restricted**. So users cannot resize columns or rows in Excel, even to auto-fit. This makes the missing-width issue (H5) worse — users can't even fix it themselves.

**Impact:** Workbook appears broken (truncated columns) and user has no workaround short of unprotecting the sheet.

---

### MEDIUM — Maintenance burden / poor error locality

#### M1. Magic threshold numbers in formula strings

| Location | Threshold | Meaning |
|----------|-----------|---------|
| `trends.py:131` | `0.05` | OK band |
| `trends.py:132` | `0.15` | OVER threshold |
| `trends.py:134` | `-0.10` | SLOW threshold |
| `capacity.py:351` | `0.8` | Amber load |
| `actuals.py:254` | `0.85` | Amber % ceiling |
| `actuals.py:265` | `12` | Amber breach months |
| `actuals.py:268` | `6` | Red breach months |

Buried inside f-string formulas, not configurable without editing Python code.

#### M2. Summary COLS_PER_YEAR arithmetic

**File:** `summary.py:62-66`
```python
COLS_PER_YEAR = 4
TOTAL_GROUP_START = YEAR_GROUP_START + num_years * COLS_PER_YEAR
```

Adding a metric (e.g., "Forecast EUR") means changing COLS_PER_YEAR AND updating all `sc + 1`, `sc + 2`, `sc + 3` offsets AND updating the sub-header list. No named mapping between column purpose and offset.

#### M3. Year-header/month-header code duplicated 4 times

The year-header + month-header block is copy-pasted across `plan.py:256-300`, `capacity.py:112-153`, `actuals.py:72-123`, and partially in `summary.py`. Each has slight variations (different validation columns, different formats). A change to the header pattern must be applied 4 times.

#### M4. `discover_layout` assumes 3-row gap between header and WP data

**File:** `reader.py:392`
```python
planned_first_wp_row = planned_header_row + 3
```

Assumes: section header → year row → month row → first WP row. If the Plan builder adds a sub-header, this offset breaks.

---

## Hardening Plan

### Step 1: Centralize layout constants (CRITICAL — fixes C1, C2)

**Create `budget_planner/layout.py`** — single source of truth for column structure.

```python
# Column structure
MONTH_COL_START = 3  # Columns A=ID, B=Name, C+=months

# Validation column offsets (relative to month_col_end)
VAL_OFFSET_SUM = 1
VAL_OFFSET_BUDGET = 2
VAL_OFFSET_DIFF = 3
# Actuals-only additional columns
VAL_OFFSET_CUM = 4
VAL_OFFSET_PCT_CEIL = 5
VAL_OFFSET_BREACH = 6

def month_col_end(project_months):
    return MONTH_COL_START + project_months - 1

def val_columns(project_months):
    """Return dict of validation column numbers."""
    end = month_col_end(project_months)
    return {
        "sum": end + VAL_OFFSET_SUM,
        "budget": end + VAL_OFFSET_BUDGET,
        "diff": end + VAL_OFFSET_DIFF,
        "cum": end + VAL_OFFSET_CUM,
        "pct_ceil": end + VAL_OFFSET_PCT_CEIL,
        "breach": end + VAL_OFFSET_BREACH,
    }
```

**Changes to each builder:** Replace local `month_col_start = 3` and derived vars with imports from `layout.py`.

**Fix C2:** Pass `config_first_wp_row` explicitly at the call site in `capacity.py:188`. Remove the default parameter.

**Files touched:** `layout.py` (new), `plan.py`, `capacity.py`, `actuals.py`, `summary.py`, `trends.py`, `meta.py`, `config.py` (move constants)
**Tests:** Existing tests pass (no behavior change). Add a test that imports from all builders and asserts they use the same `MONTH_COL_START`.
**Risk:** Low — pure refactor, values unchanged.

### Step 2: Centralize sheet name constants (HIGH — fixes H2)

**Add to `layout.py`:**
```python
SHEET_META = "_META"
SHEET_CONFIG = "Config"
SHEET_PLAN = "Plan"
SHEET_TEAM = "Team"
SHEET_CAPACITY = "Capacity"
SHEET_ACTUALS = "Actuals"
SHEET_SUMMARY = "Summary"
SHEET_TRENDS = "Trends"
```

Replace all `"Plan"`, `"Config"`, etc. string literals across builders, readers, and workbook.py with these constants. Also replace sheet name strings inside formula f-strings (e.g., `f"=Plan!A{row}"` → `f"={SHEET_PLAN}!A{row}"`).

**Files touched:** All builder modules, `reader.py`, `workbook.py`
**Tests:** Existing tests pass. Add test that all sheet names in generated workbook match constants.
**Risk:** Low — rename-safe refactor.

### Step 3: Add reader validation assertions (CRITICAL — fixes C3, C4)

Add structural assertions to readers so they fail fast with clear messages instead of silently reading wrong data.

**Config reader (`reader.py:5-37`):**
```python
# Validate expected labels in column A
label_1 = ws.cell(1, 1).value
if label_1 is None or "title" not in str(label_1).lower():
    raise ValueError(
        f"Config sheet row 1 col A expected 'Project Title:', got {label_1!r}. "
        "Was the Config sheet manually edited?"
    )
# Similar for rows 2-4
```

**Actuals reader (`reader.py:408-452`):**
Replace `first_wp_row = 4` with layout discovery: scan for "WP ID" header row, set `first_wp_row = header_row + 1`.

**discover_layout (`reader.py:354-405`):**
Replace `planned_first_wp_row = planned_header_row + 3` with a scan for the first row after the header that has a WP ID (formula or literal) in column A.

**Files touched:** `reader.py`
**Tests:** Add tests for malformed workbooks (missing headers, shifted rows) that verify ValueError with helpful message.
**Risk:** Low — adds validation, no behavior change for valid workbooks.

### Step 4: Make Trends row positions dynamic (HIGH — fixes H3)

Currently Trends hardcodes section start rows (3, 32, 40). Make them computed:

```python
section1_row = 3
# ... WP rows computed from num_wps ...
total_row = first_wp_data_row + num_wps
chart_anchor_row = total_row + 2

person_section_row = chart_anchor_row + 16  # chart height ~15 rows
person_header_row = person_section_row + 1
person_first_row = person_header_row + 1

helper_start_row = person_first_row + len(team) + 2
helper_row_planned = helper_start_row
helper_row_actual = helper_start_row + 1
helper_row_ceiling = helper_start_row + 2
helper_row_labels = helper_start_row + 3
```

**Files touched:** `trends.py`
**Tests:** Existing Trends tests should still pass (they use layout dict, not hardcoded rows). Add parametric test with 3 WPs and 15 WPs to verify no row collisions.
**Risk:** Medium — chart anchor changes could shift visuals. Verify with LibreOffice.

### Step 5: Harden reader string matching (HIGH — fixes H4)

Standardize all section header detection to:
1. Case-insensitive matching
2. Substring matching (not exact)
3. Configurable search range (not hardcoded to 20 rows)
4. Clear error messages on failure

Create a helper:
```python
def _find_header_row(ws, pattern, col=1, max_row=None, required=True):
    """Find row where cell matches pattern (case-insensitive substring)."""
    limit = max_row or ws.max_row
    for r in range(1, limit + 1):
        val = ws.cell(r, col).value
        if val and pattern.upper() in str(val).upper():
            return r
    if required:
        raise ValueError(f"Cannot find '{pattern}' in column {col}, rows 1-{limit}")
    return None
```

Replace all ad-hoc header scans with calls to this helper.

**Files touched:** `reader.py`
**Tests:** Add tests with varied header casing/whitespace.
**Risk:** Low — more permissive matching won't break valid workbooks.

### Step 6: Fix sheet protection to allow column/row resizing (HIGH — fixes H6)

**File:** `workbook.py:191-197`

After `set_password("grant")`, explicitly allow formatting:

```python
def _apply_protection(wb):
    for ws in wb.worksheets:
        ws.protection.set_password("grant")
        ws.protection.formatColumns = False
        ws.protection.formatRows = False
```

This keeps cell content protection (blue cells editable, formula cells locked) while letting users resize columns and rows freely.

**Files touched:** `workbook.py`
**Tests:** Add test asserting `formatColumns is False` and `formatRows is False` on all sheets. Update existing protection tests if they assert on default values.
**Risk:** Low — strictly less restrictive, no behavior change for cell editing.

### Step 7: Run recalc on `_changed` test output + cleanup dead code (HIGH — fixes H5)

The `_save_change()` helper in `test_rebuild_e2e.py` copies the rebuilt workbook before recalc runs. The LibreOffice macro in `recalc.py` already handles auto-fit on all sheets except `_META` — the `_changed` file just needs to go through the same pipeline.

**Fix:** In `test_rebuild_e2e.py`, move the `_save_change()` call into the recalc test class (after recalc runs), or add a separate recalc call on the saved file.

**Cleanup:** Remove the dead comment `# Column widths` in `summary.py:248`.

**Files touched:** `tests/test_rebuild_e2e.py`, `budget_planner/summary.py` (dead comment)
**Tests:** Verify `_changed` file has non-default column widths after recalc.
**Risk:** Low — test-only change + dead code removal.

### Step 8: Extract threshold constants (MEDIUM — fixes M1)

**Add to `config.py` (or `layout.py`):**
```python
# Alert thresholds
TREND_OK_BAND = 0.05
TREND_OVER_THRESHOLD = 0.15
TREND_SLOW_THRESHOLD = -0.10
CAPACITY_AMBER_LOAD = 0.80
ACTUALS_AMBER_PCT_CEIL = 0.85
ACTUALS_AMBER_BREACH_MONTHS = 12
ACTUALS_RED_BREACH_MONTHS = 6
```

Replace magic numbers in formula strings with these constants.

**Files touched:** `config.py`, `trends.py`, `capacity.py`, `actuals.py`
**Tests:** Existing tests pass. Add test that thresholds are used in CF formulas.
**Risk:** Low — values unchanged.

### Step 9: Extract shared header builder (MEDIUM — fixes M3)

Create `_build_year_month_headers()` helper in a shared location (could be `styles.py` or new `layout.py`):

```python
def build_year_month_headers(ws, start_row, project_params, month_col_start,
                             extra_headers=None):
    """Build year + month header rows. Returns (year_row, month_row, next_row)."""
```

Replace the 4 copy-pasted header blocks with calls to this helper.

**Files touched:** `plan.py`, `capacity.py`, `actuals.py`, potentially `styles.py` or `layout.py`
**Tests:** Existing tests pass. Visual verification with LibreOffice.
**Risk:** Medium — touching header generation in all sheets requires careful testing.

---

## Dependency Graph

```
Step 1 (layout constants) ─┐
                            ├──> Step 8 (thresholds)      [after Step 1]
Step 2 (sheet names)  ──────┤
                            ├──> Step 9 (header builder)   [after Steps 1+2]
Step 3 (reader validation) ─┘

Step 4 (trends dynamic)     [independent — only trends.py]
Step 5 (reader matching)    [independent — only reader.py, pairs with Step 3]
Step 6 (protection fix)     [independent — only workbook.py]
Step 7 (recalc _changed)    [independent — only test file + dead comment]
```

**Parallel groups:**
- **Group A:** Steps 1 + 2 (both touch all builders — sequential, same files)
- **Group B:** Steps 3 + 5 (both in reader.py — sequential)
- **Group C:** Step 4 (only touches trends.py — can parallel with Group B)
- **Group D:** Steps 6 + 7 (workbook.py + test file — independent of all others, quick UX wins)
- **Group E:** Step 8 (touches config.py + 3 builders — after Step 1)
- **Group F:** Step 9 (touches 3+ builders — after Steps 1 + 2)

**Recommended order:**
1. Steps 6 + 7 first (quick UX wins, independent of everything)
2. Step 1 (layout constants) — unblocks most refactoring
3. Steps 2 + 4 in parallel (sheet names + trends dynamic)
4. Steps 3 + 5 sequential (reader hardening)
5. Step 8 (thresholds)
6. Step 9 (header builder)

---

## What's NOT worth hardening

- **Config sheet layout (rows 1-4 params, row 8 WPs):** This is a deliberate design, not accidental coupling. The Config sheet reader (Step 3) adding validation is sufficient — we don't need to make the row structure dynamically discoverable because Config is generated by our own builder.
- **Builder execution order:** The orchestrator in `workbook.py` correctly sequences builders. Adding runtime dependency checks would over-engineer for N=8 sheets.
- **Layout dict schemas:** Adding TypedDicts or runtime validation for every layout dict is overkill for ~2K LOC. Step 1's centralization makes the implicit contracts explicit enough.
- **Test fixtures:** Tests were already made data-independent (session 2026-04-03). No further hardening needed.
