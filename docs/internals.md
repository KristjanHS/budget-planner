# Internals

Technical reference for the workbook generation engine: builder pipeline, reader/rebuild, layout infrastructure, conditional formatting, and cross-sheet references.

---

## Sheet Builders

Each sheet is produced by a stateless builder function. Builders receive data as parameters and return a layout dict of row/column positions consumed by downstream builders.

### Builder Interface

```python
def build_<sheet>_sheet(ws, wb, ..., *, project_params=None, work_packages=None, ...):
    # ws: openpyxl Worksheet
    # wb: parent Workbook (named ranges, cross-sheet access)
    # ...: upstream layout dicts
    # Returns: layout dict with row/column numbers for downstream builders
```

### Builder Dependency Chain

| Builder | Needs from upstream |
|---------|-------------------|
| config_sheet | (none) |
| plan | config_layout.wp_first_row |
| capacity | config_layout.person_rows, plan_layout |
| plan (patch) | capacity_layout.person_groups.wp_rows |
| actuals | plan_layout |
| summary | plan_layout + actuals_layout |
| trends | all upstream layouts |
| meta | plan_layout |

### Per-Sheet Notes

**Config** (`config_sheet.py`): Rows 1-4 project params, then WP table (`tbl_WorkPackages`), then Team Roster section. Creates named ranges `project_start` (B2) and `fte_rate` (B4). Returns `config_layout` with `wp_first_row`, `wp_last_row`, `person_rows`.

**Plan** (`plan.py`): Two-phase build. Phase 1 creates structure with empty FTE cells. Phase 2 (`_apply_planned_formulas()`) fills cells with `=SUM(Capacity!...)` after Capacity is built. Has validation columns: Sum, Budget FTE, Variance.

**Capacity** (`capacity.py`): Most complex builder. Creates per-person groups: WP rows (blue input), external project rows, TOTAL (SUM formula), REMAINING (Max FTE - TOTAL). Every person gets all WP rows. Returns `capacity_layout` with `person_groups` dict.

**Actuals** (`actuals.py`): Similar to Plan but with EUR input cells. Extra validation columns: Cum. EUR (SUMPRODUCT up to current_month), % Ceiling, Breach In (months until ceiling exhausted).

**Summary** (`summary.py`): Calendar-year groups with Plan FTE, Plan EUR, Actual EUR, Variance EUR sub-columns. Budget Overview section: Personnel Actual, Total Claimable (+25% indirect), Personnel Ceiling, Remaining, % Consumed.

**Trends** (`trends.py`): Dashboard with WP Burn Rate & Exhaustion table, Person Load table, and cumulative EUR LineChart. Hidden helper rows below person load hold chart data series.

---

## Reader and Rebuild

`reader.py` extracts structured data from an existing workbook so `rebuild_workbook()` can regenerate all sheets while preserving user-entered values. This is the most fragile module.

### Formula Chain Resolution

The reader must follow cross-sheet formula references to resolve actual values:

```
Actuals cell =Plan!A6 -> Plan!A6 =Config!A8 -> Config!A8 = "WP1"
```

`_resolve_plan_ref()` handles this, following up to 2 hops.

### Extraction Functions

| Function | Extracts | Key technique |
|----------|----------|---------------|
| `read_config_sheet(ws)` | Project params (rows 1-4), WP table (rows 8+) | Direct cell reads |
| `read_team_from_config(ws)` | Team roster (name + max_fte) | Scans from "TEAM ROSTER" header |
| `read_capacity_sheet(ws, ...)` | External commitments, per-person monthly FTE | Person header detection by bold + fill `2F5496` |
| `read_actuals_values(ws, ...)` | EUR per WP per month | Formula chain resolution |

### What Rebuild Preserves

All sheets are regenerated from scratch. Only these user-entered values survive:

| Data | Source in old workbook |
|------|----------------------|
| Project params, WP definitions, team roster, grade rates, reporting periods | Config sheet sections |
| Person-WP assignments, external commitments | Config sheet PERSON-WP ASSIGNMENTS (B2) + EXTERNAL COMMITMENTS (B3) sections |
| Monthly EUR values per WP | Actuals data cells |

Everything else (formulas, styling, CF, charts, _META/Plan/Summary/Trends) is overwritten.

### Preserved Data Structures

```python
# Capacity FTE: (person_name, wp_id) -> [monthly values]
preserved_capacity: dict[tuple[str, str], list[float | None]]

# Actuals EUR: wp_id -> [monthly EUR values]
preserved_actuals: dict[str, list[float | None]]
```

### Orphan Detection

- Person removed from team: Capacity entries dropped (warning)
- WP removed from config: Capacity + Actuals entries dropped (warning)
- External commitment references missing team member: hard error

All orphaned data is logged and included in `_META` sheet's LAST REBUILD LOG.

### Row-Span Resolution (named-range-primary)

Reader functions for Config tables (`read_team_from_config`,
`read_assignments_from_config`, `read_externals_from_config`) are **named-range
primary**: they resolve `roster_ftes` / `assign_persons` / `ext_persons` via
`wb.defined_names` as the authoritative row span, then iterate within and
tolerate internal blank rows. Excel auto-extends these named ranges when the
user inserts rows inside them, so the reader's bounds stay in sync with live
data after hand-edits. Only when a named range is missing (legacy workbooks)
does the reader fall back to sentinel-based scanning that stops at the first
blank. See `_iter_named_range_rows` in `reader.py` and shape-drift tests in
`tests/test_rebuild_shape_drift.py` / `tests/test_rebuild_shapes.py`.

### Known Fragility

1. **Person header detection** relies on specific font/fill styling. Manual restyling breaks detection.
2. **Formula chain resolution** assumes patterns like `=Config!A8`. Custom formulas are not followed.
3. **Sentinel-scan sections** (WP BUDGET, WP SCHEDULE, RATE BY GRADE, reporting periods) still stop at the first blank in column A. Use filled rows when hand-editing these, or add named ranges to extend the same treatment as roster/assignments/externals.

---

## Layout and Styling

`layout.py` and `styles.py` centralize constants that every builder imports.

### Month Column Positioning

All sheets with monthly data share the same column structure:

```
Col A (1) = WP ID / Person name
Col B (2) = Label
Col C (3) = Month 1              (MONTH_COL_START = 3)
...
Col N     = Month N              (month_col_end(N))
Col N+1   = Sum                  (val_columns(N)["sum"])
Col N+2   = Budget               (val_columns(N)["budget"])
Col N+3   = Variance             (val_columns(N)["diff"])
Col N+4   = Cum. EUR             (Actuals only)
Col N+5   = % Ceiling            (Actuals only)
Col N+6   = Breach In            (Actuals only)
```

Validation columns are offsets from the last month column, so changing project duration auto-repositions them.

### Alert Thresholds (layout.py)

| Constant | Value | Used by |
|----------|-------|---------|
| `CAPACITY_AMBER_LOAD` | 0.80 | Capacity + Trends person load |
| `ACTUALS_AMBER_PCT_CEIL` | 0.85 | Actuals % ceiling |
| `ACTUALS_AMBER_BREACH_MONTHS` | 12 | Actuals breach column |
| `ACTUALS_RED_BREACH_MONTHS` | 6 | Actuals breach column |
| `TREND_OK_BAND` | 0.05 | Trends burn delta |
| `TREND_OVER_THRESHOLD` | 0.15 | Trends burn delta |
| `TREND_SLOW_THRESHOLD` | -0.10 | Trends burn delta |

### style_cell() Auto-Unlock

When font color is blue (`0000FF`), `style_cell()` automatically sets `Protection(locked=False)`. When sheet protection is re-enabled, all blue input cells become editable with no extra code.

### Column Grouping

`apply_year_groups()` creates collapsible column groups by calendar year. The group containing `current_month` is expanded; all others are collapsed.

---

## Conditional Formatting Rules

### Config Sheet
- WP start/end columns: **yellow** when start month > end month

### Plan Sheet
- Planned variance column: **yellow** when `ABS(value) > 0.01`
- Actual remaining column: **red** when `value < 0`

### Team Sheet
- WP share SUM row: **yellow** when `ABS(SUM - 1.0) > 0.001`

### Capacity Sheet
- TOTAL row per person: **amber** when `TOTAL > 0.8 * Max FTE`, **red** when `TOTAL > Max FTE`
- REMAINING row per person: **red** when `REMAINING < 0`

### Actuals Sheet
- Negative variance: **red**
- % Ceiling: **amber** > 85%, **red** > 100%
- Breach In: **amber** < 12 months, **red** < 6 months

### Summary Sheet
- Remaining EUR column: **red** when `value < 0`

---

## Cross-Sheet Reference Map

| Target Cell | References | Via |
|-------------|-----------|-----|
| **Config -> Plan** | | |
| `project_start` | Config!$B$2 | Named range |
| `fte_rate` | Config!$B$4 | Named range |
| **Plan -> Config** | | |
| Plan global cells (start, rate) | Config params | Green formula |
| **Plan -> Capacity** | | |
| Plan Planned FTE (monthly) | Capacity per-person WP rows | `=SUM(Capacity!{cells})` green formula |
| **Capacity -> Plan** | | |
| Capacity WP ID/Name | Plan config rows | `=Plan!A{row}` / `=Plan!B{row}` green formula |
| **Capacity -> Config** | | |
| Person header | Config roster name | `=Config!A{row}` formula |
| REMAINING | Config Max FTE | `=Config!B{row}-{TOTAL}` formula |
| **Summary -> Plan** | | |
| Summary WP ID/Name | Plan config rows | Green formula |
| Year N Plan/Act FTE | Plan monthly ranges | `=SUM(Plan!{range})` |
| Remaining EUR | Plan ceiling - Actual total | Formula |
| **Trends -> multiple** | | |
| WP start/end | Config refs | Formula |
| Cumulative/ceiling | Actuals refs | Formula |
| TOTAL rows | Capacity refs | Formula |
| Planned FTE (chart) | Plan refs | Formula |
