"""
Budget Planner — Verification Script

Runs the same checks used during initial build to verify workbook correctness.
Run after every generation: python verify.py output/budget_plan.xlsx

Exit code 0 = all checks pass, 1 = failures found.
"""

import sys
import json
import subprocess
import openpyxl

def verify(filepath):
    errors = []
    warnings = []

    # ── 1. Formula recalculation ──
    print("Step 1: Recalculating formulas...")
    result = subprocess.run(
        ["python3", "scripts/recalc.py", filepath],
        capture_output=True, text=True
    )
    recalc = json.loads(result.stdout)
    if recalc["status"] != "success":
        errors.append(f"Recalc failed: {recalc['total_errors']} formula errors")
        for err_type, details in recalc.get("error_summary", {}).items():
            errors.append(f"  {err_type}: {details['count']} at {details['locations']}")
    else:
        print(f"  OK: {recalc['total_formulas']} formulas, 0 errors")

    # ── 2. Load workbook with calculated values ──
    print("Step 2: Loading workbook (data_only)...")
    wb = openpyxl.load_workbook(filepath, data_only=True)

    # ── 3. Check sheet existence ──
    print("Step 3: Checking sheet structure...")
    expected_sheets = ["_META", "Plan", "Summary"]
    for s in expected_sheets:
        if s not in wb.sheetnames:
            errors.append(f"Missing sheet: {s}")
        else:
            print(f"  OK: Sheet '{s}' exists")

    plan = wb["Plan"]
    summary = wb["Summary"]

    # ── 4. Check global parameters ──
    print("Step 4: Checking global parameters...")
    project_start = plan["B2"].value
    fte_rate = plan["E2"].value
    if project_start is None:
        errors.append("Plan!B2 (project_start) is empty")
    else:
        print(f"  Project start: {project_start}")
    if fte_rate is None or fte_rate <= 0:
        errors.append(f"Plan!E2 (fte_rate) is invalid: {fte_rate}")
    else:
        print(f"  FTE rate: €{fte_rate}")

    # ── 5. Check named ranges ──
    print("Step 5: Checking named ranges...")
    defined_names = {}
    for name in wb.defined_names:
        dn = wb.defined_names[name]
        defined_names[dn.name] = dn.attr_text
    for name, expected_ref in [("project_start", "Plan!$B$2"), ("fte_rate", "Plan!$E$2")]:
        if name not in defined_names:
            errors.append(f"Missing named range: {name}")
        elif defined_names[name] != expected_ref:
            errors.append(f"Named range '{name}' points to {defined_names[name]}, expected {expected_ref}")
        else:
            print(f"  OK: {name} -> {defined_names[name]}")

    # ── 6. Detect config table ──
    print("Step 6: Detecting config table...")
    config_header_row = None
    for r in range(1, 20):
        val = plan.cell(r, 1).value
        if val == "WP ID":
            config_header_row = r
            break
    if config_header_row is None:
        errors.append("Cannot find config table header row (looking for 'WP ID' in column A)")
        return errors, warnings

    config_first_wp = config_header_row + 1
    config_last_wp = config_first_wp
    while plan.cell(config_last_wp + 1, 1).value and not str(plan.cell(config_last_wp + 1, 2).value).upper().startswith("TOTAL"):
        config_last_wp += 1
    num_wps = config_last_wp - config_first_wp + 1
    print(f"  Found {num_wps} WPs in rows {config_first_wp}-{config_last_wp}")

    # ── 7. Validate config data ──
    print("Step 7: Validating config data...")
    wp_configs = []
    for r in range(config_first_wp, config_last_wp + 1):
        wp = {
            "row": r,
            "id": plan.cell(r, 1).value,
            "name": plan.cell(r, 2).value,
            "fte": plan.cell(r, 3).value,
            "start": plan.cell(r, 4).value,
            "end": plan.cell(r, 5).value,
            "active": plan.cell(r, 6).value,
            "budget_eur": plan.cell(r, 7).value,
        }
        wp_configs.append(wp)

        if wp["fte"] is None or wp["fte"] <= 0:
            errors.append(f"  {wp['id']}: Total FTE is invalid ({wp['fte']})")
        if wp["start"] is None or wp["end"] is None:
            errors.append(f"  {wp['id']}: Start/End month is empty")
        elif wp["start"] > wp["end"]:
            errors.append(f"  {wp['id']}: Start month ({wp['start']}) > End month ({wp['end']})")
        elif wp["active"] != wp["end"] - wp["start"] + 1:
            errors.append(f"  {wp['id']}: Active months formula wrong: {wp['active']} != {wp['end'] - wp['start'] + 1}")

        expected_budget = wp["fte"] * fte_rate if wp["fte"] and fte_rate else None
        if expected_budget and wp["budget_eur"] and abs(wp["budget_eur"] - expected_budget) > 1:
            errors.append(f"  {wp['id']}: Budget EUR {wp['budget_eur']} != FTE {wp['fte']} * rate {fte_rate} = {expected_budget}")

    if not errors:
        print(f"  OK: All {num_wps} WP configs valid")

    # ── 8. Find planned and actual blocks ──
    print("Step 8: Locating Planned and Actual blocks...")
    planned_header = None
    actuals_header = None
    for r in range(config_last_wp + 1, plan.max_row + 1):
        val = str(plan.cell(r, 1).value or "")
        if "PLANNED" in val.upper():
            planned_header = r
        elif "ACTUAL" in val.upper():
            actuals_header = r

    if planned_header is None:
        errors.append("Cannot find PLANNED FTE block header")
        return errors, warnings
    if actuals_header is None:
        errors.append("Cannot find ACTUAL FTE block header")
        return errors, warnings

    print(f"  Planned block header: row {planned_header}")
    print(f"  Actual block header: row {actuals_header}")

    # ── 9. Validate planned block — variance check ──
    print("Step 9: Checking planned FTE validation columns...")
    # Find the first WP row in planned block (3 rows after header: header, year row, month row)
    plan_first_wp = planned_header + 3
    TOLERANCE = 0.01

    for i, wp in enumerate(wp_configs):
        r = plan_first_wp + i
        # Validation columns are the last 3 populated columns
        # Find them by checking the row
        sum_val = None
        budget_val = None
        variance_val = None

        # Scan from right to find validation cols
        for c in range(plan.max_column, 2, -1):
            v = plan.cell(r, c).value
            if v is not None:
                if variance_val is None:
                    variance_val = v
                elif budget_val is None:
                    budget_val = v
                elif sum_val is None:
                    sum_val = v
                    break

        if sum_val is not None and budget_val is not None and variance_val is not None:
            if abs(variance_val) > TOLERANCE:
                warnings.append(f"  {wp['id']}: Planned variance = {variance_val:.4f} (rounding, adjust manually)")
            else:
                pass  # OK
        else:
            errors.append(f"  {wp['id']}: Cannot read validation columns at row {r}")

    planned_variance_issues = [w for w in warnings if "Planned variance" in w]
    if not planned_variance_issues:
        print(f"  OK: All planned FTE sums match budgets exactly")
    else:
        print(f"  {len(planned_variance_issues)} WPs have minor rounding variance (see warnings)")

    # ── 10. Validate actual block — should be all zeros initially ──
    print("Step 10: Checking actual FTE block (should be zeros initially)...")
    act_first_wp = actuals_header + 3

    for i, wp in enumerate(wp_configs):
        r = act_first_wp + i
        sum_val = None
        for c in range(plan.max_column, 2, -1):
            v = plan.cell(r, c).value
            if v is not None:
                # Third from right is sum
                remaining = v
                c2 = c - 1
                budget_v = plan.cell(r, c2).value
                c3 = c2 - 1
                sum_v = plan.cell(r, c3).value
                if sum_v is not None and sum_v != 0:
                    warnings.append(f"  {wp['id']}: Actual sum is {sum_v}, expected 0 for fresh workbook")
                break

    act_warnings = [w for w in warnings if "Actual sum" in w]
    if not act_warnings:
        print(f"  OK: All actuals are zero")

    # ── 11. Check inactive months are empty/zero ──
    print("Step 11: Checking inactive month cells...")
    MONTH_COL_START = 3
    inactive_violations = 0

    for i, wp in enumerate(wp_configs):
        plan_r = plan_first_wp + i
        act_r = act_first_wp + i
        for m in range(60):
            month_num = m + 1
            col = MONTH_COL_START + m
            if month_num < wp["start"] or month_num > wp["end"]:
                for check_row in [plan_r, act_r]:
                    val = plan.cell(check_row, col).value
                    if val is not None and val != "" and val != 0:
                        inactive_violations += 1
                        errors.append(f"  {wp['id']} M{month_num} row {check_row}: value '{val}' in inactive month")

    if inactive_violations == 0:
        print(f"  OK: No data in inactive month cells")

    # ── 12. Check Summary sheet ──
    print("Step 12: Checking Summary sheet...")
    sum_wp_count = 0
    for r in range(1, summary.max_row + 1):
        val = summary.cell(r, 1).value
        if val and str(val).startswith("WP") and val != "WP ID":
            sum_wp_count += 1

    if sum_wp_count != num_wps:
        errors.append(f"Summary has {sum_wp_count} WP rows, expected {num_wps}")
    else:
        print(f"  OK: {sum_wp_count} WP rows in Summary")

    # ── 13. Check _META sheet exists and has content ──
    print("Step 13: Checking _META sheet...")
    meta = wb["_META"]
    meta_content = str(meta["A1"].value or "")
    if len(meta_content) < 100:
        errors.append("_META sheet content is too short or empty")
    elif "MODIFICATION GUIDE" not in meta_content:
        warnings.append("_META sheet may be missing modification instructions")
    else:
        print(f"  OK: _META sheet has {len(meta_content)} chars of documentation")

    return errors, warnings


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python verify.py <path_to_xlsx>")
        sys.exit(1)

    filepath = sys.argv[1]
    print(f"Verifying: {filepath}\n")

    errors, warnings = verify(filepath)

    print("\n" + "=" * 50)
    if warnings:
        print(f"\nWARNINGS ({len(warnings)}):")
        for w in warnings:
            print(f"  ⚠ {w}")

    if errors:
        print(f"\nERRORS ({len(errors)}):")
        for e in errors:
            print(f"  ✗ {e}")
        print("\nRESULT: FAIL")
        sys.exit(1)
    else:
        print("\nRESULT: PASS")
        sys.exit(0)
