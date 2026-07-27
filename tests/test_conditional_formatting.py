"""Tests for conditional formatting rules across all sheets."""

import pytest
import openpyxl

from budget_planner.layout import _calendar_year_groups
from budget_planner.config import TEAM, PROJECT_MONTHS, PROJECT_START, DEMO_CAPACITY_ASSIGNMENTS


# ── Fixtures ──


@pytest.fixture(scope="module")
def workbook(assigned_wb_path):
    return openpyxl.load_workbook(assigned_wb_path, data_only=False)


# ── Helpers ──


def _collect_cf_rules(ws):
    """Collect all CF rules from a worksheet.

    Returns a list of (sqref, formula, fill_fgColor) tuples.
    """
    results = []
    for cf in ws.conditional_formatting:
        sqref = str(cf.sqref)
        for rule in cf.rules:
            formulas = list(rule.formula) if rule.formula else []
            fg = None
            if rule.dxf and rule.dxf.fill:
                # CF fills use bgColor (DXF context); fall back to fgColor
                fill = rule.dxf.fill
                if fill.bgColor and str(fill.bgColor.rgb) != "00000000":
                    fg = str(fill.bgColor.rgb)
                elif fill.fgColor and str(fill.fgColor.rgb) != "00000000":
                    fg = str(fill.fgColor.rgb)
            results.append((sqref, formulas, fg))
    return results


# ── Capacity sheet conditional formatting ──


class TestCapacityConditionalFormatting:
    """Capacity sheet CF: yellow per WP variance + 2 per person (amber/red TOTAL).

    REMAINING + EU TOTAL rows dropped from Capacity (2026-04-28); per-year EU
    ceiling CF on Actuals-Hours deleted in the posted-cumulative redesign
    (2026-05-09); EU compliance is now a manual check. REMAINING < 0 CF
    folded onto TOTAL > Max FTE (still red, single rule per person)."""

    def test_capacity_rule_count(self, workbook):
        """Planned-side CF + Actual-view band CF + simplified to-date CF.

        Planned side:
          - yellow per WP row (variance < 0)
          - 2 per person (TOTAL amber+red — folded REMAINING < 0 onto
            TOTAL > Max FTE, same red rule)
        Actual-view band (right-side grouped, raw monthly):
          - 1 variance-red rule per (WP row + external row) per person
          - 2 per person (utilisation amber+red on Actual TOTAL row)
        To-date block (2 cols):
          - 1 per person (Variance PM red < 0 on TOTAL only).
        """
        from budget_planner.config import EXTERNAL_COMMITMENTS

        ws = workbook["Capacity"]
        rules = _collect_cf_rules(ws)
        total_wp_rows = len(DEMO_CAPACITY_ASSIGNMENTS)
        total_ext_rows = len(EXTERNAL_COMMITMENTS)
        planned = total_wp_rows + 2 * len(TEAM)
        actual_band = (
            total_wp_rows
            + total_ext_rows  # variance reds
            + 2 * len(TEAM)  # utilisation amber+red
        )
        todate = 1 * len(TEAM)  # Variance PM < 0 on TOTAL only
        expected = planned + actual_band + todate
        assert len(rules) == expected, f"Expected {expected} CF rules on Capacity, got {len(rules)}"

    @pytest.mark.parametrize("person_idx", range(len(TEAM)))
    def test_capacity_total_amber_rule(self, workbook, person_idx):
        """Each person's TOTAL row has an amber fill for >0.8 * Max FTE."""
        ws = workbook["Capacity"]
        rules = _collect_cf_rules(ws)
        # Rules per person: N yellow (one per WP row) + amber TOTAL + red TOTAL.
        # Find the amber rule by scanning for FFD966 fill with >0.8 pattern.
        amber_rules = [
            (i, r)
            for i, r in enumerate(rules)
            if r[2] is not None and r[2].endswith("FFD966") and any(">0.8*Config!$B$" in f for f in r[1])
        ]
        assert len(amber_rules) >= person_idx + 1, (
            f"Expected at least {person_idx + 1} amber TOTAL rules, got {len(amber_rules)}"
        )
        _, (sqref, formulas, fg_color) = amber_rules[person_idx]
        assert ">0.8*Config!$B$" in formulas[0], (
            f"Amber formula should reference >0.8*Config!$B$<row>, got: {formulas[0]}"
        )

    @pytest.mark.parametrize("person_idx", range(len(TEAM)))
    def test_capacity_total_red_rule(self, workbook, person_idx):
        """Each person's TOTAL row has a red fill for > Max FTE.

        Folds the prior REMAINING-row red CF onto TOTAL: a monthly TOTAL >
        Max FTE is by definition REMAINING < 0.
        """
        ws = workbook["Capacity"]
        rules = _collect_cf_rules(ws)
        # Find red TOTAL rules: FFC7CE fill with >Config! pattern (no 0.8, no SUM).
        red_total_rules = [
            (i, r)
            for i, r in enumerate(rules)
            if r[2] is not None
            and r[2].endswith("FFC7CE")
            and any(">Config!$B$" in f for f in r[1])
            and not any("0.8" in f for f in r[1])
            and not any("SUM(" in f for f in r[1])
        ]
        assert len(red_total_rules) >= person_idx + 1, (
            f"Expected at least {person_idx + 1} red TOTAL rules, got {len(red_total_rules)}"
        )
        _, (sqref, formulas, fg_color) = red_total_rules[person_idx]
        assert ">Config!$B$" in formulas[0]
        assert "0.8" not in formulas[0]


# ── Summary sheet conditional formatting ──


class TestSummaryConditionalFormatting:
    """Summary has year variance (red), project total variance (red), remaining (amber + red)."""

    def test_summary_year_variance_red(self, workbook):
        """Each year's Variance EUR column has a red fill for < 0."""
        ws = workbook["Summary"]
        rules = _collect_cf_rules(ws)
        num_years = len(list(_calendar_year_groups(PROJECT_START, PROJECT_MONTHS)))
        # Match by: RED_FILL + <0 formula + range sqref (contains ':').
        # Exclude the Task 10 Forecast EUR rule which uses "<0.8*...".
        variance_red_rules = [
            r
            for r in rules
            if r[2] is not None
            and r[2].endswith("FFC7CE")
            and any("<0" in f for f in r[1])
            and not any("0.8*" in f for f in r[1])
            and ":" in r[0]
        ]
        # Year variance + project total variance = num_years + 1 total range-based red rules
        assert len(variance_red_rules) >= num_years, (
            f"Expected at least {num_years} red variance range rules, got {len(variance_red_rules)}"
        )
        for rule in variance_red_rules[:num_years]:
            sqref, formulas, fg_color = rule
            assert len(formulas) == 1
            assert "<0" in formulas[0], f"Year variance formula should contain '<0', got: {formulas[0]}"

    def test_summary_project_total_variance_red(self, workbook):
        """Project total Variance EUR column has a red fill for < 0."""
        ws = workbook["Summary"]
        rules = _collect_cf_rules(ws)
        num_years = len(list(_calendar_year_groups(PROJECT_START, PROJECT_MONTHS)))
        # Match by: RED_FILL + <0 formula + range sqref (contains ':').
        # Exclude the Task 10 Forecast EUR rule which uses "<0.8*...".
        variance_red_rules = [
            r
            for r in rules
            if r[2] is not None
            and r[2].endswith("FFC7CE")
            and any("<0" in f for f in r[1])
            and not any("0.8*" in f for f in r[1])
            and ":" in r[0]
        ]
        # Budget Overview: num_years + 1 (year variance + project total)
        # RP Summary: num_rps + 1 (per-RP variance + project total)
        from budget_planner.config import REPORTING_PERIODS

        num_rps = len(REPORTING_PERIODS)
        expected = (num_years + 1) + (num_rps + 1)
        assert len(variance_red_rules) == expected, (
            f"Expected {expected} red variance range rules "
            f"(year + total + RP + RP total), got {len(variance_red_rules)}"
        )

    def test_summary_forecast_eur_red(self, workbook):
        """Project-end Forecast EUR column has a single red CF when
        Forecast < 0.8 × project total Plan EUR. The posted-cumulative
        rewrite collapsed per-year Forecast columns into one project-end
        column, so only ONE red rule remains (not num_years + 1).
        """
        ws = workbook["Summary"]
        rules = _collect_cf_rules(ws)
        # Match by: RED_FILL + "0.8*" + range sqref (contains ':')
        fcst_red_rules = [
            r
            for r in rules
            if r[2] is not None and r[2].endswith("FFC7CE") and any("0.8*" in f for f in r[1]) and ":" in r[0]
        ]
        assert len(fcst_red_rules) == 1, (
            f"Expected exactly 1 red Forecast EUR CF rule on the project-end column, got {len(fcst_red_rules)}"
        )
        # Empty-sentinel guard required: posted_idx=0 returns "" from the
        # forecast formula, and `"" < 0.8*X` would fire spuriously without
        # the `<>""` clause.
        assert any('<>""' in f for f in fcst_red_rules[0][1]), (
            f'Project-end Forecast red CF must guard against empty-sentinel with `<>""`, got: {fcst_red_rules[0][1]}'
        )

    def test_summary_remaining_amber(self, workbook):
        """Remaining cell has amber fill when < 10% of ceiling."""
        ws = workbook["Summary"]
        rules = _collect_cf_rules(ws)
        # Match by: AMBER_FILL (only one amber rule on Summary)
        amber_rules = [
            r for r in rules if r[2] is not None and r[2].endswith("FFD966") and any("*0.1" in f for f in r[1])
        ]
        assert len(amber_rules) == 1, f"Expected exactly 1 amber remaining rule, got {len(amber_rules)}"
        sqref, formulas, fg_color = amber_rules[0]
        assert len(formulas) == 1
        assert "*0.1" in formulas[0], f"Remaining amber formula should contain '*0.1', got: {formulas[0]}"

    def test_summary_remaining_red(self, workbook):
        """Remaining cell has red fill when < 0."""
        ws = workbook["Summary"]
        rules = _collect_cf_rules(ws)
        # Match by: RED_FILL + <0 formula + single-cell sqref (no ':')
        red_single_cell_rules = [
            r
            for r in rules
            if r[2] is not None and r[2].endswith("FFC7CE") and any("<0" in f for f in r[1]) and ":" not in r[0]
        ]
        assert len(red_single_cell_rules) == 1, (
            f"Expected exactly 1 red single-cell rule (remaining), got {len(red_single_cell_rules)}"
        )
        sqref, formulas, fg_color = red_single_cell_rules[0]
        assert len(formulas) == 1
        assert "<0" in formulas[0], f"Remaining red formula should contain '<0', got: {formulas[0]}"


# ── Plan sheet conditional formatting ──


class TestPlanConditionalFormatting:
    """Plan sheet has yellow fill on variance column when ABS() > 0.01."""

    def test_plan_variance_red(self, workbook):
        """Planned variance column has red fill for negative values."""
        ws = workbook["Plan"]
        rules = _collect_cf_rules(ws)
        # Plan has exactly 1 CF rule (red on negative variance)
        assert len(rules) == 1, f"Expected 1 CF rule on Plan, got {len(rules)}"
        sqref, formulas, fg_color = rules[0]
        assert len(formulas) == 1
        assert "<0" in formulas[0], f"Plan variance formula should contain '<0', got: {formulas[0]}"
        assert fg_color is not None and fg_color.endswith("FFC7CE"), f"Expected RED_FILL (FFC7CE), got: {fg_color}"


# ── Actuals sheet conditional formatting ──


class TestActualsConditionalFormatting:
    """Actuals sheet has red fill on variance column when < 0."""

    def test_actuals_variance_red(self, workbook):
        """Actuals variance column has red fill for negative values.

        Posted-cumulative redesign (2026-05-09) deleted the forecast
        monitoring block on Actuals (5 -i + 5 +i CFs gone), reducing
        the total rule count from 21 → 11.

        Post-redesign: 11 rules
            - Actual-EUR -i: variance red, %ceiling amber/red, breach amber/red = 5
            - Actual-EUR +i mirror: variance red, %ceiling amber/red, breach amber/red = 5
            - Reconciliation grey-fill (Task 8): future-month grey = 1
        """
        ws = workbook["Actuals"]
        rules = _collect_cf_rules(ws)
        assert len(rules) == 11, f"Expected 11 CF rules on Actuals, got {len(rules)}"
        # First rule is variance red on the Actual-EUR -i block
        sqref, formulas, fg_color = rules[0]
        assert len(formulas) == 1
        assert "<0" in formulas[0], f"Actuals variance formula should contain '<0', got: {formulas[0]}"
        assert fg_color is not None and fg_color.endswith("FFC7CE"), f"Expected RED_FILL (FFC7CE), got: {fg_color}"
