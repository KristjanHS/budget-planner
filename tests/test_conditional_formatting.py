"""Tests for conditional formatting rules across all sheets."""

import pytest
import openpyxl

from budget_planner.workbook import generate_workbook
from budget_planner.config import TEAM, PROJECT_MONTHS


# ── Fixtures ──

@pytest.fixture(scope="module")
def workbook(tmp_path_factory):
    """Generate a full demo workbook and load it."""
    out = tmp_path_factory.mktemp("cf") / "test.xlsx"
    generate_workbook(str(out))
    return openpyxl.load_workbook(str(out), data_only=False)


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
            if rule.dxf and rule.dxf.fill and rule.dxf.fill.fgColor:
                fg = str(rule.dxf.fill.fgColor.rgb)
            results.append((sqref, formulas, fg))
    return results


# ── Capacity sheet conditional formatting ──

class TestCapacityConditionalFormatting:
    """Capacity sheet has 3 CF rules per person (amber TOTAL, red TOTAL, red REMAINING)."""

    def test_capacity_rule_count(self, workbook):
        """3 rules per person x len(TEAM) persons = expected total."""
        ws = workbook["Capacity"]
        rules = _collect_cf_rules(ws)
        expected = 3 * len(TEAM)
        assert len(rules) == expected, (
            f"Expected {expected} CF rules on Capacity, got {len(rules)}"
        )

    @pytest.mark.parametrize("person_idx", range(len(TEAM)))
    def test_capacity_total_amber_rule(self, workbook, person_idx):
        """Each person's TOTAL row has an amber fill for >0.8 * Max FTE."""
        ws = workbook["Capacity"]
        rules = _collect_cf_rules(ws)
        # Rules are added in order: amber TOTAL, red TOTAL, red REMAINING per person
        rule_idx = person_idx * 3  # amber is first of each triple
        sqref, formulas, fg_color = rules[rule_idx]
        assert len(formulas) == 1
        assert ">0.8*Team!$B$" in formulas[0], (
            f"Amber formula should reference >0.8*Team!$B$<row>, got: {formulas[0]}"
        )
        assert fg_color is not None and fg_color.endswith("FFD966"), (
            f"Expected AMBER_FILL (FFD966), got: {fg_color}"
        )

    @pytest.mark.parametrize("person_idx", range(len(TEAM)))
    def test_capacity_total_red_rule(self, workbook, person_idx):
        """Each person's TOTAL row has a red fill for > Max FTE."""
        ws = workbook["Capacity"]
        rules = _collect_cf_rules(ws)
        rule_idx = person_idx * 3 + 1  # red TOTAL is second of each triple
        sqref, formulas, fg_color = rules[rule_idx]
        assert len(formulas) == 1
        assert ">Team!$B$" in formulas[0], (
            f"Red TOTAL formula should reference >Team!$B$<row>, got: {formulas[0]}"
        )
        # Must NOT contain the 0.8 multiplier (that's the amber rule)
        assert "0.8" not in formulas[0], (
            f"Red TOTAL formula should not have 0.8 multiplier: {formulas[0]}"
        )
        assert fg_color is not None and fg_color.endswith("FFC7CE"), (
            f"Expected RED_FILL (FFC7CE), got: {fg_color}"
        )

    @pytest.mark.parametrize("person_idx", range(len(TEAM)))
    def test_capacity_remaining_red_rule(self, workbook, person_idx):
        """Each person's REMAINING row has a red fill for < 0."""
        ws = workbook["Capacity"]
        rules = _collect_cf_rules(ws)
        rule_idx = person_idx * 3 + 2  # red REMAINING is third of each triple
        sqref, formulas, fg_color = rules[rule_idx]
        assert len(formulas) == 1
        assert "<0" in formulas[0], (
            f"REMAINING formula should contain '<0', got: {formulas[0]}"
        )
        assert fg_color is not None and fg_color.endswith("FFC7CE"), (
            f"Expected RED_FILL (FFC7CE), got: {fg_color}"
        )


# ── Summary sheet conditional formatting ──

class TestSummaryConditionalFormatting:
    """Summary has year variance (red), project total variance (red), remaining (amber + red)."""

    def test_summary_year_variance_red(self, workbook):
        """Each year's Variance EUR column has a red fill for < 0."""
        ws = workbook["Summary"]
        rules = _collect_cf_rules(ws)
        num_years = (PROJECT_MONTHS + 11) // 12
        # Year variance rules are the first num_years rules
        for y in range(num_years):
            sqref, formulas, fg_color = rules[y]
            assert len(formulas) == 1
            assert "<0" in formulas[0], (
                f"Year {y+1} variance formula should contain '<0', got: {formulas[0]}"
            )
            assert fg_color is not None and fg_color.endswith("FFC7CE"), (
                f"Year {y+1}: expected RED_FILL (FFC7CE), got: {fg_color}"
            )

    def test_summary_project_total_variance_red(self, workbook):
        """Project total Variance EUR column has a red fill for < 0."""
        ws = workbook["Summary"]
        rules = _collect_cf_rules(ws)
        num_years = (PROJECT_MONTHS + 11) // 12
        # Project total variance is added after all year variance rules
        sqref, formulas, fg_color = rules[num_years]
        assert len(formulas) == 1
        assert "<0" in formulas[0], (
            f"Project total variance formula should contain '<0', got: {formulas[0]}"
        )
        assert fg_color is not None and fg_color.endswith("FFC7CE"), (
            f"Expected RED_FILL (FFC7CE), got: {fg_color}"
        )

    def test_summary_remaining_amber(self, workbook):
        """Remaining cell has amber fill when < 10% of ceiling."""
        ws = workbook["Summary"]
        rules = _collect_cf_rules(ws)
        num_years = (PROJECT_MONTHS + 11) // 12
        # Remaining amber is after year variance + project total variance
        amber_idx = num_years + 1
        sqref, formulas, fg_color = rules[amber_idx]
        assert len(formulas) == 1
        assert "*0.1" in formulas[0], (
            f"Remaining amber formula should contain '*0.1', got: {formulas[0]}"
        )
        assert fg_color is not None and fg_color.endswith("FFD966"), (
            f"Expected AMBER_FILL (FFD966), got: {fg_color}"
        )

    def test_summary_remaining_red(self, workbook):
        """Remaining cell has red fill when < 0."""
        ws = workbook["Summary"]
        rules = _collect_cf_rules(ws)
        num_years = (PROJECT_MONTHS + 11) // 12
        # Remaining red is the last summary rule
        red_idx = num_years + 2
        sqref, formulas, fg_color = rules[red_idx]
        assert len(formulas) == 1
        assert "<0" in formulas[0], (
            f"Remaining red formula should contain '<0', got: {formulas[0]}"
        )
        assert fg_color is not None and fg_color.endswith("FFC7CE"), (
            f"Expected RED_FILL (FFC7CE), got: {fg_color}"
        )


# ── Plan sheet conditional formatting ──

class TestPlanConditionalFormatting:
    """Plan sheet has yellow fill on variance column when ABS() > 0.01."""

    def test_plan_variance_yellow(self, workbook):
        """Planned variance column has yellow fill for nonzero differences."""
        ws = workbook["Plan"]
        rules = _collect_cf_rules(ws)
        # Plan has exactly 1 CF rule (yellow on variance)
        assert len(rules) == 1, f"Expected 1 CF rule on Plan, got {len(rules)}"
        sqref, formulas, fg_color = rules[0]
        assert len(formulas) == 1
        assert "ABS(" in formulas[0] and ">0.01" in formulas[0], (
            f"Plan variance formula should contain 'ABS(...)>0.01', got: {formulas[0]}"
        )
        assert fg_color is not None and fg_color.endswith("FFFF00"), (
            f"Expected YELLOW_FILL (FFFF00), got: {fg_color}"
        )


# ── Actuals sheet conditional formatting ──

class TestActualsConditionalFormatting:
    """Actuals sheet has red fill on variance column when < 0."""

    def test_actuals_variance_red(self, workbook):
        """Actuals variance column has red fill for negative values."""
        ws = workbook["Actuals"]
        rules = _collect_cf_rules(ws)
        # Actuals has exactly 1 CF rule (red on variance)
        assert len(rules) == 1, f"Expected 1 CF rule on Actuals, got {len(rules)}"
        sqref, formulas, fg_color = rules[0]
        assert len(formulas) == 1
        assert "<0" in formulas[0], (
            f"Actuals variance formula should contain '<0', got: {formulas[0]}"
        )
        assert fg_color is not None and fg_color.endswith("FFC7CE"), (
            f"Expected RED_FILL (FFC7CE), got: {fg_color}"
        )
