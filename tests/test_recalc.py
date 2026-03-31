"""LibreOffice recalculation tests — auto-skip if unavailable."""


def test_recalculation_zero_errors(recalc_result):
    assert recalc_result["status"] == "success", \
        f"Recalc found {recalc_result['total_errors']} errors"
    assert recalc_result["total_errors"] == 0


def test_formula_count_positive(recalc_result):
    assert recalc_result["total_formulas"] > 0, "No formulas found in workbook"
