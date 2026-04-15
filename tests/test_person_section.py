from budget_planner.person_section import plan_person_layout


def test_plan_person_layout_assigns_sequential_rows():
    layout = plan_person_layout(
        start_row=10,
        wp_ids=["WP1", "WP2"],
        external_names=["TeachingX"],
        include_remaining=True,
        include_eu_total=True,
    )
    assert layout.header_row == 10
    assert layout.wp_rows == {"WP1": 11, "WP2": 12}
    assert layout.external_rows == {"TeachingX": 13}
    assert layout.cum_hours_todate_row is None
    assert layout.actual_pm_todate_row is None
    assert layout.total_row == 14
    assert layout.remaining_row == 15
    assert layout.eu_total_row == 16
    assert layout.next_row == 17


def test_plan_person_layout_without_optional_rows():
    layout = plan_person_layout(
        start_row=5,
        wp_ids=["WP1"],
        external_names=[],
        include_remaining=False,
        include_eu_total=False,
    )
    assert layout.total_row == 7
    assert layout.remaining_row is None
    assert layout.eu_total_row is None
    assert layout.cum_hours_todate_row is None
    assert layout.actual_pm_todate_row is None
    assert layout.next_row == 8


def test_plan_person_layout_with_hours_todate_summary():
    """When ``include_hours_todate_summary`` is True, two summary rows sit
    between externals and TOTAL — shifting TOTAL + EU TOTAL down by 2."""
    layout = plan_person_layout(
        start_row=10,
        wp_ids=["WP1"],
        external_names=["ProjectX"],
        include_remaining=False,
        include_eu_total=True,
        include_hours_todate_summary=True,
    )
    assert layout.header_row == 10
    assert layout.wp_rows == {"WP1": 11}
    assert layout.external_rows == {"ProjectX": 12}
    assert layout.cum_hours_todate_row == 13
    assert layout.actual_pm_todate_row == 14
    assert layout.total_row == 15
    assert layout.remaining_row is None
    assert layout.eu_total_row == 16
    assert layout.next_row == 17
