from budget_planner.person_section import plan_person_layout


def test_plan_person_layout_assigns_sequential_rows():
    layout = plan_person_layout(
        start_row=10,
        wp_ids=["WP1", "WP2"],
        external_names=["TeachingX"],
        include_remaining=False,
    )
    assert layout.header_row == 10
    assert layout.wp_rows == {"WP1": 11, "WP2": 12}
    assert layout.external_rows == {"TeachingX": 13}
    assert layout.total_row == 14
    assert layout.remaining_row is None
    assert layout.next_row == 15
