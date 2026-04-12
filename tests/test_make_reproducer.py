def test_make_reproducer_end_to_end(tmp_path):
    from budget_planner.workbook import init_workbook
    from scripts.make_reproducer import make_reproducer
    from scripts.shape_fingerprint import extract_fingerprint

    real = tmp_path / "real.xlsx"
    init_workbook(str(real))
    synth = tmp_path / "synth.xlsx"

    make_reproducer(str(real), str(synth))

    assert synth.exists()
    assert extract_fingerprint(str(synth))["config_sections"] == \
           extract_fingerprint(str(real))["config_sections"]
