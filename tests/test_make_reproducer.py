def test_make_reproducer_writes_both_artifacts_into_out_dir(tmp_path):
    from budget_planner.workbook import init_workbook
    from scripts.make_reproducer import FP_FILENAME, SYNTH_FILENAME, make_reproducer
    from scripts.shape_fingerprint import extract_fingerprint

    real = tmp_path / "real.xlsx"
    init_workbook(str(real))
    out_dir = tmp_path / "artifacts"

    fp_path, synth_path = make_reproducer(str(real), out_dir=out_dir)

    assert fp_path == out_dir / FP_FILENAME
    assert synth_path == out_dir / SYNTH_FILENAME
    assert fp_path.exists()
    assert synth_path.exists()
    assert extract_fingerprint(str(synth_path))["config_sections"] == extract_fingerprint(str(real))["config_sections"]


def test_make_reproducer_overwrites_prior_artifacts(tmp_path):
    from budget_planner.workbook import init_workbook
    from scripts.make_reproducer import FP_FILENAME, make_reproducer

    real = tmp_path / "real.xlsx"
    init_workbook(str(real))
    out_dir = tmp_path / "artifacts"
    out_dir.mkdir()
    stale = out_dir / FP_FILENAME
    stale.write_text('{"stale": true}')

    make_reproducer(str(real), out_dir=out_dir)

    assert "stale" not in stale.read_text()
