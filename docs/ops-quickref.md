# Ops Quickref (AI-facing)

Short pointers for agents. Full human runbooks live under `docs/human/` and are `.claudeignore`'d — load only when explicitly asked.

## LibreOffice is a hard dependency

`soffice` is the formula recalc engine. Without it, ~40% of the test suite skips silently. Install with `sudo apt-get install -y libreoffice-calc`.

## Common failure modes

| Symptom | Fix |
|---------|-----|
| Tests report "skipped"; `demo_recalc` fixture unavailable | Install LibreOffice. |
| `recalc.py` succeeds but `data_only=True` cells are `None` | Delete `~/.config/libreoffice/4/user/basic/Standard/`, retry. |
| `recalc.py` exits 124 (timeout) | Kill stale soffice + remove lock: `pkill -f soffice && rm -f ~/.config/libreoffice/.~lock.*`, retry. |
| LO crashes with `--headless` in a container | `unset DISPLAY`, retry. |

## Debugging against the user's real workbook

Never open the user's confidential `output/*.xlsx` file. Generate a shape-preserving synthetic via `python -m scripts.make_reproducer <path>` — outputs `tmp/synthetic.xlsx` + `tmp/fp.json` with deterministic fake values (`Person_01`, `WP_A`). The `tests/test_privacy_canary.py` canary guards the pipeline against cell-value leaks. Ask the user for the synthetic path; do not ask for the real file.

Full runbooks (pre-flight script, macro-dir init, per-failure XML probes, reproducer workflow detail) live in `docs/human/` — ignored by default.
