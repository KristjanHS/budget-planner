"""Fully automated: real.xlsx -> tmp/synthetic.xlsx + tmp/fp.json.

Always writes to the repo's ``tmp/`` directory, overwriting prior runs, so
debugging sessions can reliably reference the same canonical paths.

Usage:
    python -m scripts.make_reproducer <real.xlsx>
    python -m scripts.make_reproducer <real.xlsx> --out-dir some/other/dir

The real workbook can live anywhere (absolute path recommended). The
synthetic artifacts always land inside the chosen out-dir.
"""

from __future__ import annotations

import json
import sys
from pathlib import Path

from scripts.shape_fingerprint import extract_fingerprint
from scripts.build_from_shape import build_from_fingerprint

DEFAULT_OUT_DIR = Path("tmp")
FP_FILENAME = "fp.json"
SYNTH_FILENAME = "synthetic.xlsx"


def make_reproducer(
    real_path: str,
    out_dir: str | Path = DEFAULT_OUT_DIR,
) -> tuple[Path, Path]:
    """Extract fingerprint + build synthetic, both into out_dir. Returns (fp_path, synth_path)."""
    out_dir = Path(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    fp_path = out_dir / FP_FILENAME
    synth_path = out_dir / SYNTH_FILENAME

    fp = extract_fingerprint(real_path)
    fp_path.write_text(json.dumps(fp, indent=2, sort_keys=True), encoding="utf-8")
    build_from_fingerprint(fp, str(synth_path))
    return fp_path, synth_path


def main(argv: list[str]) -> int:
    if len(argv) < 2 or argv[1] in {"-h", "--help"}:
        print(__doc__, file=sys.stderr)
        return 2
    real = argv[1]
    out_dir: str | Path = DEFAULT_OUT_DIR
    if "--out-dir" in argv:
        out_dir = argv[argv.index("--out-dir") + 1]
    fp_path, synth_path = make_reproducer(real, out_dir=out_dir)
    print(f"Fingerprint: {fp_path}")
    print(f"Synthetic:   {synth_path}")
    print("Both are safe to share — value-free by construction.")
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv))
