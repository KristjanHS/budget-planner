"""Fully automated: real.xlsx -> synthetic.xlsx (safe to share, same shape).

Usage: python scripts/make_reproducer.py <real.xlsx> <synthetic.xlsx>
       python scripts/make_reproducer.py <real.xlsx> <synthetic.xlsx> --keep-fingerprint fp.json
"""
from __future__ import annotations

import json
import sys
from pathlib import Path

from scripts.shape_fingerprint import extract_fingerprint
from scripts.build_from_shape import build_from_fingerprint


def make_reproducer(real_path: str, out_path: str, keep_fingerprint: str | None = None) -> None:
    fp = extract_fingerprint(real_path)
    if keep_fingerprint:
        Path(keep_fingerprint).write_text(json.dumps(fp, indent=2, sort_keys=True))
    build_from_fingerprint(fp, out_path)


def main(argv: list[str]) -> int:
    if len(argv) < 3:
        print(__doc__, file=sys.stderr)
        return 2
    real, out = argv[1], argv[2]
    keep = None
    if "--keep-fingerprint" in argv:
        keep = argv[argv.index("--keep-fingerprint") + 1]
    make_reproducer(real, out, keep_fingerprint=keep)
    print(f"Synthetic reproducer written to {out}. Safe to share.")
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv))
