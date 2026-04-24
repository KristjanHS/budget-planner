"""
xlsx structural diff — emit logic + structure of two workbooks while suppressing values.

Consumer: an AI agent that wrote v1 of budget_planner/**, asked to replicate user-driven
manual edits in code without ever seeing budget numbers, rates, hours, or names.

Default output is `--mode=logic` (drops pure-style cell entries + style sub-dicts
from mixed entries) with compact JSON, to keep the result small enough for the
consumer's context window. Use `--mode=full --pretty` for human-readable output
that includes every styling delta.

Schema versions
---------------
- v1 — original shape (xlsx structural diff design).
- v2 — Stage 1 noise-suppression additions (additive; no renames / removals):
  - `noise_filtered_summary` top-level dict (always emitted, zero counts as ``0``)
    counting cosmetic deltas collapsed by canonicalization predicates
    (``number_format_normalized``, ``tab_color_alpha_normalized``).
  - ``--keep-noise`` CLI: retain such records and tag each with ``noise_class``
    (``"format"`` or ``"color"``).
  - N4 fix: ``added``/``removed`` branches now preserve ``value`` for
    string-literal cells the same way ``kind_change`` does (subject to the
    existing D2 label-zone + SENSITIVE_LITERAL_RE safety filters).

Designs: docs/plans/archived/2026-04-23-xlsx-structural-diff-design.md (v1)
         docs/plans/archived/2026-04-23-diff-noise-suppression-stage-1-design.md (v2)
"""

from __future__ import annotations

import argparse
import datetime as _dt
import hashlib
import json
import re
import sys
from collections import Counter
from pathlib import Path
from typing import Any

from openpyxl import load_workbook
from openpyxl.cell.cell import Cell
from openpyxl.utils import column_index_from_string, get_column_letter
from openpyxl.workbook.defined_name import DefinedName

SCHEMA_VERSION = 2

# D2 — safe-literal sensitive-value regex.
SENSITIVE_LITERAL_RE = re.compile(
    r"[\d][\d,.\s]*[€%]"          # currency/percent-ish
    r"|@"                          # email markers
    r"|\b\d{4}-\d{2}-\d{2}\b"     # ISO dates
    r"|\b\d+[,.]\d+\b"            # decimals
    r"|\b\d{4,}\b"                # 4+ digit numbers (IDs, years, round euros)
)

# D6 — A1 reference parsing.
A1_REF_RE = re.compile(r"(\$?)([A-Z]{1,3})(\$?)(\d+)")

# D2 — label-zone defaults.
LABEL_COL_A = 1
LABEL_FIRST_ROWS = 3

# D7 — Config sheet hard-suppress column B values.
CONFIG_SHEET_NAME = "Config"
CONFIG_SUPPRESS_COL = 2

# D6 — heuristic thresholds.
SHIFT_MIN_FORMULAS = 5
SHIFT_THRESHOLD = 0.80


# ---------- normalization (D3) ----------


def _build_merge_anchor_map(ws: Any) -> dict[tuple[int, int], tuple[int, int]]:
    """Map every (row, col) inside a merged range to its top-left anchor."""
    anchors: dict[tuple[int, int], tuple[int, int]] = {}
    for rng in ws.merged_cells.ranges:
        anchor = (rng.min_row, rng.min_col)
        for r in range(rng.min_row, rng.max_row + 1):
            for c in range(rng.min_col, rng.max_col + 1):
                anchors[(r, c)] = anchor
    return anchors


# Shared/array-formula anchor handling (D3): openpyxl drops most shared-formula
# refs on read — only the anchor cell carries the formula string; member cells
# come back with `value is None`. The merge-anchor map above already filters
# merged-range members; for shared formulas we rely on openpyxl's own behavior
# (members read as empty → naturally excluded).


# ---------- style + cell signature ----------


def _style_fingerprint(cell: Cell) -> dict[str, str]:
    """Extract style fields explicitly. Missing values normalize to "none"."""
    font = cell.font
    fill = cell.fill
    border = cell.border
    align = cell.alignment

    def _color_rgb(color: Any) -> str:
        if color is None:
            return "none"
        rgb = getattr(color, "rgb", None)
        if rgb is None or not isinstance(rgb, str):
            return "none"
        return rgb

    fill_color = "none"
    if getattr(fill, "patternType", None) == "solid":
        fill_color = _color_rgb(getattr(fill, "fgColor", None))

    def _border_style(side: Any) -> str:
        if side is None:
            return "none"
        s = getattr(side, "style", None)
        return s if s else "none"

    return {
        "bold": str(bool(getattr(font, "bold", False))).lower(),
        "italic": str(bool(getattr(font, "italic", False))).lower(),
        "font_color": _color_rgb(getattr(font, "color", None)),
        "fill_color": fill_color,
        "border_left": _border_style(getattr(border, "left", None)),
        "border_right": _border_style(getattr(border, "right", None)),
        "border_top": _border_style(getattr(border, "top", None)),
        "border_bottom": _border_style(getattr(border, "bottom", None)),
        "halign": getattr(align, "horizontal", None) or "none",
        "valign": getattr(align, "vertical", None) or "none",
        "wrap_text": str(bool(getattr(align, "wrap_text", False))).lower(),
    }


def _cell_kind(cell: Cell) -> str:
    val = cell.value
    if val is None:
        return "empty"
    if isinstance(val, str) and val.startswith("="):
        return "formula"
    return "literal"


def _cell_signature(cell: Cell) -> dict[str, Any]:
    kind = _cell_kind(cell)
    sig: dict[str, Any] = {
        "kind": kind,
        "number_format": cell.number_format or "General",
        "style": _style_fingerprint(cell),
    }
    if kind == "formula":
        sig["formula"] = cell.value
    elif kind == "literal":
        sig["data_type"] = cell.data_type
    return sig


# ---------- safe-literal emit (D2 + D7) ----------


def _is_label_zone(sheet_name: str, row: int, col: int) -> bool:
    if col == LABEL_COL_A:
        return True
    if row == 1:
        return True
    if row <= LABEL_FIRST_ROWS:
        return True
    return False


def _emit_literal_value(
    sheet_name: str, row: int, col: int, cell: Cell
) -> tuple[str | None, bool]:
    """Return (value_to_emit, value_was_suppressed_due_to_regex).

    value_to_emit is None when not in label zone or non-string.
    """
    val = cell.value
    if not isinstance(val, str):
        return None, False
    if cell.data_type != "s":
        return None, False
    if len(val) > 40:
        return None, False
    # D7 — Config!B always suppressed regardless of D2.
    if sheet_name == CONFIG_SHEET_NAME and col == CONFIG_SUPPRESS_COL:
        return None, True  # pretend regex hit so consumer knows value existed
    if not _is_label_zone(sheet_name, row, col):
        return None, False
    if SENSITIVE_LITERAL_RE.search(val):
        return None, True
    return val, False


# ---------- noise canonicalization (Stage 1) ----------


def _canon_numfmt(s: Any) -> Any:
    """Canonical form of a number-format string for noise-suppression comparison.

    Collapses case + Excel backslash escapes so cosmetic differences compare equal:
    - ``mmm\\ yyyy`` ≡ ``MMM YYYY``
    - ``\\€#,##0``   ≡ ``€#,##0``

    Preserves meaningful differences: ``"0.00"`` ≠ ``"0.0000"``;
    ``"mmm yyyy"`` ≠ ``"mm yyyy"``. Non-string inputs pass through.
    """
    if not isinstance(s, str):
        return s
    return s.upper().replace("\\", "")


def _canon_argb(s: Any) -> Any:
    """Canonical form of an 8-char ARGB color for noise-suppression comparison.

    Excel ignores the alpha byte for tab colors: ``FF`` and ``00`` both read as
    fully opaque. Collapse both onto a sentinel prefix so they compare equal:
    - ``"FF0000FF"`` ≡ ``"000000FF"`` → both → ``"__0000FF"``.

    Non-opaque alphas (e.g. ``80``) pass through uppercased so partial
    transparency, if ever intentional, still registers as a real delta.
    Non-string inputs pass through unchanged.
    """
    if isinstance(s, str) and len(s) == 8 and s[:2].upper() in ("FF", "00"):
        return "__" + s[2:].upper()
    if isinstance(s, str):
        return s.upper()
    return s


def _compare_numfmt(a: Any, b: Any, counter: dict[str, int]) -> tuple[bool, str | None]:
    """Return ``(equal, noise_class)`` for a number-format comparison.

    ``equal`` is True on raw match OR canonical match. When canonical match
    drops a raw mismatch, the counter's ``number_format_normalized`` key is
    bumped and ``noise_class == "format"`` is returned for optional tagging.
    """
    if a == b:
        return True, None
    if _canon_numfmt(a) == _canon_numfmt(b):
        counter["number_format_normalized"] += 1
        return True, "format"
    return False, None


def _compare_argb(a: Any, b: Any, counter: dict[str, int]) -> tuple[bool, str | None]:
    """Return ``(equal, noise_class)`` for an ARGB color comparison.

    Same shape as :func:`_compare_numfmt` — when the only delta is the alpha
    byte collapsing under :func:`_canon_argb`, bumps ``tab_color_alpha_normalized``
    and returns ``noise_class == "color"``.
    """
    if a == b:
        return True, None
    if _canon_argb(a) == _canon_argb(b):
        counter["tab_color_alpha_normalized"] += 1
        return True, "color"
    return False, None


def _new_noise_counter() -> dict[str, int]:
    """Fresh zeroed ``noise_filtered_summary`` accumulator. Always emitted."""
    return {"number_format_normalized": 0, "tab_color_alpha_normalized": 0}


# ---------- per-sheet build ----------


def _collect_cells(ws: Any) -> dict[tuple[int, int], Cell]:
    """Collect every non-empty cell, mapping merged-range members to their anchor."""
    anchor_map = _build_merge_anchor_map(ws)
    cells: dict[tuple[int, int], Cell] = {}
    for row in ws.iter_rows():
        for cell in row:
            coord = (cell.row, cell.column)
            anchored = anchor_map.get(coord, coord)
            if anchored != coord:
                # member of merged range — skip; anchor cell will be visited.
                continue
            if cell.value is None and cell.number_format in ("General", None):
                # truly empty + default format — skip only if no styling either.
                style = _style_fingerprint(cell)
                blank_style = {
                    "bold": "false",
                    "italic": "false",
                    "font_color": "none",
                    "fill_color": "none",
                    "border_left": "none",
                    "border_right": "none",
                    "border_top": "none",
                    "border_bottom": "none",
                    "halign": "none",
                    "valign": "none",
                    "wrap_text": "false",
                }
                if style == blank_style:
                    continue
            cells[coord] = cell
    return cells


def _coord_str(coord: tuple[int, int]) -> str:
    return f"{get_column_letter(coord[1])}{coord[0]}"


def _coord_sort_key(coord_str: str) -> tuple[int, str]:
    """Sort A1 coords by (row, col) — natural reading order."""
    m = re.match(r"([A-Z]+)(\d+)", coord_str)
    if not m:
        return (0, coord_str)
    col_letters, row_str = m.group(1), m.group(2)
    return (int(row_str), col_letters)


# ---------- defined names (D5) ----------


def _defined_names_map(wb: Any) -> dict[str, str]:
    out: dict[str, str] = {}
    for name in wb.defined_names:
        dn = wb.defined_names[name]
        if isinstance(dn, DefinedName):
            out[name] = dn.attr_text or ""
    return out


def _diff_defined_names(
    a_names: dict[str, str], b_names: dict[str, str]
) -> dict[str, Any]:
    a_keys = set(a_names)
    b_keys = set(b_names)
    added = {k: b_names[k] for k in sorted(b_keys - a_keys)}
    removed = {k: a_names[k] for k in sorted(a_keys - b_keys)}
    changed: dict[str, dict[str, str]] = {}
    for k in sorted(a_keys & b_keys):
        if a_names[k] != b_names[k]:
            changed[k] = {"old": a_names[k], "new": b_names[k]}
    return {"added": added, "removed": removed, "changed": changed}


# ---------- data validation diff ----------


def _validation_signatures(ws: Any) -> list[dict[str, Any]]:
    sigs: list[dict[str, Any]] = []
    dvs = getattr(ws, "data_validations", None)
    if dvs is None:
        return sigs
    for dv in getattr(dvs, "dataValidation", []) or []:
        ranges = sorted(str(r) for r in (dv.sqref.ranges if dv.sqref else []))
        formula = (dv.formula1 or "") if hasattr(dv, "formula1") else ""
        list_hash: str | None = None
        if (dv.type or "") == "list" and formula:
            list_hash = "sha256:" + hashlib.sha256(formula.encode("utf-8")).hexdigest()
        sigs.append(
            {
                "range": ",".join(ranges),
                "type": dv.type or "",
                "formula": formula,
                "list_hash": list_hash,
            }
        )
    return sorted(sigs, key=lambda s: s["range"])


def _diff_validations(a_sigs: list[dict[str, Any]], b_sigs: list[dict[str, Any]]) -> list[dict[str, Any]]:
    changes: list[dict[str, Any]] = []
    a_by_range = {s["range"]: s for s in a_sigs}
    b_by_range = {s["range"]: s for s in b_sigs}
    all_ranges = sorted(set(a_by_range) | set(b_by_range))
    for rng in all_ranges:
        a = a_by_range.get(rng)
        b = b_by_range.get(rng)
        if a is None and b is not None:
            changes.append({"range": rng, "added": True, "type": b["type"], "list_hash_new": b["list_hash"]})
        elif b is None and a is not None:
            changes.append({"range": rng, "removed": True, "type": a["type"], "list_hash_old": a["list_hash"]})
        elif a is not None and b is not None and a != b:
            entry: dict[str, Any] = {"range": rng}
            if a["type"] != b["type"]:
                entry["type_change"] = {"old": a["type"], "new": b["type"]}
            if a["formula"] != b["formula"]:
                entry["formula_change"] = True  # never emit formula contents — may carry hardcoded list
            if a["list_hash"] != b["list_hash"]:
                entry["list_hash_old"] = a["list_hash"]
                entry["list_hash_new"] = b["list_hash"]
            changes.append(entry)
    return changes


# ---------- shift heuristic (D6) ----------


def _parse_refs(formula: str) -> list[tuple[int, int]]:
    """Return list of (row, col_index) tuples in positional order."""
    refs: list[tuple[int, int]] = []
    for m in A1_REF_RE.finditer(formula):
        col_letters = m.group(2)
        row = int(m.group(4))
        refs.append((row, column_index_from_string(col_letters)))
    return refs


def _shift_hints(
    changed_formulas: dict[tuple[int, int], tuple[str, str]],
    a_dim: dict[str, int],
    b_dim: dict[str, int],
) -> list[dict[str, Any]]:
    """Detect row/col shift heuristics per D6. Returns 0–2 hints."""
    deltas: list[tuple[int, int, tuple[int, int]]] = []  # (row_delta, col_delta, source_coord)
    for coord, (old_f, new_f) in changed_formulas.items():
        old_refs = _parse_refs(old_f)
        new_refs = _parse_refs(new_f)
        if len(old_refs) != len(new_refs) or not old_refs:
            continue
        for (orow, ocol), (nrow, ncol) in zip(old_refs, new_refs):
            deltas.append((nrow - orow, ncol - ocol, coord))

    contributing_coords: set[tuple[int, int]] = {d[2] for d in deltas}
    if len(contributing_coords) < SHIFT_MIN_FORMULAS:
        return []

    # Per-direction pools: a row shift dominates the col_delta==0 sub-pool;
    # a column shift dominates the row_delta==0 sub-pool. (D6: "The 80% threshold
    # applies per direction" — checking 80% of the whole multiset would make a
    # mixed row+col shift mathematically impossible.)
    row_pool = [(rd, cd) for rd, cd, _ in deltas if cd == 0 and rd != 0]
    col_pool = [(rd, cd) for rd, cd, _ in deltas if rd == 0 and cd != 0]

    row_candidates: Counter[int] = Counter(rd for rd, _ in row_pool)
    col_candidates: Counter[int] = Counter(cd for _, cd in col_pool)

    row_hint: dict[str, Any] | None = None
    col_hint: dict[str, Any] | None = None

    if row_candidates and row_pool:
        best_row, best_row_count = row_candidates.most_common(1)[0]
        if best_row_count / len(row_pool) >= SHIFT_THRESHOLD and (
            (best_row > 0 and b_dim["max_row"] > a_dim["max_row"])
            or (best_row < 0 and b_dim["max_row"] < a_dim["max_row"])
        ):
            row_hint = _build_hint("row_shift", best_row, deltas, "row", "high")

    if col_candidates and col_pool:
        best_col, best_col_count = col_candidates.most_common(1)[0]
        if best_col_count / len(col_pool) >= SHIFT_THRESHOLD and (
            (best_col > 0 and b_dim["max_col"] > a_dim["max_col"])
            or (best_col < 0 and b_dim["max_col"] < a_dim["max_col"])
        ):
            col_hint = _build_hint("column_shift", best_col, deltas, "col", "high")

    # Mixed: both directions present at threshold per direction.
    if row_hint is not None and col_hint is not None:
        row_hint["confidence"] = "medium"
        col_hint["confidence"] = "medium"
        row_hint["note"] = "mixed row+col shift — verify"
        col_hint["note"] = "mixed row+col shift — verify"

    out: list[dict[str, Any]] = []
    if row_hint is not None:
        out.append(row_hint)
    if col_hint is not None:
        out.append(col_hint)
    return out


def _build_hint(
    kind: str,
    delta: int,
    deltas: list[tuple[int, int, tuple[int, int]]],
    direction: str,
    confidence: str,
) -> dict[str, Any]:
    """Build a row_shift/column_shift hint from delta multiset."""
    matching_coords: list[tuple[int, int]] = []
    for rd, cd, coord in deltas:
        if direction == "row" and cd == 0 and rd == delta:
            matching_coords.append(coord)
        elif direction == "col" and rd == 0 and cd == delta:
            matching_coords.append(coord)
    # inferred_start (heuristic-only per D6): plan spec uses smallest old-side
    # *referenced* row/col minus delta, but we only have the formula-cell coord
    # in the deltas tuple — using coord[0]/coord[1] as a proxy. For self- or
    # near-self-referencing formulas (the common case) this produces the right
    # answer; for formulas that reference rows far from their own coord it
    # may be off. Documented as heuristic in the design.
    inferred_start: int | None = None
    if direction == "row":
        rows_referenced = []
        for rd, cd, coord in deltas:
            if cd == 0 and rd == delta:
                rows_referenced.append(coord[0])
        if rows_referenced:
            inferred_start = min(rows_referenced) - delta
    else:
        cols_referenced = []
        for rd, cd, coord in deltas:
            if rd == 0 and cd == delta:
                cols_referenced.append(coord[1])
        if cols_referenced:
            inferred_start = min(cols_referenced) - delta

    sample_coords = sorted({_coord_str(c) for c in matching_coords}, key=_coord_sort_key)[:5]
    payload: dict[str, Any] = {
        "kind": kind,
        "delta": delta,
        "confidence": confidence,
        "sample_coords": sample_coords,
    }
    if inferred_start is not None:
        if kind == "row_shift":
            payload["inferred_start_row"] = inferred_start
        else:
            payload["inferred_start_col"] = inferred_start
    return payload


# ---------- per-sheet diff ----------


def _diff_sheet(
    sheet_name: str,
    ws_a: Any | None,
    ws_b: Any | None,
    counter: dict[str, int],
    keep_noise: bool,
) -> dict[str, Any]:
    """Build the per-sheet diff dict."""
    structure = _diff_structure(ws_a, ws_b, counter, keep_noise)
    cells_diff: dict[str, Any] = {}
    changed_formulas: dict[tuple[int, int], tuple[str, str]] = {}

    a_cells = _collect_cells(ws_a) if ws_a is not None else {}
    b_cells = _collect_cells(ws_b) if ws_b is not None else {}

    a_coords = set(a_cells)
    b_coords = set(b_cells)

    # Added: in B only.
    for coord in sorted(b_coords - a_coords):
        cell = b_cells[coord]
        coord_s = _coord_str(coord)
        kind = _cell_kind(cell)
        entry: dict[str, Any] = {"kind": kind}
        if kind == "formula":
            entry["added"] = cell.value
        elif kind == "literal":
            entry["added"] = True
            entry["data_type"] = cell.data_type
            value, suppressed = _emit_literal_value(sheet_name, coord[0], coord[1], cell)
            if value is not None:
                entry["value"] = value
            elif suppressed:
                entry["value_suppressed"] = True
        else:
            entry["added"] = True
        cells_diff[coord_s] = entry

    # Removed: in A only.
    for coord in sorted(a_coords - b_coords):
        cell = a_cells[coord]
        coord_s = _coord_str(coord)
        kind = _cell_kind(cell)
        entry = {"kind": kind, "removed": True}
        if kind == "formula":
            entry["old"] = cell.value
        elif kind == "literal":
            entry["data_type"] = cell.data_type
            # N4 fix: mirror the `added` branch — serialize string literal value
            # through the label-zone + regex safety filter.
            value, suppressed = _emit_literal_value(
                sheet_name, coord[0], coord[1], cell
            )
            if value is not None:
                entry["value"] = value
            elif suppressed:
                entry["value_suppressed"] = True
        cells_diff[coord_s] = entry

    # Changed: in both, signatures differ.
    for coord in sorted(a_coords & b_coords):
        ca = a_cells[coord]
        cb = b_cells[coord]
        sa = _cell_signature(ca)
        sb = _cell_signature(cb)
        if sa == sb:
            continue
        coord_s = _coord_str(coord)
        entry = _build_change_entry(
            sheet_name, coord, ca, cb, sa, sb, counter, keep_noise
        )
        if entry:
            cells_diff[coord_s] = entry
        if sa["kind"] == "formula" and sb["kind"] == "formula" and sa.get("formula") != sb.get("formula"):
            changed_formulas[coord] = (sa["formula"], sb["formula"])

    # Sort cells_diff by natural A1 order.
    sorted_cells = dict(sorted(cells_diff.items(), key=lambda kv: _coord_sort_key(kv[0])))

    sheet_out: dict[str, Any] = {
        "structure": structure,
        "cells": sorted_cells,
    }

    a_validations = _validation_signatures(ws_a) if ws_a is not None else []
    b_validations = _validation_signatures(ws_b) if ws_b is not None else []
    val_changes = _diff_validations(a_validations, b_validations)
    if val_changes:
        sheet_out["validation_changes"] = val_changes

    if ws_a is not None and ws_b is not None:
        a_dim = {
            "max_row": int(ws_a.max_row or 0),
            "max_col": int(ws_a.max_column or 0),
        }
        b_dim = {
            "max_row": int(ws_b.max_row or 0),
            "max_col": int(ws_b.max_column or 0),
        }
        hints = _shift_hints(changed_formulas, a_dim, b_dim)
        if hints:
            sheet_out["hints"] = hints

    return sheet_out


def _build_change_entry(
    sheet_name: str,
    coord: tuple[int, int],
    ca: Cell,
    cb: Cell,
    sa: dict[str, Any],
    sb: dict[str, Any],
    counter: dict[str, int],
    keep_noise: bool,
) -> dict[str, Any]:
    """Produce a single cells diff entry for a coord present on both sides."""
    # If kind itself changed, that subsumes everything.
    if sa["kind"] != sb["kind"]:
        entry: dict[str, Any] = {
            "kind": "kind_change",
            "old_kind": sa["kind"],
            "new_kind": sb["kind"],
        }
        if sb["kind"] == "formula":
            entry["new_formula"] = sb.get("formula")
        if sa["kind"] == "formula":
            entry["old_formula"] = sa.get("formula")
        return entry

    kind = sa["kind"]
    entry: dict[str, Any] = {"kind": kind}
    has_payload = False

    # Formula change (logic-bearing — primary kind).
    if kind == "formula" and sa.get("formula") != sb.get("formula"):
        entry["old"] = sa["formula"]
        entry["new"] = sb["formula"]
        has_payload = True

    # Literal type change (data_type flip).
    if kind == "literal" and sa.get("data_type") != sb.get("data_type"):
        entry["type_change"] = {"old": sa.get("data_type"), "new": sb.get("data_type")}
        value, suppressed = _emit_literal_value(sheet_name, coord[0], coord[1], cb)
        if value is not None:
            entry["value"] = value
        elif suppressed:
            entry["value_suppressed"] = True
        has_payload = True

    # number_format change — combine into the same entry rather than overwriting kind.
    # Canonicalization predicate absorbs Excel-cosmetic differences (case, escape
    # backslashes); counter bumps when a raw mismatch collapses to canonical match.
    # keep_noise retains the record and stamps `noise_class`.
    if sa["number_format"] != sb["number_format"]:
        equal, noise_class = _compare_numfmt(
            sa["number_format"], sb["number_format"], counter
        )
        if not equal or (keep_noise and noise_class):
            entry["number_format"] = {
                "old": sa["number_format"],
                "new": sb["number_format"],
            }
            if keep_noise and noise_class:
                entry["noise_class"] = noise_class
            has_payload = True

    # style change — emit changed fields under nested `style` block so the parent
    # `kind` keeps its semantic meaning (formula/literal/empty) when other facets
    # also changed. Standalone style edits get kind="style" promoted.
    if sa["style"] != sb["style"]:
        style_diff = {
            field: f"{sa['style'][field]}→{sb['style'][field]}"
            for field in sorted(sa["style"])
            if sa["style"][field] != sb["style"][field]
        }
        if has_payload:
            entry["style"] = style_diff
        else:
            entry = {"kind": "style", **style_diff}
            has_payload = True

    if not has_payload:
        return {}
    return entry


def _diff_structure(
    ws_a: Any | None,
    ws_b: Any | None,
    counter: dict[str, int],
    keep_noise: bool,
) -> dict[str, Any]:
    out: dict[str, Any] = {}

    a_dim = {"max_row": int(ws_a.max_row or 0), "max_col": int(ws_a.max_column or 0)} if ws_a else {"max_row": 0, "max_col": 0}
    b_dim = {"max_row": int(ws_b.max_row or 0), "max_col": int(ws_b.max_column or 0)} if ws_b else {"max_row": 0, "max_col": 0}
    if a_dim != b_dim:
        out["dimensions"] = {"old": a_dim, "new": b_dim}

    a_merged = sorted(str(r) for r in (ws_a.merged_cells.ranges if ws_a else []))
    b_merged = sorted(str(r) for r in (ws_b.merged_cells.ranges if ws_b else []))
    added_m = sorted(set(b_merged) - set(a_merged))
    removed_m = sorted(set(a_merged) - set(b_merged))
    if added_m:
        out["merged_added"] = added_m
    if removed_m:
        out["merged_removed"] = removed_m

    a_freeze = ws_a.freeze_panes if ws_a else None
    b_freeze = ws_b.freeze_panes if ws_b else None
    if a_freeze != b_freeze:
        out["frozen_panes"] = {"old": a_freeze, "new": b_freeze}

    a_tab = ws_a.sheet_properties.tabColor.rgb if ws_a and ws_a.sheet_properties.tabColor else None
    b_tab = ws_b.sheet_properties.tabColor.rgb if ws_b and ws_b.sheet_properties.tabColor else None
    if a_tab != b_tab:
        # Canonicalization absorbs ARGB alpha-byte noise (FF vs 00 both opaque).
        equal, noise_class = _compare_argb(a_tab, b_tab, counter)
        if not equal or (keep_noise and noise_class):
            tab_entry: dict[str, Any] = {"old": a_tab, "new": b_tab}
            if keep_noise and noise_class:
                tab_entry["noise_class"] = noise_class
            out["tab_color"] = tab_entry

    a_hidden = (ws_a.sheet_state if ws_a else None) == "hidden"
    b_hidden = (ws_b.sheet_state if ws_b else None) == "hidden"
    if a_hidden != b_hidden:
        out["hidden"] = {"old": a_hidden, "new": b_hidden}

    return out


# ---------- top-level diff ----------


def diff_workbooks(
    path_a: Path, path_b: Path, keep_noise: bool = False
) -> dict[str, Any]:
    wb_a = load_workbook(filename=str(path_a), data_only=False)
    wb_b = load_workbook(filename=str(path_b), data_only=False)

    a_sheets = set(wb_a.sheetnames)
    b_sheets = set(wb_b.sheetnames)

    sheets_added = sorted(b_sheets - a_sheets)
    sheets_removed = sorted(a_sheets - b_sheets)

    a_names = _defined_names_map(wb_a)
    b_names = _defined_names_map(wb_b)
    defined_names = _diff_defined_names(a_names, b_names)

    workbook_block: dict[str, Any] = {
        "sheets_added": sheets_added,
        "sheets_removed": sheets_removed,
        "protection_changed": bool(wb_a.security != wb_b.security),
    }

    noise_counter = _new_noise_counter()

    sheet_diffs: dict[str, dict[str, Any]] = {}
    for name in sorted(a_sheets | b_sheets):
        ws_a = wb_a[name] if name in a_sheets else None
        ws_b = wb_b[name] if name in b_sheets else None
        sheet_diffs[name] = _diff_sheet(
            name, ws_a, ws_b, noise_counter, keep_noise
        )

    formula_changes = 0
    cells_added = 0
    cells_removed = 0
    style_changes = 0
    hints_emitted = 0
    for sd in sheet_diffs.values():
        for entry in sd.get("cells", {}).values():
            # formula_changes covers any formula activity: edit, add, or removal —
            # the consumer's "skip body if zero" triage relies on this catching
            # added/removed formulas too, not just edits.
            if entry.get("kind") == "formula":
                formula_changes += 1
            if entry.get("kind") == "style" or "style" in entry:
                style_changes += 1
            if entry.get("added") is not None and entry.get("added") is not False:
                cells_added += 1
            if entry.get("removed"):
                cells_removed += 1
        hints_emitted += len(sd.get("hints", []))

    summary = {
        "sheets_added": len(sheets_added),
        "sheets_removed": len(sheets_removed),
        "defined_names_changed": (
            len(defined_names["added"])
            + len(defined_names["removed"])
            + len(defined_names["changed"])
        ),
        "formula_changes": formula_changes,
        "cells_added": cells_added,
        "cells_removed": cells_removed,
        "style_changes": style_changes,
        "hints_emitted": hints_emitted,
    }

    return {
        "meta": {
            "file_a": str(path_a),
            "file_b": str(path_b),
            "generated_at": _dt.datetime.now(_dt.timezone.utc).isoformat(timespec="seconds").replace("+00:00", "Z"),
            "schema_version": SCHEMA_VERSION,
        },
        "summary": summary,
        "noise_filtered_summary": noise_counter,
        "defined_names": defined_names,
        "workbook": workbook_block,
        "sheets": sheet_diffs,
    }


def _filter_logic_mode(diff: dict[str, Any]) -> dict[str, Any]:
    """Drop pure-style cell entries and style sub-dicts from mixed entries.

    Keeps the full `summary` block so the consumer sees suppressed counts
    (e.g. `style_changes: 500`) even when bodies are hidden.

    Edge case: `kind == "empty"` entries with `added`/`removed` are also dropped.
    These represent styled-empty cells, usually decoration. The one case this
    misses is a formula cleared to a styled placeholder (e.g. greyed blank
    replacing a deleted formula); consumers replicating logic edits should
    cross-check formula removals against any unexpected styled-empty appearances
    in `--mode=full` if such a pattern is suspected.
    """
    for sd in diff["sheets"].values():
        kept: dict[str, Any] = {}
        for coord, entry in sd.get("cells", {}).items():
            kind = entry.get("kind")
            if kind == "style":
                continue
            if kind == "empty":
                # Styled empty cells added/removed — decoration, not logic.
                continue
            if "style" in entry:
                entry = {k: v for k, v in entry.items() if k != "style"}
            kept[coord] = entry
        sd["cells"] = kept
    return diff


def _pick_oldest_newest(dir_path: Path) -> tuple[Path, Path]:
    if not dir_path.is_dir():
        raise SystemExit(f"--dir {dir_path} is not a directory")
    candidates = [
        p for p in dir_path.iterdir()
        if p.is_file() and p.suffix.lower() == ".xlsx" and not p.name.startswith("~$")
    ]
    if len(candidates) < 2:
        raise SystemExit(
            f"--dir {dir_path}: need at least 2 .xlsx files (found {len(candidates)})"
        )
    candidates.sort(key=lambda p: p.stat().st_mtime)
    return candidates[0], candidates[-1]


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description="Emit a confidentiality-safe structural diff between two .xlsx files."
    )
    parser.add_argument("file_a", type=Path, nargs="?")
    parser.add_argument("file_b", type=Path, nargs="?")
    parser.add_argument(
        "--dir",
        dest="dir_path",
        type=Path,
        help="Directory containing ≥2 .xlsx files; diffs oldest vs newest by mtime (excludes ~$ lock files).",
    )
    parser.add_argument("--out", type=Path, default=Path("diff.json"))
    parser.add_argument(
        "--mode",
        choices=["logic", "full"],
        default="logic",
        help="logic (default): drop pure-style cell entries and style sub-dicts — "
        "keeps consumer context small. full: emit every styling delta.",
    )
    parser.add_argument(
        "--pretty",
        action="store_true",
        help="Pretty-print JSON with indent=2. Default is compact (one-line) for "
        "smaller file size when an AI agent is the consumer.",
    )
    parser.add_argument(
        "--keep-noise",
        action="store_true",
        help="Retain records that would otherwise be dropped by noise "
        "canonicalization (number-format case/escape, ARGB alpha); tag each "
        "surviving record with `noise_class`.",
    )
    args = parser.parse_args(argv)

    if args.dir_path is not None:
        if args.file_a is not None or args.file_b is not None:
            parser.error("--dir is mutually exclusive with positional file arguments")
        file_a, file_b = _pick_oldest_newest(args.dir_path)
        print(f"Oldest : {file_a}")
        print(f"Newest : {file_b}")
    else:
        if args.file_a is None or args.file_b is None:
            parser.error("provide two positional file arguments or --dir")
        file_a, file_b = args.file_a, args.file_b

    if args.out.exists():
        stamp = _dt.datetime.now().strftime("%Y%m%d_%H%M%S")
        archive = args.out.with_name(f"{args.out.stem}.{stamp}{args.out.suffix}")
        args.out.rename(archive)
        print(f"Archived previous result to {archive}")

    diff = diff_workbooks(file_a, file_b, keep_noise=args.keep_noise)
    diff["meta"]["mode"] = args.mode
    if args.mode == "logic":
        diff = _filter_logic_mode(diff)
    if args.pretty:
        payload = json.dumps(diff, sort_keys=True, indent=2, ensure_ascii=False)
    else:
        payload = json.dumps(diff, sort_keys=True, separators=(",", ":"), ensure_ascii=False)
    args.out.write_text(payload, encoding="utf-8")
    return 0


if __name__ == "__main__":
    sys.exit(main())
