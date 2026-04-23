"""Tests for scripts/xlsx_structural_diff.py.

Constructs both input workbooks in-memory; no fixtures.
Assertions check JSON shape, not byte-exact output (D6 + acceptance criteria).
"""

from __future__ import annotations

import json
from pathlib import Path
from typing import Any

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation

from scripts.xlsx_structural_diff import diff_workbooks, main


def _active(wb: Workbook):
    ws = wb.active
    assert ws is not None
    return ws


def _save(wb: Workbook, path: Path) -> Path:
    wb.save(path)
    return path


def _diff(tmp_path: Path, build_a, build_b) -> dict[str, Any]:
    wb_a = Workbook()
    wb_b = Workbook()
    build_a(wb_a)
    build_b(wb_b)
    a = _save(wb_a, tmp_path / "a.xlsx")
    b = _save(wb_b, tmp_path / "b.xlsx")
    return diff_workbooks(a, b)


# ---------- formulas ----------


def test_formula_changed_emits_old_and_new(tmp_path: Path):
    def build_a(wb):
        ws = wb.active
        ws.title = "S"
        ws["C5"] = "=SUM(A1:A4)"

    def build_b(wb):
        ws = wb.active
        ws.title = "S"
        ws["C5"] = "=SUM(A1:A5)"

    out = _diff(tmp_path, build_a, build_b)
    entry = out["sheets"]["S"]["cells"]["C5"]
    assert entry["kind"] == "formula"
    assert entry["old"] == "=SUM(A1:A4)"
    assert entry["new"] == "=SUM(A1:A5)"
    assert out["summary"]["formula_changes"] == 1


def test_formula_added(tmp_path: Path):
    def build_a(wb):
        wb.active.title = "S"

    def build_b(wb):
        ws = wb.active
        ws.title = "S"
        ws["D5"] = "=C5*1.2"

    out = _diff(tmp_path, build_a, build_b)
    entry = out["sheets"]["S"]["cells"]["D5"]
    assert entry["kind"] == "formula"
    assert entry["added"] == "=C5*1.2"


def test_formula_removed(tmp_path: Path):
    def build_a(wb):
        ws = wb.active
        ws.title = "S"
        ws["D5"] = "=C5*1.2"

    def build_b(wb):
        wb.active.title = "S"

    out = _diff(tmp_path, build_a, build_b)
    entry = out["sheets"]["S"]["cells"]["D5"]
    assert entry.get("removed") is True
    assert entry["old"] == "=C5*1.2"


# ---------- literals ----------


def test_literal_type_flip_value_suppressed_for_numbers(tmp_path: Path):
    """E5 numeric → string. Even though it becomes a string, E5 is outside any label
    zone (col != A, row > 3) so the value field is not emitted."""
    def build_a(wb):
        ws = wb.active
        ws.title = "S"
        ws["E5"] = 42

    def build_b(wb):
        ws = wb.active
        ws.title = "S"
        ws["E5"] = "Subtotal"

    out = _diff(tmp_path, build_a, build_b)
    entry = out["sheets"]["S"]["cells"]["E5"]
    # Per design D2: literal type flip stays under kind="literal" with type_change subfield.
    assert entry["kind"] == "literal"
    assert entry["type_change"] == {"old": "n", "new": "s"}
    # Outside label zone — no value emitted.
    assert "value" not in entry


def test_label_zone_column_a_emits_string_value(tmp_path: Path):
    """A row label (column A, row 7) must be emitted in `value`."""
    def build_a(wb):
        ws = wb.active
        ws.title = "S"
        ws["A7"] = "Old label"

    def build_b(wb):
        ws = wb.active
        ws.title = "S"
        ws["A7"] = "New label"

    out = _diff(tmp_path, build_a, build_b)
    # Same kind+type — no entry unless something else changed. So this asserts
    # via an added cell instead: change row to be added.
    # Re-do with added in B side.
    wb_a = Workbook()
    _active(wb_a).title = "S"
    wb_b = Workbook()
    ws_b = _active(wb_b)
    ws_b.title = "S"
    ws_b["A8"] = "Header label"
    a = _save(wb_a, tmp_path / "a2.xlsx")
    b = _save(wb_b, tmp_path / "b2.xlsx")
    out = diff_workbooks(a, b)
    entry = out["sheets"]["S"]["cells"]["A8"]
    assert entry["value"] == "Header label"


def test_config_b_value_never_emitted(tmp_path: Path):
    """Config!B values must be suppressed even if they look like a benign label."""
    def build_a(wb):
        ws = wb.active
        ws.title = "Config"

    def build_b(wb):
        ws = wb.active
        ws.title = "Config"
        ws["B2"] = "Round"  # short, no regex hit, but still must be suppressed

    out = _diff(tmp_path, build_a, build_b)
    entry = out["sheets"]["Config"]["cells"]["B2"]
    assert "value" not in entry
    assert entry.get("value_suppressed") is True


def test_safe_literal_regex_suppresses_decimals(tmp_path: Path):
    """A1 (column-A label zone) string that matches the sensitive regex is suppressed."""
    def build_a(wb):
        ws = wb.active
        ws.title = "S"

    def build_b(wb):
        ws = wb.active
        ws.title = "S"
        ws["A2"] = "12,500.00"  # decimal — sensitive

    out = _diff(tmp_path, build_a, build_b)
    entry = out["sheets"]["S"]["cells"]["A2"]
    assert "value" not in entry
    assert entry.get("value_suppressed") is True


# ---------- format / style ----------


def test_number_format_change_only(tmp_path: Path):
    def build_a(wb):
        ws = wb.active
        ws.title = "S"
        ws["G5"] = 1
        ws["G5"].number_format = "0.00"

    def build_b(wb):
        ws = wb.active
        ws.title = "S"
        ws["G5"] = 1
        ws["G5"].number_format = "0.00%"

    out = _diff(tmp_path, build_a, build_b)
    entry = out["sheets"]["S"]["cells"]["G5"]
    # kind preserves cell nature ("literal"); number_format change rides as a subfield.
    assert entry["kind"] == "literal"
    assert entry["number_format"]["old"] == "0.00"
    assert entry["number_format"]["new"] == "0.00%"


def test_style_change_only(tmp_path: Path):
    def build_a(wb):
        ws = wb.active
        ws.title = "S"
        ws["F5"] = 1

    def build_b(wb):
        ws = wb.active
        ws.title = "S"
        ws["F5"] = 1
        ws["F5"].font = Font(bold=True)
        ws["F5"].fill = PatternFill(start_color="FFE2EFDA", end_color="FFE2EFDA", fill_type="solid")

    out = _diff(tmp_path, build_a, build_b)
    entry = out["sheets"]["S"]["cells"]["F5"]
    assert entry["kind"] == "style"
    assert entry["bold"] == "false→true"
    assert "fill_color" in entry
    assert out["summary"]["style_changes"] == 1


# ---------- structure ----------


def test_merged_range_added(tmp_path: Path):
    def build_a(wb):
        wb.active.title = "S"

    def build_b(wb):
        ws = wb.active
        ws.title = "S"
        ws.merge_cells("A1:D1")

    out = _diff(tmp_path, build_a, build_b)
    assert "A1:D1" in out["sheets"]["S"]["structure"]["merged_added"]


def test_defined_name_added_removed_changed(tmp_path: Path):
    def build_a(wb):
        wb.active.title = "S"
        wb.defined_names["project_start"] = DefinedName("project_start", attr_text="S!$B$2")
        wb.defined_names["legacy"] = DefinedName("legacy", attr_text="S!$Z$1")

    def build_b(wb):
        wb.active.title = "S"
        wb.defined_names["project_start"] = DefinedName("project_start", attr_text="S!$B$3")
        wb.defined_names["fte_rate"] = DefinedName("fte_rate", attr_text="S!$B$4")

    out = _diff(tmp_path, build_a, build_b)
    dn = out["defined_names"]
    assert "fte_rate" in dn["added"]
    assert "legacy" in dn["removed"]
    assert dn["changed"]["project_start"]["old"] == "S!$B$2"
    assert dn["changed"]["project_start"]["new"] == "S!$B$3"
    assert out["summary"]["defined_names_changed"] == 3


def test_sheet_added_and_removed(tmp_path: Path):
    def build_a(wb):
        wb.active.title = "Keep"
        wb.create_sheet("Gone")

    def build_b(wb):
        wb.active.title = "Keep"
        wb.create_sheet("New")

    out = _diff(tmp_path, build_a, build_b)
    assert out["workbook"]["sheets_added"] == ["New"]
    assert out["workbook"]["sheets_removed"] == ["Gone"]
    assert out["summary"]["sheets_added"] == 1
    assert out["summary"]["sheets_removed"] == 1


def test_data_validation_added(tmp_path: Path):
    def build_a(wb):
        wb.active.title = "S"

    def build_b(wb):
        ws = wb.active
        ws.title = "S"
        dv = DataValidation(type="list", formula1='"a,b,c"')
        dv.add("B2:B10")
        ws.add_data_validation(dv)

    out = _diff(tmp_path, build_a, build_b)
    changes = out["sheets"]["S"]["validation_changes"]
    assert any(c.get("range") == "B2:B10" and c.get("added") for c in changes)


# ---------- shift heuristic (D6) ----------


def _populate_shift_grid(ws, row_offset: int = 0):
    """Fill rows 5-10 col B with formulas pointing at A in the same row, plus extras."""
    for r in range(5, 12):
        ws.cell(row=r + row_offset, column=2, value=f"=A{r + row_offset}+1")


def test_row_shift_hint_high_confidence(tmp_path: Path):
    """Rows 5-11 in A populated with formulas; B-side shifts everything down by 1
    and adds an extra terminal row so max_row goes up."""
    def build_a(wb):
        ws = wb.active
        ws.title = "Plan"
        for r in range(5, 12):
            ws.cell(row=r, column=2, value=f"=A{r}+1")
        ws.cell(row=15, column=1, value="anchor")  # set max_row

    def build_b(wb):
        ws = wb.active
        ws.title = "Plan"
        for r in range(5, 12):
            ws.cell(row=r, column=2, value=f"=A{r + 1}+1")
        ws.cell(row=16, column=1, value="anchor")  # max_row went up by 1

    out = _diff(tmp_path, build_a, build_b)
    hints = out["sheets"]["Plan"].get("hints", [])
    row_hints = [h for h in hints if h["kind"] == "row_shift"]
    assert len(row_hints) == 1
    assert row_hints[0]["delta"] == 1
    assert row_hints[0]["confidence"] == "high"
    assert len(row_hints[0]["sample_coords"]) >= 1


def test_row_shift_dimension_guard_blocks_hint(tmp_path: Path):
    """Same shifted formulas, but max_row unchanged → no hint."""
    def build_a(wb):
        ws = wb.active
        ws.title = "Plan"
        for r in range(5, 12):
            ws.cell(row=r, column=2, value=f"=A{r}+1")
        ws.cell(row=20, column=1, value="anchor")

    def build_b(wb):
        ws = wb.active
        ws.title = "Plan"
        for r in range(5, 12):
            ws.cell(row=r, column=2, value=f"=A{r + 1}+1")
        ws.cell(row=20, column=1, value="anchor")  # unchanged

    out = _diff(tmp_path, build_a, build_b)
    hints = out["sheets"]["Plan"].get("hints", [])
    assert all(h["kind"] != "row_shift" for h in hints)


def test_column_shift_hint_high_confidence(tmp_path: Path):
    """Symmetric to row shift, shifting refs from column A to column B and growing max_col."""
    def build_a(wb):
        ws = wb.active
        ws.title = "Plan"
        for r in range(5, 12):
            ws.cell(row=r, column=10, value=f"=A{r}+1")
        ws.cell(row=1, column=15, value="anchor")  # max_col

    def build_b(wb):
        ws = wb.active
        ws.title = "Plan"
        for r in range(5, 12):
            ws.cell(row=r, column=10, value=f"=B{r}+1")  # col delta +1
        ws.cell(row=1, column=16, value="anchor")  # max_col grew

    out = _diff(tmp_path, build_a, build_b)
    hints = out["sheets"]["Plan"].get("hints", [])
    col_hints = [h for h in hints if h["kind"] == "column_shift"]
    assert len(col_hints) == 1
    assert col_hints[0]["delta"] == 1
    assert col_hints[0]["confidence"] == "high"


def test_mixed_shift_emits_both_with_medium_confidence(tmp_path: Path):
    """Half of changed formulas carry only-row delta, half carry only-col delta."""
    def build_a(wb):
        ws = wb.active
        ws.title = "Plan"
        for r in range(5, 11):
            ws.cell(row=r, column=2, value=f"=A{r}+1")  # row-only after shift
        for r in range(5, 11):
            ws.cell(row=r, column=20, value=f"=K{r}+1")  # col-only after shift
        ws.cell(row=15, column=1, value="anchor")
        ws.cell(row=1, column=25, value="anchor2")

    def build_b(wb):
        ws = wb.active
        ws.title = "Plan"
        for r in range(5, 11):
            ws.cell(row=r, column=2, value=f"=A{r + 1}+1")  # row+1
        for r in range(5, 11):
            ws.cell(row=r, column=20, value=f"=L{r}+1")  # col+1 (K→L)
        ws.cell(row=16, column=1, value="anchor")  # max_row +1
        ws.cell(row=1, column=26, value="anchor2")  # max_col +1

    out = _diff(tmp_path, build_a, build_b)
    hints = out["sheets"]["Plan"].get("hints", [])
    kinds = {h["kind"]: h for h in hints}
    assert "row_shift" in kinds
    assert "column_shift" in kinds
    assert kinds["row_shift"]["confidence"] == "medium"
    assert kinds["column_shift"]["confidence"] == "medium"
    assert "note" in kinds["row_shift"]


def test_subthreshold_no_hint(tmp_path: Path):
    """Only 3 shifted formulas — below noise floor of 5."""
    def build_a(wb):
        ws = wb.active
        ws.title = "Plan"
        for r in range(5, 8):
            ws.cell(row=r, column=2, value=f"=A{r}+1")

    def build_b(wb):
        ws = wb.active
        ws.title = "Plan"
        for r in range(5, 8):
            ws.cell(row=r, column=2, value=f"=A{r + 1}+1")

    out = _diff(tmp_path, build_a, build_b)
    assert "hints" not in out["sheets"]["Plan"] or out["sheets"]["Plan"]["hints"] == []


# ---------- acceptance / sanity ----------


def test_identical_workbooks_have_zero_summary(tmp_path: Path):
    def build(wb):
        ws = wb.active
        ws.title = "S"
        ws["A1"] = "Header"
        ws["B2"] = "=A1"

    out = _diff(tmp_path, build, build)
    s = out["summary"]
    assert s["formula_changes"] == 0
    assert s["cells_added"] == 0
    assert s["cells_removed"] == 0
    assert s["style_changes"] == 0
    assert s["hints_emitted"] == 0
    assert all(not sd["cells"] for sd in out["sheets"].values())


def test_main_writes_diff_json(tmp_path: Path):
    wb_a = Workbook()
    ws_a = _active(wb_a)
    ws_a.title = "S"
    ws_a["C5"] = "=SUM(A1:A4)"
    wb_b = Workbook()
    ws_b = _active(wb_b)
    ws_b.title = "S"
    ws_b["C5"] = "=SUM(A1:A5)"
    a = _save(wb_a, tmp_path / "a.xlsx")
    b = _save(wb_b, tmp_path / "b.xlsx")
    out_path = tmp_path / "diff.json"
    rc = main([str(a), str(b), "--out", str(out_path)])
    assert rc == 0
    parsed = json.loads(out_path.read_text())
    assert parsed["meta"]["schema_version"] == 1
    assert parsed["sheets"]["S"]["cells"]["C5"]["new"] == "=SUM(A1:A5)"


def test_formula_and_style_changing_together_emits_both(tmp_path: Path):
    """A cell whose formula AND style both changed must surface BOTH facets —
    early-return ordering used to drop the style change silently."""
    def build_a(wb):
        ws = _active(wb)
        ws.title = "S"
        ws["C5"] = "=SUM(A1:A4)"

    def build_b(wb):
        ws = _active(wb)
        ws.title = "S"
        ws["C5"] = "=SUM(A1:A5)"
        ws["C5"].font = Font(bold=True)

    out = _diff(tmp_path, build_a, build_b)
    entry = out["sheets"]["S"]["cells"]["C5"]
    assert entry["kind"] == "formula"
    assert entry["new"] == "=SUM(A1:A5)"
    # Style must travel under a nested key, not be discarded.
    assert "style" in entry
    assert entry["style"]["bold"] == "false→true"


def test_no_decimal_value_leaks_when_only_data_changed(tmp_path: Path):
    """If the only change is a numeric literal value flip, the diff body must not
    contain the actual numbers."""
    def build_a(wb):
        ws = wb.active
        ws.title = "Money"
        ws["B5"] = 1234.56

    def build_b(wb):
        ws = wb.active
        ws.title = "Money"
        ws["B5"] = 9876.54

    out = _diff(tmp_path, build_a, build_b)
    serialized = json.dumps(out)
    # The actual numeric values must not appear anywhere in the JSON.
    assert "1234.56" not in serialized
    assert "9876.54" not in serialized
