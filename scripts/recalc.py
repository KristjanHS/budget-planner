"""
Excel Formula Recalculation Script
Recalculates all formulas in an Excel file using LibreOffice
"""

import json
import os
import platform
import subprocess
import sys
import zipfile
from pathlib import Path

from office.soffice import get_soffice_env

from openpyxl import load_workbook

MACRO_DIR_MACOS = "~/Library/Application Support/LibreOffice/4/user/basic/Standard"
MACRO_DIR_LINUX = "~/.config/libreoffice/4/user/basic/Standard"
MACRO_FILENAME = "Module1.xba"

RECALCULATE_MACRO = """<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE script:module PUBLIC "-//OpenOffice.org//DTD OfficeDocument 1.0//EN" "module.dtd">
<script:module xmlns:script="http://openoffice.org/2000/script" script:name="Module1" script:language="StarBasic">
    Sub RecalculateAndSave()
      ThisComponent.calculateAll()

      ' Auto-fit column widths on all sheets
      Dim oSheets As Object
      oSheets = ThisComponent.getSheets()
      Dim i As Long
      For i = 0 To oSheets.getCount() - 1
        Dim oSheet As Object
        oSheet = oSheets.getByIndex(i)
        Dim oCols As Object
        oCols = oSheet.getColumns()
        Dim j As Long
        For j = 0 To oCols.getCount() - 1
          Dim oCol As Object
          oCol = oCols.getByIndex(j)
          oCol.OptimalWidth = True
          ' Add padding (150 = 1.5mm in 1/100mm units)
          oCol.Width = oCol.Width + 150
        Next j
      Next i

      ThisComponent.store()
      ThisComponent.close(True)
    End Sub
</script:module>"""


def has_gtimeout():
    try:
        subprocess.run(
            ["gtimeout", "--version"], capture_output=True, timeout=1, check=False
        )
        return True
    except (FileNotFoundError, subprocess.TimeoutExpired):
        return False


def setup_libreoffice_macro():
    macro_dir = os.path.expanduser(
        MACRO_DIR_MACOS if platform.system() == "Darwin" else MACRO_DIR_LINUX
    )
    macro_file = os.path.join(macro_dir, MACRO_FILENAME)

    if (
        os.path.exists(macro_file)
        and "RecalculateAndSave" in Path(macro_file).read_text()
    ):
        return True

    if not os.path.exists(macro_dir):
        subprocess.run(
            ["soffice", "--headless", "--terminate_after_init"],
            capture_output=True,
            timeout=10,
            env=get_soffice_env(),
        )
        os.makedirs(macro_dir, exist_ok=True)

    try:
        Path(macro_file).write_text(RECALCULATE_MACRO)
        return True
    except Exception:
        return False


def recalc(filename, timeout=30):
    if not Path(filename).exists():
        return {"error": f"File {filename} does not exist"}

    abs_path = str(Path(filename).absolute())

    if not setup_libreoffice_macro():
        return {"error": "Failed to setup LibreOffice macro"}

    cmd = [
        "soffice",
        "--headless",
        "--norestore",
        "vnd.sun.star.script:Standard.Module1.RecalculateAndSave?language=Basic&location=application",
        abs_path,
    ]

    if platform.system() == "Linux":
        cmd = ["timeout", str(timeout)] + cmd
    elif platform.system() == "Darwin" and has_gtimeout():
        cmd = ["gtimeout", str(timeout)] + cmd

    # Preserve chart XML — LibreOffice strips series data references on save
    chart_backups = {}
    with zipfile.ZipFile(abs_path, "r") as zin:
        for name in zin.namelist():
            if name.startswith("xl/charts/") and name.endswith(".xml"):
                chart_backups[name] = zin.read(name)

    result = subprocess.run(cmd, capture_output=True, text=True, env=get_soffice_env())

    if result.returncode != 0 and result.returncode != 124:
        error_msg = result.stderr or "Unknown error during recalculation"
        if "Module1" in error_msg or "RecalculateAndSave" not in error_msg:
            return {"error": "LibreOffice macro not configured properly"}
        return {"error": error_msg}

    # Restore chart XML that LibreOffice mangled
    if chart_backups:
        tmp_path = abs_path + ".tmp"
        with zipfile.ZipFile(abs_path, "r") as zin, \
             zipfile.ZipFile(tmp_path, "w") as zout:
            for item in zin.infolist():
                if item.filename in chart_backups:
                    zout.writestr(item, chart_backups[item.filename])
                else:
                    zout.writestr(item, zin.read(item.filename))
        os.replace(tmp_path, abs_path)

    try:
        wb = load_workbook(filename, data_only=True)

        # Hard errors: indicate formula bugs. #N/A is tracked separately
        # because NA() is an intentional idiom (LineChart series clipping,
        # lookup sentinels) — Excel itself produces it by design, not as
        # a failure mode.
        excel_errors = [
            "#VALUE!",
            "#DIV/0!",
            "#REF!",
            "#NAME?",
            "#NULL!",
            "#NUM!",
        ]
        error_details = {err: [] for err in excel_errors}
        na_locations: list[str] = []
        total_errors = 0

        for sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
            for row in ws.iter_rows():
                for cell in row:
                    if cell.value is not None and isinstance(cell.value, str):
                        if "#N/A" in cell.value:
                            na_locations.append(f"{sheet_name}!{cell.coordinate}")
                            continue
                        for err in excel_errors:
                            if err in cell.value:
                                location = f"{sheet_name}!{cell.coordinate}"
                                error_details[err].append(location)
                                total_errors += 1
                                break

        wb.close()

        result = {
            "status": "success" if total_errors == 0 else "errors_found",
            "total_errors": total_errors,
            "error_summary": {},
            "na_count": len(na_locations),
        }

        for err_type, locations in error_details.items():
            if locations:
                result["error_summary"][err_type] = {
                    "count": len(locations),
                    "locations": locations[:20],
                }
        if na_locations:
            result["na_summary"] = {
                "count": len(na_locations),
                "locations": na_locations[:20],
            }

        wb_formulas = load_workbook(filename, data_only=False)
        formula_count = 0
        for sheet_name in wb_formulas.sheetnames:
            ws = wb_formulas[sheet_name]
            for row in ws.iter_rows():
                for cell in row:
                    if (
                        cell.value
                        and isinstance(cell.value, str)
                        and cell.value.startswith("=")
                    ):
                        formula_count += 1
        wb_formulas.close()

        result["total_formulas"] = formula_count

        return result

    except Exception as e:
        return {"error": str(e)}


def main():
    if len(sys.argv) < 2:
        print("Usage: python recalc.py <excel_file> [timeout_seconds]")
        print("\nRecalculates all formulas in an Excel file using LibreOffice")
        print("\nReturns JSON with error details:")
        print("  - status: 'success' or 'errors_found'")
        print("  - total_errors: Total number of Excel errors found")
        print("  - total_formulas: Number of formulas in the file")
        print("  - error_summary: Breakdown by error type with locations")
        print("    - #VALUE!, #DIV/0!, #REF!, #NAME?, #NULL!, #NUM!")
        print("  - na_count: #N/A cells (intentional NA() — e.g., chart clipping)")
        sys.exit(1)

    filename = sys.argv[1]
    timeout = int(sys.argv[2]) if len(sys.argv) > 2 else 30

    result = recalc(filename, timeout)
    print(json.dumps(result, indent=2))


if __name__ == "__main__":
    main()
