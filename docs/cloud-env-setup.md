# Cloud Environment Setup

## Required System Dependencies

### LibreOffice (CRITICAL)

LibreOffice is the project's formula calculation engine — not optional.
Without it, ~40% of the test suite is untestable (tests skip, not fail).

```bash
sudo apt-get update && sudo apt-get install -y libreoffice-calc
soffice --version                    # verify installed
soffice --headless --terminate_after_init  # verify headless works
```

### LibreOffice Macro Directory

The recalc script installs a StarBasic macro. If the directory is missing,
let LO create it: `soffice --headless --terminate_after_init`

## Pre-Flight Verification Script

Run before starting work on the project:

```bash
#!/bin/bash
set -e
echo "=== Environment Pre-Flight Check ==="

# 1. Python + package
echo -n "Python: "; python3 --version
echo -n "Package installed: "
python3 -c "import budget_planner; print('OK')" 2>/dev/null || echo "FAIL - run: uv pip install -e '.[dev]'"

# 2. LibreOffice
echo -n "LibreOffice: "
if command -v soffice &>/dev/null; then
    soffice --version 2>/dev/null | head -1
else
    echo "FAIL - CRITICAL: install libreoffice-calc"
    exit 1
fi

# 3. Macro directory
MACRO_DIR="$HOME/.config/libreoffice/4/user/basic/Standard"
echo -n "Macro dir: "
if [ -d "$MACRO_DIR" ]; then
    echo "OK ($MACRO_DIR)"
else
    echo "Missing — initializing..."
    soffice --headless --terminate_after_init 2>/dev/null
    mkdir -p "$MACRO_DIR"
    echo "Created $MACRO_DIR"
fi

# 4. End-to-end recalc probe
echo -n "Recalc probe: "
python3 -c "
from budget_planner.workbook import init_workbook
import tempfile, os, subprocess, json
path = os.path.join(tempfile.mkdtemp(), 'probe.xlsx')
init_workbook(path)
r = subprocess.run(['python3', 'scripts/recalc.py', path], capture_output=True, text=True)
if r.returncode != 0:
    print('FAIL: recalc returned ' + str(r.returncode))
else:
    data = json.loads(r.stdout)
    print('OK — ' + str(data.get('total_formulas', 0)) + ' formulas computed')
" 2>&1

echo "=== Pre-flight complete ==="
```

## Common Failure Modes

| # | Symptom | Cause | Fix |
|---|---------|-------|-----|
| 1 | `demo_recalc` fixture skips; tests show "skipped" | LO not installed | `sudo apt-get install -y libreoffice-calc` |
| 2 | `recalc.py` succeeds but all formula cells are `None` with `data_only=True` | Macro didn't execute (`calculateAll()` never ran) | Delete macro dir and retry: `rm -rf ~/.config/libreoffice/4/user/basic/Standard/ && python scripts/recalc.py file.xlsx` |
| 3 | `recalc.py` returns exit code 124 (timeout) | Stale LO process holds a lock | `pkill -f soffice && rm -f ~/.config/libreoffice/.~lock.* && retry` |
| 4 | LO crashes on startup even with `--headless` | Container display config interferes | `unset DISPLAY` then retry |

**Diagnosis for failure #2** — verify whether LO wrote cached values:

```bash
python -m budget_planner init -o /tmp/test.xlsx
python scripts/recalc.py /tmp/test.xlsx
python3 -c "
import zipfile
with zipfile.ZipFile('/tmp/test.xlsx') as z:
    xml = z.read('xl/worksheets/sheet2.xml').decode()
    count = xml.count('<v>')
    print(f'<v> elements in Plan sheet: {count}')
    if count == 0: print('PROBLEM: LO did not write cached values')
"
```
