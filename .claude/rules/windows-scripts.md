---
paths:
  - "*.bat"
  - "*.cmd"
  - "**/*.bat"
  - "**/*.cmd"
---
# Windows Batch Wrappers

Do not spawn `powershell.exe` or `pwsh.exe` from a project `.bat`. Keep the wrapper pure cmd; push any non-trivial logic (filesystem enumeration with filters, sorts, regex, timestamps) into the Python script the `.bat` activates `.venv` and calls.

- **Why:** user runs these wrappers in `cmd.exe`, not PowerShell. Cmd's two-pass backtick parser inside `for /f \`powershell ...\`` leaks PS stderr into the captured stdout — the calling script then treats the error text as a value and plows on with garbage paths. `^|` escaping inside the quoted `-Command` string is fragile and has failed repeatedly here.
- **How to apply:** triggers on any edit to `*.bat` / `*.cmd` in this repo. If the task seems to need PS (`Get-ChildItem` with filters, `Get-Date`, `Sort-Object`), instead: (a) add a flag or subcommand to the Python target that does the work, (b) call python once from the `.bat` with args assembled from plain `set` / `%VAR%` expansion. Wrapper stays a ≤30-line launcher: existence checks, `.venv\Scripts\activate.bat`, one python call, errorlevel check, `pause`.
