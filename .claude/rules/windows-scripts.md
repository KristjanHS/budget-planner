---
paths:
  - "*.bat"
  - "*.cmd"
  - "**/*.bat"
  - "**/*.cmd"
---
# Windows Batch Wrappers

Do not spawn `powershell.exe` / `pwsh.exe` from a project `.bat` — `cmd.exe`'s `for /f` backtick parser leaks PS stderr into captured stdout. Keep wrappers pure cmd (`.venv\Scripts\activate.bat` + one python call + errorlevel + `pause`); push enumeration/regex/timestamp logic into the Python target as a flag or subcommand.
