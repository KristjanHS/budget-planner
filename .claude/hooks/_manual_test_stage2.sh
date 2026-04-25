#!/usr/bin/env bash
# Manual test cases for docs-only-plan-gate.py and external-process-test-gate.py.
# Run from repo root: bash .claude/hooks/_manual_test_stage2.sh
#
# Each test prints PASS/FAIL based on expected exit code.

set -u
cd "$(dirname "$0")/../.."
PROJ="$(pwd)"
HOOK_PLAN=".claude/hooks/docs-only-plan-gate.py"
HOOK_TEST=".claude/hooks/external-process-test-gate.py"
STAMP="$HOME/.claude/sessions/$(python3 -c "import hashlib; print(hashlib.md5(b'$PROJ').hexdigest()[:8])")/pytest-pass-stamp"

pass=0; fail=0
check() {
    # check <name> <expected_exit> <actual_exit>
    if [ "$2" = "$3" ]; then
        echo "PASS: $1 (exit=$3)"
        pass=$((pass+1))
    else
        echo "FAIL: $1 (expected=$2, actual=$3)"
        fail=$((fail+1))
    fi
}

run_pre() {
    # run_pre <hook> <event> <tool> <payload_json> -> exit code
    echo "$4" \
      | CLAUDE_PROJECT_DIR="$PROJ" python3 "$1" >/dev/null 2>&1
    echo $?
}

echo "=== docs-only-plan-gate (stamp absent) ==="
rm -f "$STAMP"

# Block path
ec=$(run_pre "$HOOK_PLAN" PreToolUse Bash '{
  "hook_event_name":"PreToolUse","tool_name":"Bash",
  "tool_input":{"command":"git mv docs/plans/2026-04-25-foo.md docs/plans/archived/2026-04-25-foo.md"}}')
check "block: git mv plan to archived/ (no stamp)" 2 "$ec"

ec=$(run_pre "$HOOK_PLAN" PreToolUse Bash '{
  "hook_event_name":"PreToolUse","tool_name":"Bash",
  "tool_input":{"command":"git mv \"docs/plans/X.md\" \"docs/plans/archived/X.md\""}}')
check "block: quoted git mv to archived/ (no stamp)" 2 "$ec"

# Allow paths (no git mv pattern)
ec=$(run_pre "$HOOK_PLAN" PreToolUse Bash '{
  "hook_event_name":"PreToolUse","tool_name":"Bash",
  "tool_input":{"command":"git status"}}')
check "allow: unrelated git command" 0 "$ec"

ec=$(run_pre "$HOOK_PLAN" PreToolUse Bash '{
  "hook_event_name":"PreToolUse","tool_name":"Bash",
  "tool_input":{"command":"git log docs/plans/archived/X.md"}}')
check "allow: git log mentioning archived/" 0 "$ec"

ec=$(run_pre "$HOOK_PLAN" PreToolUse Bash '{
  "hook_event_name":"PreToolUse","tool_name":"Bash",
  "tool_input":{"command":"git mv docs/plans/X.md docs/plans/Y.md"}}')
check "allow: git mv plan -> sibling plan (no archived/)" 0 "$ec"

# Override sentinel
ec=$(run_pre "$HOOK_PLAN" PreToolUse Bash '{
  "hook_event_name":"PreToolUse","tool_name":"Bash",
  "tool_input":{"command":"git mv docs/plans/X.md docs/plans/archived/X.md  # <!-- docs-only-plan-gate-override: pytest just verified manually -->"}}')
check "allow: override sentinel present" 0 "$ec"

# Stamp fresh path
mkdir -p "$(dirname "$STAMP")" && date -u +%FT%TZ > "$STAMP"
ec=$(run_pre "$HOOK_PLAN" PreToolUse Bash '{
  "hook_event_name":"PreToolUse","tool_name":"Bash",
  "tool_input":{"command":"git mv docs/plans/X.md docs/plans/archived/X.md"}}')
check "allow: stamp fresh" 0 "$ec"

# Stamp stale path
touch -d "2 hours ago" "$STAMP"
ec=$(run_pre "$HOOK_PLAN" PreToolUse Bash '{
  "hook_event_name":"PreToolUse","tool_name":"Bash",
  "tool_input":{"command":"git mv docs/plans/X.md docs/plans/archived/X.md"}}')
check "block: stamp stale (>30min)" 2 "$ec"

echo
echo "=== docs-only-plan-gate (PostToolUse stamp updater) ==="
rm -f "$STAMP"

# Successful pytest -> writes stamp
ec=$(run_pre "$HOOK_PLAN" PostToolUse Bash '{
  "hook_event_name":"PostToolUse","tool_name":"Bash",
  "tool_input":{"command":"pytest tests/ -q"},
  "tool_response":{"exit_code":0,"stdout":"606 passed in 12.3s","stderr":""}}')
check "stamp: pytest exit 0" 0 "$ec"
[ -f "$STAMP" ] && echo "  → stamp present (good)" || { echo "  → stamp MISSING (bad)"; fail=$((fail+1)); }

rm -f "$STAMP"
ec=$(run_pre "$HOOK_PLAN" PostToolUse Bash '{
  "hook_event_name":"PostToolUse","tool_name":"Bash",
  "tool_input":{"command":"pytest tests/ -q"},
  "tool_response":{"exit_code":1,"stdout":"1 failed, 605 passed","stderr":""}}')
check "no-stamp: pytest exit 1" 0 "$ec"
[ ! -f "$STAMP" ] && echo "  → stamp absent (good)" || { echo "  → stamp PRESENT (bad)"; fail=$((fail+1)); }

ec=$(run_pre "$HOOK_PLAN" PostToolUse Bash '{
  "hook_event_name":"PostToolUse","tool_name":"Bash",
  "tool_input":{"command":"pytest --help"},
  "tool_response":{"exit_code":0}}')
check "no-stamp: pytest --help (no tests/)" 0 "$ec"
[ ! -f "$STAMP" ] && echo "  → stamp absent (good)" || { echo "  → stamp PRESENT (bad)"; fail=$((fail+1)); }

# Stdout-fallback path (no exit_code field)
rm -f "$STAMP"
ec=$(run_pre "$HOOK_PLAN" PostToolUse Bash '{
  "hook_event_name":"PostToolUse","tool_name":"Bash",
  "tool_input":{"command":".venv/bin/pytest tests/test_config_sheet.py -q"},
  "tool_response":{"stdout":"5 passed in 0.8s","stderr":""}}')
check "stamp: stdout fallback (passed, no exit_code)" 0 "$ec"
[ -f "$STAMP" ] && echo "  → stamp present (good)" || { echo "  → stamp MISSING (bad)"; fail=$((fail+1)); }

echo
echo "=== external-process-test-gate (Write/Edit) ==="

# Block: docker via subprocess
ec=$(run_pre "$HOOK_TEST" PreToolUse Write "{
  \"hook_event_name\":\"PreToolUse\",\"tool_name\":\"Write\",
  \"tool_input\":{\"file_path\":\"$PROJ/tests/test_docker_e2e.py\",
  \"content\":\"import subprocess\\n\\ndef test_x():\\n    subprocess.run([\\\"docker\\\", \\\"run\\\", \\\"alpine\\\"])\\n\"}}")
check "block: subprocess docker call in tests/" 2 "$ec"

# Block: playwright import
ec=$(run_pre "$HOOK_TEST" PreToolUse Write "{
  \"hook_event_name\":\"PreToolUse\",\"tool_name\":\"Write\",
  \"tool_input\":{\"file_path\":\"$PROJ/tests/test_browser.py\",
  \"content\":\"from playwright.sync_api import sync_playwright\\ndef test_x(): pass\\n\"}}")
check "block: playwright import in tests/" 2 "$ec"

# Block: requests.get
ec=$(run_pre "$HOOK_TEST" PreToolUse Write "{
  \"hook_event_name\":\"PreToolUse\",\"tool_name\":\"Write\",
  \"tool_input\":{\"file_path\":\"$PROJ/tests/test_api.py\",
  \"content\":\"import requests\\ndef test_x():\\n    requests.get('http://example.com')\\n\"}}")
check "block: requests.get in tests/" 2 "$ec"

# Block: Edit adding kubectl subprocess
ec=$(run_pre "$HOOK_TEST" PreToolUse Edit "{
  \"hook_event_name\":\"PreToolUse\",\"tool_name\":\"Edit\",
  \"tool_input\":{\"file_path\":\"$PROJ/tests/test_existing.py\",
  \"old_string\":\"def test_old():\\n    pass\",
  \"new_string\":\"def test_new():\\n    subprocess.run(['kubectl', 'get', 'pods'])\"}}")
check "block: Edit adds kubectl subprocess" 2 "$ec"

# Allow: under tests/manual/
ec=$(run_pre "$HOOK_TEST" PreToolUse Write "{
  \"hook_event_name\":\"PreToolUse\",\"tool_name\":\"Write\",
  \"tool_input\":{\"file_path\":\"$PROJ/tests/manual/test_docker_e2e.py\",
  \"content\":\"import subprocess\\nsubprocess.run([\\\"docker\\\", \\\"run\\\"])\\n\"}}")
check "allow: external-process under tests/manual/" 0 "$ec"

# Allow: legitimate subprocess.run (not docker/kubectl/etc) — recalc.py
ec=$(run_pre "$HOOK_TEST" PreToolUse Write "{
  \"hook_event_name\":\"PreToolUse\",\"tool_name\":\"Write\",
  \"tool_input\":{\"file_path\":\"$PROJ/tests/test_recalc.py\",
  \"content\":\"import subprocess\\nsubprocess.run(['python', 'scripts/recalc.py', 'output/x.xlsx'])\\n\"}}")
check "allow: subprocess.run for project recalc (not external CLI)" 0 "$ec"

# Allow: conftest.py edit
ec=$(run_pre "$HOOK_TEST" PreToolUse Edit "{
  \"hook_event_name\":\"PreToolUse\",\"tool_name\":\"Edit\",
  \"tool_input\":{\"file_path\":\"$PROJ/tests/conftest.py\",
  \"old_string\":\"x\",\"new_string\":\"y\\nsubprocess.run(['docker','run'])\"}}")
check "allow: conftest.py exempt (project-specific recalc fixtures)" 0 "$ec"

# Allow: override sentinel
ec=$(run_pre "$HOOK_TEST" PreToolUse Write "{
  \"hook_event_name\":\"PreToolUse\",\"tool_name\":\"Write\",
  \"tool_input\":{\"file_path\":\"$PROJ/tests/test_x.py\",
  \"content\":\"# <!-- external-process-test-gate-override: cross-checking against live API in this one-off -->\\nimport requests\\nrequests.get('http://x')\\n\"}}")
check "allow: override sentinel in content" 0 "$ec"

# Allow: outside tests/
ec=$(run_pre "$HOOK_TEST" PreToolUse Write "{
  \"hook_event_name\":\"PreToolUse\",\"tool_name\":\"Write\",
  \"tool_input\":{\"file_path\":\"$PROJ/budget_planner/x.py\",
  \"content\":\"import subprocess\\nsubprocess.run(['docker','run'])\\n\"}}")
check "allow: not under tests/" 0 "$ec"

# Allow: empty new_string
ec=$(run_pre "$HOOK_TEST" PreToolUse Edit "{
  \"hook_event_name\":\"PreToolUse\",\"tool_name\":\"Edit\",
  \"tool_input\":{\"file_path\":\"$PROJ/tests/test_x.py\",\"old_string\":\"import requests\",\"new_string\":\"\"}}")
check "allow: Edit removes external-process code (empty new_string)" 0 "$ec"

# Allow: tests/__init__.py exempt
ec=$(run_pre "$HOOK_TEST" PreToolUse Write "{
  \"hook_event_name\":\"PreToolUse\",\"tool_name\":\"Write\",
  \"tool_input\":{\"file_path\":\"$PROJ/tests/__init__.py\",\"content\":\"\"}}")
check "allow: tests/__init__.py exempt" 0 "$ec"

echo
echo "=== Latency check ==="
t0=$(date +%s%N)
for i in $(seq 1 10); do
    echo '{"hook_event_name":"PreToolUse","tool_name":"Bash","tool_input":{"command":"git status"}}' \
      | CLAUDE_PROJECT_DIR="$PROJ" python3 "$HOOK_PLAN" >/dev/null 2>&1
done
t1=$(date +%s%N)
ms_per=$(( (t1 - t0) / 10000000 ))
echo "docs-only-plan-gate: ~${ms_per}ms per invocation (target <50ms)"

t0=$(date +%s%N)
for i in $(seq 1 10); do
    echo "{\"hook_event_name\":\"PreToolUse\",\"tool_name\":\"Write\",\"tool_input\":{\"file_path\":\"$PROJ/budget_planner/x.py\",\"content\":\"x\"}}" \
      | CLAUDE_PROJECT_DIR="$PROJ" python3 "$HOOK_TEST" >/dev/null 2>&1
done
t1=$(date +%s%N)
ms_per=$(( (t1 - t0) / 10000000 ))
echo "external-process-test-gate: ~${ms_per}ms per invocation (target <50ms)"

# Cleanup stamp from tests
rm -f "$STAMP"

echo
echo "=== Result: $pass passed, $fail failed ==="
exit "$fail"
