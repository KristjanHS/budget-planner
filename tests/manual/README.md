# tests/manual/ — manual-tier tests

Tests that drive **external systems** (docker, kubectl, headless browsers, live
HTTP services) live here. They are **not** collected by the default
`pytest tests/ -q` run — invoke explicitly:

```sh
pytest tests/manual/ -q
```

## Why a separate tier

External-process tests are slow, environmentally sensitive, and prone to
false-greens when the harness silently degrades (e.g. a docker daemon stops,
a remote service times out and the test passes because the mock fallback
fires). Keeping them out of the default suite preserves the signal of the
fast suite.

The `.claude/hooks/external-process-test-gate.py` hook blocks new test files
under `tests/` (outside this directory) that contain docker/kubectl/playwright/
selenium/`requests.<verb>(` patterns — those must move here, or carry an
`<!-- external-process-test-gate-override: <reason ≥10> -->` sentinel
explaining the exemption.

## Source rule

`~/.claude/rules/testing.md` — *"External-process tests (docker, kubectl,
headless browser) go in manual tier from day one."*
