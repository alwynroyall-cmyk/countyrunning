# CI / PR Policy (Phase 1)

This repository uses a minimal CI workflow:

- Workflow: `Windows CI`
- Job: `test-and-stage-checks`
- Trigger: pull requests and pushes

## Required PR Gate

Set branch protection on the target branch (for example `main`) so this check must pass before merge:

- Required status check: `Windows CI / test-and-stage-checks`

## Why This Is Required

The check verifies:

1. Unit tests for core scoring and validation modules
2. Staged chain checks (1-4) in safe mode for CI environments

This provides a lightweight but real quality gate while keeping tooling minimal.

## Portability Note

The codebase supports Windows, macOS, and Linux as of v6.1.0. The CI workflow currently runs on Windows. If a Linux or macOS runner is added in future, no code changes are required.
