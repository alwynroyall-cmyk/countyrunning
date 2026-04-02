# CI / PR Policy (Phase 1)

This repository now uses a minimal Windows-only CI workflow:

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
