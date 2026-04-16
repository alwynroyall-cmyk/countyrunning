# Installation Guide

This document explains how to install and update the WRRL Admin Suite as a Python package.

## Prerequisites

- Windows machine with Python 3.11 installed.
- A Python virtual environment is strongly recommended.
- The package includes only Python dependencies. External runtime requirements are not bundled.

### External requirements

- `docx2pdf` requires Microsoft Word if PDF conversion is needed.
- `playwright` may require browser binaries to be installed separately by running `playwright install` after installation.

## Install from a built wheel

If you already have a built distribution artifact, install it with:

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
python -m pip install dist\league_scorer-8.3.0-py3-none-any.whl
```

After installation, launch the GUI with:

```powershell
wrrl-admin
```

or directly with Python:

```powershell
python -m league_scorer.graphical.qt
```

## Build the package locally

If you want to build a release package from the repository source:

1. Create and activate a virtual environment.
2. Install build tooling:

   ```powershell
   python -m pip install --upgrade pip
   python -m pip install build
   ```

3. Run the build command from the repository root:

   ```powershell
   python -m build
   ```

This creates:

- `dist\league_scorer-8.3.0-py3-none-any.whl`
- `dist\league_scorer-8.3.0.tar.gz`

## Update to a new release

1. Build the next release wheel with the updated version.
2. Install or upgrade on the target machine:

   ```powershell
   python -m pip install --upgrade dist\league_scorer-<new-version>-py3-none-any.whl
   ```

## Verify installation

Run a quick import check:

```powershell
python -c "import league_scorer; print(league_scorer.__version__)"
```

Expected output:

```text
8.3.0
```

## Notes

- `requirements.txt` contains the Python dependency policy for the application.
- The package is designed to include only Python code and dependencies; it does not bundle machine-specific external software.
- If you need to install browser support for Playwright, run:

```powershell
playwright install
```

- If Word-based PDF output is required, install Microsoft Word on the target computer.
