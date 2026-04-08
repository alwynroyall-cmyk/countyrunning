# WRRL League AI — Python Dependencies

This document lists all Python package dependencies required to run the WRRL League AI application.

## Required Python Packages

- **pandas** >= 1.5.0  
  Data analysis and manipulation library. Used for reading, processing, and aggregating race and club data.

- **openpyxl** >= 3.0.10  
  Excel file reader/writer. Used for reading and writing `.xlsx` files (race results, club lists, outputs).

- **xlrd** >= 2.0.1  
  Legacy Excel support (`.xls` files).

- **Pillow** >= 10.0.0  
  Python Imaging Library. Used for image processing (e.g., generating timeline PNGs). Font paths are resolved per-platform (Windows, macOS, Linux).

- **python-docx** >= 1.1.0  
  Microsoft Word document library. Used for generating `.docx` race reports and league summaries.

- **docx2pdf** >= 0.1.8  
  Converts `.docx` files to `.pdf` format for report distribution. Requires Microsoft Word on Windows; PDF output is skipped on macOS/Linux unless a compatible converter is available.

- **playwright** >= 1.50.0  
  Browser automation used by Sporthive import tooling. After installing, run `playwright install chromium` once.

## Installation

All dependencies can be installed with:

```sh
pip install -r requirements.txt
```

## Platform Support

As of v6.1.0, the application supports Windows, macOS, and Linux. The scoring pipeline, GUI, and all file-handling code are cross-platform. PDF export requires Microsoft Word (Windows) or a separately configured converter (macOS/Linux).

## Notes
- All dependencies are listed in `requirements.txt` in the project root.
- The application uses standard library modules (`os`, `sys`, `threading`, `datetime`, `pathlib`, etc.) which do not require separate installation.
- If you add new features that require extra packages, update both `requirements.txt` and this document.
