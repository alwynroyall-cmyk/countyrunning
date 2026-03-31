# WRRL League Scorer — Python Dependencies

This document lists all Python package dependencies required to run the Wiltshire Road Running League (WRRL) League Scorer application.

## Required Python Packages

- **pandas** >= 1.5.0  
  Data analysis and manipulation library. Used for reading, processing, and aggregating race and club data.

- **openpyxl** >= 3.0.10  
  Excel file reader/writer. Used for reading and writing `.xlsx` files (race results, club lists, outputs).

- **Pillow** >= 10.0.0  
  Python Imaging Library. Used for image processing (e.g., generating timeline PNGs).

- **python-docx** >= 1.1.0  
  Microsoft Word document library. Used for generating `.docx` race reports and league summaries.

- **docx2pdf** >= 0.1.8  
  Converts `.docx` files to `.pdf` format for report distribution.

## Installation

All dependencies can be installed with:

```sh
pip install -r requirements.txt
```

## Notes
- All dependencies are listed in `requirements.txt` in the project root.
- The application may use additional standard library modules (e.g., `os`, `datetime`, `csv`), but these do not require separate installation.
- If you add new features that require extra packages, update both `requirements.txt` and this document.
