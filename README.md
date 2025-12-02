# Matrix Processor

A web application to process Excel/CSV files and create intersection matrices.

## 🚀 Quick Start

**Double-click `START.bat`** to launch the application.

**Requirements:** Python 3.6+ (usually pre-installed on most systems)

The app will automatically install required packages (`pandas`, `openpyxl`) on first run.

---

## Features

1. **Multi-file Upload** - Load Excel (.xlsx, .xls) or CSV files
2. **Tab Selection** - Choose which sheets to process
3. **Column Selection** - Select Y axis, X axis, and optional secondary X axis
4. **Matrix Configuration** - Merge matrices or keep them independent
5. **Matrix Computation** - Creates intersection matrices (marks intersections with 1)
6. **Excel Export** - Exports results to a new Excel file

---

## How to Use

1. Double-click `START.bat`
2. Upload your Excel/CSV files
3. Select the tabs you want to process
4. Choose columns for Y axis, X axis (and optionally secondary X axis)
5. Configure how matrices should be created (merged or independent)
6. Compute and export to Excel

---

## Manual Start

If the batch file doesn't work, open a terminal and run:

```bash
python app.py
```

---

## Troubleshooting

**"Python is not installed"** - Download from https://python.org/

**App doesn't open in browser** - Manually open http://localhost:8080

**Package installation fails** - Run manually:
```bash
pip install pandas openpyxl
```
