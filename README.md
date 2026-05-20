# CSV to Excel Automation using Python

This project automates loading data from a CSV file into an Excel workbook using Python.

The script:

- Reads data from a CSV file
- Inserts it into a specific Excel sheet
- Starts insertion from a configurable row
- Writes into configurable columns
- Adds prefixes dynamically (Example: `USERSTORY`)
- Uses a separate configuration file for easy maintenance

---

# Project Structure

```text
pb_data/
│
├── automate_csv_excel.py
├── config.py
├── pb_csv.csv
└── CRMTeam StatusReport_JUNE2026.xlsx
```

---

# Fresh Windows Setup Guide

## Step 1 — Install Python

Download Python from:

https://www.python.org/downloads/

### Important During Installation

While installing:

✅ Check:

```text
Add Python to PATH
```

Then click:

```text
Install Now
```

---

## Step 2 — Verify Python Installation

Open Command Prompt and run:

```bash
python --version
```

Expected Output:

```text
Python 3.x.x
```

Also verify pip:

```bash
pip --version
```

---

# Step 3 — Create Project Folder

Create this folder manually:

```text
C:\Users\Lenovo\OneDrive\Desktop\pb_data
```

---

# Step 4 — Copy Required Files

Place these files inside `pb_data`:

```text
automate_csv_excel.py
config.py
pb_csv.csv
CRMTeam StatusReport_JUNE2026.xlsx
```

---

# Step 5 — Install Required Python Packages

Open Command Prompt and run:

```bash
pip install pandas openpyxl
```

Packages used:

| Package | Purpose |
|---|---|
| pandas | Read CSV files |
| openpyxl | Read/write Excel files |

---

# Step 6 — Create `config.py`

Create a file named:

```text
config.py
```

Paste this:

```python
from pathlib import Path

# =========================================================
# MAIN DATA FOLDER
# =========================================================

DATA_FOLDER = Path(
    r"C:\Users\Lenovo\OneDrive\Desktop\pb_data"
)

# =========================================================
# FILE NAMES
# =========================================================

CSV_FILE_NAME = "pb_csv.csv"

EXCEL_FILE_NAME = "CRMTeam StatusReport_JUNE2026.xlsx"

# =========================================================
# EXCEL SETTINGS
# =========================================================

SHEET_NAME = "ProjectStatusReport(May)-AccSca"

START_ROW = 65

TARGET_COLUMNS = ["E", "F"]

# =========================================================
# COLUMN PREFIXES
# =========================================================

COLUMN_PREFIXES = {
    "E": "USERSTORY"
}
```

---

# Step 7 — Create `automate_csv_excel.py`

Create file:

```text
automate_csv_excel.py
```

Paste this:

```python
"""
CSV to Excel Automation Script
"""

import pandas as pd

from openpyxl import load_workbook
from openpyxl.utils import column_index_from_string

from config import (
    DATA_FOLDER,
    CSV_FILE_NAME,
    EXCEL_FILE_NAME,
    SHEET_NAME,
    START_ROW,
    TARGET_COLUMNS,
    COLUMN_PREFIXES
)


# =========================================================
# BUILD FILE PATHS
# =========================================================

csv_path = DATA_FOLDER / CSV_FILE_NAME

excel_path = DATA_FOLDER / EXCEL_FILE_NAME


# =========================================================
# DEBUG INFORMATION
# =========================================================

print("======================================")
print("CSV Path:", csv_path)
print("Excel Path:", excel_path)
print("CSV Exists:", csv_path.exists())
print("Excel Exists:", excel_path.exists())
print("======================================")


# =========================================================
# VALIDATE FILES
# =========================================================

if not csv_path.exists():
    raise FileNotFoundError(
        f"CSV file not found:\n{csv_path}"
    )

if not excel_path.exists():
    raise FileNotFoundError(
        f"Excel file not found:\n{excel_path}"
    )


# =========================================================
# READ CSV FILE
# =========================================================

print(f"Reading CSV file:\n{csv_path}")

# Include first row also as data
df = pd.read_csv(csv_path, header=None)

if df.empty:
    raise Exception("CSV file is empty.")

required_column_count = len(TARGET_COLUMNS)

if len(df.columns) < required_column_count:
    raise Exception(
        f"CSV must contain at least "
        f"{required_column_count} columns."
    )

# Only use required columns
df = df.iloc[:, :required_column_count]


# =========================================================
# LOAD EXCEL WORKBOOK
# =========================================================

print(f"Opening Excel workbook:\n{excel_path}")

workbook = load_workbook(excel_path)

if SHEET_NAME not in workbook.sheetnames:
    raise Exception(
        f"Sheet '{SHEET_NAME}' not found."
    )

worksheet = workbook[SHEET_NAME]


# =========================================================
# WRITE DATA INTO EXCEL
# =========================================================

print("Writing data into Excel...")

for row_offset, row_data in enumerate(df.values):

    excel_row = START_ROW + row_offset

    for col_offset, value in enumerate(row_data):

        excel_column_letter = TARGET_COLUMNS[col_offset]

        excel_column_number = column_index_from_string(
            excel_column_letter
        )

        # Apply Prefix
        prefix = COLUMN_PREFIXES.get(
            excel_column_letter
        )

        if prefix:
            value = f"{prefix}{value}"

        worksheet.cell(
            row=excel_row,
            column=excel_column_number,
            value=value
        )


# =========================================================
# SAVE OUTPUT
# =========================================================

output_file = (
    DATA_FOLDER /
    "CRMTeam StatusReport_JUNE2026_UPDATED.xlsx"
)

workbook.save(output_file)


# =========================================================
# SUCCESS
# =========================================================

print("======================================")
print("Data inserted successfully.")
print("Output File:")
print(output_file)
print("======================================")
```

---

# Step 8 — Run the Script

Open Command Prompt:

```bash
cd C:\Users\Lenovo\OneDrive\Desktop\pb_data
```

Then run:

```bash
python automate_csv_excel.py
```

---

# Expected Successful Output

```text
======================================
CSV Path: ...
Excel Path: ...
CSV Exists: True
Excel Exists: True
======================================

Reading CSV file...
Opening Excel workbook...
Writing data into Excel...

======================================
Data inserted successfully.
Output File:
CRMTeam StatusReport_JUNE2026_UPDATED.xlsx
======================================
```

---

# Customization

Edit only `config.py`.

## Change Excel Sheet

```python
SHEET_NAME = "WeeklyReport"
```

---

## Change Start Row

```python
START_ROW = 100
```

---

## Change Columns

```python
TARGET_COLUMNS = ["B", "C"]
```

---

## Change Prefixes

```python
COLUMN_PREFIXES = {
    "B": "TASK-",
    "C": "BUG-"
}
```

---

# Troubleshooting

## Error: `ModuleNotFoundError`

Install packages:

```bash
pip install pandas openpyxl
```

---

## Error: `CSV file not found`

Verify file exists:

```text
pb_csv.csv
```

inside:

```text
C:\Users\Lenovo\OneDrive\Desktop\pb_data
```

---

## Error: `Sheet not found`

Verify sheet name exactly matches:

```python
SHEET_NAME
```

including spaces and brackets.

---

# Future Enhancements

Possible future improvements:

- Multiple sheets support
- Multiple CSV files
- Excel formatting
- Logging
- GUI interface
- Auto email reports
- Scheduler automation
- Database integration

---

# Technologies Used

- Python
- Pandas
- OpenPyXL

---

# License

MIT License
