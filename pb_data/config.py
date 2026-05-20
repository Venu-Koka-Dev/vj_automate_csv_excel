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

CSV_FILE_NAME = "pb_csv.txt"

EXCEL_FILE_NAME = "CRMTeam StatusReport_JUNE2026.xlsx"

# =========================================================
# EXCEL SETTINGS
# =========================================================

SHEET_NAME = "ProjectStatusReport(May)-AccSca"

START_ROW = 65

TARGET_COLUMNS = ["E", "F"]

# =========================================================
# CUSTOM PREFIX SETTINGS
# =========================================================

COLUMN_PREFIXES = {
    "E": "USERSTORY"
}