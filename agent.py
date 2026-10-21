from pathlib import Path
import json
import os
from collections import Counter

import pandas as pd


BASE_DIR = Path(__file__).parent
OUTPUT_DIR = BASE_DIR / "output"

# Total HC is not available in Source A or Source B.
# It can be overridden at deployment level without changing code:
# CSU_TOTAL_HC=761
DEFAULT_TOTAL_HC = int(os.getenv("CSU_TOTAL_HC", "761"))

MONTH_NAMES = {
    1: "January",
    2: "February",
    3: "March",
    4: "April",
    5: "May",
    6: "June",
    7: "July",
    8: "August",
    9: "September",
    10: "October",
    11: "November",
    12: "December",
}

SOURCE_A_REQUIRED = [
    "Legal Entity Name",
    "Work Type Name",
    "Timesheets Billable",
    "Timesheets Non Billable",
    "All Timesheets",
    "Completion %",
    "Man Period Name",
]

SOURCE_B_REQUIRED = [
    "Person Number",
    "Mois",
    "Year",
    "Time Entry",
]

# Historical PoC controls that were already validated.
# These allow Jul/Aug 2026 to continue working even when a newer
# Source B upload contains only Sep/Oct.
HISTORICAL_CONTROLS = {
    (2026, 7): {
        "hc_tagged": 656,
        "total_hc": 761,
        "expected_working_days": 23.0,
        "hc_compliant": 633,
        "overbooked": 21,
        "control_source": "validated historical control",
    },
    (2026, 8): {
        "hc_tagged": 656,
        "total_hc": 761,
        "expected_working_days": 21.0,
        "hc_compliant": 623,
        "overbooked": 40,
        "control_source": "validated historical control",
    },
}


def _clean_header(value):
    if value is None:
        return ""
    return " ".join(str(value).replace("\n", " ").split()).strip()


def load_sheet_containing_columns(path, required_columns):
    """
    Find and load the worksheet containing all required columns.
    Also tolerates title rows above the real header.
    """
    path = Path(path)
    if not path.exists():
        raise FileNotFoundError(f"Input file not found: {path}")

    excel = pd.ExcelFile(path, engine="openpyxl")
    required = {_clean_header(c) for c in required_columns}
    checked = []

    for sheet_name in excel.sheet_names:
        preview = pd.read_excel(
            path,
            sheet_name=sheet_name,
            header=None,
            nrows=20,
