"""
Configuration for the Excel/CSV/PPTX merger & Sprint MBR Reporting System.

Edit the values below to match your data before running main.py.
"""

# ── Folder paths ─────────────────────────────────────────────────────────
SOURCE_FOLDER = "./sample_data"        # folder to scan for .xlsx/.xls/.csv/.pptx
OUTPUT_FILE = "./output/master_file.xlsx"
DASHBOARD_FILE = "./output/dashboard.html"   # standalone HTML dashboard, same data
MBR_REPORT_FILE = "./output/mbr_report.xlsx"  # MBR report with KPIs and charts

# ── Column matching ──────────────────────────────────────────────────────
# Column names are matched case-insensitively and with whitespace stripped.
# List every name you'd expect to mean "the same column" together.
COLUMN_ALIASES = {
    "ID": ["id", "record id", "order id", "ref", "reference", "ref no", "ref_no"],
    "Date": ["date", "order date", "date/time", "datetime", "timestamp", "updated"],
    # Sprint-specific column aliases
    "Sprint": ["sprint", "week starting on", "week", "sprint period", "iteration"],
    "Team": ["team", "team name", "project", "project name", "sub project"],
    "Planned_SP": ["planned story points", "planned sp", "planned", "planned points"],
    "Actual_Delivered": ["actual delivered", "delivered", "actual sp", "completed sp"],
    "V2": ["v2", "v2 sp", "v2 points"],
    "LT": ["lt", "lt sp", "lt points"],
    "Rework": ["rework", "rework sp"],
    "Reopened_LT": ["reopened sp (lt)", "reopened lt", "lt reopened", "reopened"],
    "V2_Carryover": ["v2 carryover", "v2 carry over"],
    "LT_Carryover": ["lt carryover", "lt carry over"],
    "QA_Passed": ["qa passed", "qa pass", "quality passed"],
    "Delivery_Manager": ["delivery manager", "dm", "manager"],
    "Project_Manager": ["project manager/lead", "project manager", "pm", "lead"],
    "Status": ["status", "overall status", "project status"],
}

# ── Deduplication ─────────────────────────────────────────────────────────
# Canonical column name (after alias resolution) used to identify "the same
# record" across files. Rows sharing this value will be deduped, keeping
# only the one with the latest DEDUP_DATE_COLUMN.
DEDUP_ID_COLUMN = "ID"

# Canonical column name used to decide which duplicate is "latest".
DEDUP_DATE_COLUMN = "Date"

# If a row has no usable ID value, what to do: "keep" (treat as unique,
# never deduped) or "drop" (exclude from Master sheet, still in All_Records).
NO_ID_POLICY = "keep"

# If True, rows whose date value can't be parsed are still kept (placed
# last / treated as oldest). If False, they're excluded from the Master
# sheet (still appear in All_Records).
KEEP_UNPARSEABLE_DATES = True

# ── Date parsing ──────────────────────────────────────────────────────────
# Slash-separated dates like 07/01/2026 are ambiguous (day/month or
# month/day?). Set this explicitly to match your data instead of letting
# the script guess:
#   False -> month/day/year (US style, e.g. 07/01/2026 = Jan 7)
#   True  -> day/month/year (e.g. 07/01/2026 = Jul 1)
DATE_DAYFIRST = False

# Extract tables found inside slides into the same row-structure as the
# Excel/CSV data (only works well if the table headers match known columns).
EXTRACT_PPTX_TABLES = True

# Extract slide titles / body text into a separate PPTX_Text sheet.
EXTRACT_PPTX_TEXT = True

# Extract embedded chart titles + category/series data into PPTX_Charts.
EXTRACT_PPTX_CHARTS = True

# ── File types to scan ────────────────────────────────────────────────────
EXCEL_EXTENSIONS = (".xlsx", ".xlsm", ".xls")
CSV_EXTENSIONS = (".csv", ".tsv")
PPTX_EXTENSIONS = (".pptx",)

# ── Sprint MBR Reporting Configuration ───────────────────────────────────

# KPI Calculation Settings
ENABLE_MBR_REPORTING = True  # Set to False to use legacy mode only

# KPI Formulas (these match your requirements)
# Reopened % = Reopened SP / Delivered SP
# Carryover % = Carryover / Planned SP
# Completion Ratio = Actual / Planned
# QA metrics tracking enabled by default

# PowerPoint slide extraction patterns
PPTX_GOAL_SLIDE_PATTERNS = [
    "Goals Achieved in",  # Matches "Goals Achieved in May", "Goals Achieved in June", etc.
    "Goals Achieved",
    "Sprint Goals",
    "Accomplishments",
]

# Chart generation settings
GENERATE_CHARTS = True
CHART_TYPES = [
    "sprint_trend",          # Sprint/Kanban trend (week-wise)
    "team_comparison",       # Team comparison charts
    "planned_vs_actual",     # Planned vs Actual delivery
    "quality_metrics",       # Quality metrics trends
    "rework_trend",          # Rework trend over time
    "reopened_trend",        # Reopened trend over time
    "v2_quality",            # V2 quality metrics
]

# Time period for analysis (months to look back from latest data)
ANALYSIS_PERIOD_MONTHS = 6

# Team filters (leave empty to include all teams)
TEAMS_TO_INCLUDE = []  # e.g., ["LendingTree", "SEOD", "Consumer Web"]

# Month name formats for goal slide extraction
MONTH_NAMES = [
    "January", "February", "March", "April", "May", "June",
    "July", "August", "September", "October", "November", "December"
]

# Output sheet names for MBR report
MBR_SHEET_NAMES = {
    "consolidated": "Consolidated Data",
    "kpi_summary": "KPI Summary",
    "sprint_trends": "Sprint Trends",
    "team_comparison": "Team Comparison",
    "quality_metrics": "Quality Metrics",
    "goals_achieved": "Goals Achieved",
    "run_log": "Run Log",
}
