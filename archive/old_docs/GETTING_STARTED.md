# Getting Started with Sprint MBR Reporting System

## 🎯 Quick Start Guide

### Step 1: Install Dependencies

```bash
py -m pip install pandas openpyxl python-pptx
```

### Step 2: Organize Your Data

Create a folder structure like this:
```
your_sprint_data/
├── Team_A_Sprint_Data.xlsx
├── Team_B_Sprint_Data.xlsx
├── Team_C_Sprint_Data.xlsx
├── LT-SimplifiedMetrics.xlsx
└── presentations/
    ├── Team_A_Goals_May.pptx
    ├── Team_B_Goals_May.pptx
    └── Final_MBR.pptx
```

### Step 3: Configure the System

Edit `config.py` to match your data:

```python
# Set your data folder
SOURCE_FOLDER = "./your_sprint_data"

# Enable MBR reporting
ENABLE_MBR_REPORTING = True

# Column names will be auto-matched using aliases
```

### Step 4: Run the System

```bash
py main.py
```

## 📊 What Gets Generated

After running, you'll find these files in the `output/` folder:

1. **`master_file.xlsx`** - Legacy consolidated data
2. **`dashboard.html`** - Interactive web dashboard
3. **`mbr_report.xlsx`** - ⭐ NEW Sprint MBR Report with:
   - Consolidated Data with calculated KPIs
   - KPI Summary
   - Sprint Trends with charts
   - Goals Achieved (extracted from PPTs)
   - Run Log

## 🔧 Understanding Your Data

### Required Columns

Your Excel files should have columns like:
- Sprint/Week (e.g., "SEOD 10/07 - 10/21")
- Planned Story Points
- Actual Delivered
- V2, LT (if tracking separately)
- Rework
- Reopened SP
- Carryover
- QA Passed

### Column Matching

The system automatically matches different column names:
- "planned story points" = "planned sp" = "Planned_SP"
- "actual delivered" = "delivered" = "Actual_Delivered"

All matching is case-insensitive!

## 📈 KPIs Calculated

The system automatically calculates:

1. **Reopened %** = Reopened SP / Delivered SP
2. **Carryover %** = Carryover / Planned SP
3. **Completion Ratio** = Actual / Planned
4. **QA Pass Rate** = QA Passed / Total Delivered

## 🎤 PowerPoint Goal Extraction

The system extracts slides with titles matching:
- "Goals Achieved in May" (or any month)
- "Sprint Goals"
- "Accomplishments"

All content from matching slides is consolidated in the final report.

## ⚙️ Advanced Configuration

### Customize Goal Slide Patterns

```python
PPTX_GOAL_SLIDE_PATTERNS = [
    "Goals Achieved in",
    "Sprint Goals",
    "Your Custom Pattern",
]
```

### Filter by Teams

```python
TEAMS_TO_INCLUDE = ["LendingTree", "SEOD"]  # Empty = all teams
```

### Adjust Analysis Period

```python
ANALYSIS_PERIOD_MONTHS = 6  # Look back 6 months
```

## 🐛 Troubleshooting

### "Column not found" errors
- Check `config.py` COLUMN_ALIASES
- Add your column names to the appropriate alias list

### "Cannot convert to float" errors
- The system now auto-extracts numbers from text
- e.g., "5 points" → 5

### No goal slides extracted
- Check slide titles match PPTX_GOAL_SLIDE_PATTERNS
- Slides must be in .pptx files in SOURCE_FOLDER

## 📁 Sample Data

Sample data is included in `sample_data/` folder:
- `LT-SimplifiedMetrics(SEOD).csv` - Sprint metrics example
- `orders_*.xlsx` - Legacy test data
- `q1_review.pptx` - PowerPoint extraction example

## 🚀 Next Steps

1. **Customize column mappings** in config.py
2. **Add your team data** to a folder
3. **Run the system** and check the MBR report
4. **Review the Run_Log sheet** for any warnings

## 💡 Tips

- Always check the **Run_Log** sheet first
- The system handles mixed data types automatically
- Charts are generated if GENERATE_CHARTS = True
- Legacy mode works alongside MBR reporting

## 📞 Need Help?

Check the main README.md for full documentation and feature details.
