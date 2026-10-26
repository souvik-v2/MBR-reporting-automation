# 🚀 Quick Start - MBR Workflow

## Get Started in 2 Minutes

### Step 1: Run the Workflow

```bash
py main.py --mbr-workflow
```

### Step 2: Check Your Outputs

All files are in the `output/` folder:

#### ✅ Graph Data Summary

Open `graph_data_summary.xlsx`:

- See metrics for all 6 projects
- V2 Carryover, LT Carryover, Rework, Reopened SP
- Latest week's data automatically extracted

#### ✅ Goals Achieved Report

Open `goals_achieved_report.xlsx`:

- All "Goals Achieved in May" content
- Text, tables, and metadata
- Source file and slide number for each goal

#### ✅ Consolidated PowerPoint (AUTOMATED!)

Open `consolidated_mbr.pptx`:

- **Automatically created** - no manual work needed!
- All goal slides inserted with perfect formatting
- Images, charts, and layouts preserved
- Ready to present!

> **How?** Uses Windows COM automation to control PowerPoint directly

---

## ✨ What's New - Automated PowerPoint Consolidation!

### No Manual Work Required!

The workflow now **automatically creates** the consolidated PowerPoint using Windows COM automation.

### Requirements

- Microsoft PowerPoint installed (Windows only)
- `pywin32` package (auto-installed with requirements)

### Fallback to Manual

If automated consolidation fails, instructions are in:

```
output/ppt_consolidation_instructions.txt
```

Follow the 2-minute manual process as a backup.

---

## What Each File Contains

### `graph_data_summary.xlsx`

```
Project      | Week           | V2 Carryover | LT Carryover | Rework | Reopened SP
-----------------------------------------------------------------------------
CANOPYNEXT   | LCW 4/6 - 4/20 | 5            | 11           | 2      | 0
CONSUMERWEB  | LCW 4/6 - 4/20 | 5            | 11           | 2      | 0
FLEX         | LCW 4/6 - 4/20 | 5            | 11           | 2      | 0
...
```

### `goals_achieved_report.xlsx`

```
Source File                              | Slide | Title                  | Text Content
--------------------------------------------------------------------------------------------
V2Solutions_SEOD_May_2026_Report.pptx   | 3     | Goals Achieved in May  | [Full content]
```

### `ppt_consolidation_instructions.txt`

```
SLIDES TO EXTRACT:
1. File: V2Solutions_SEOD_May_2026_Report.pptx
   Slide: #3
   Title: Goals Achieved in May

MANUAL CONSOLIDATION STEPS:
1. Open template: sample_data/final-data/06-01-2026.pptx
2. Copy slide #3 from source file
3. Paste into template after slide 5
4. Save as: consolidated_mbr_manual.pptx
```

---

## Customization

### Change the Month

Edit `mbr_workflow.py`:

```python
run_mbr_workflow(month="June")  # Change "May" to "June"
```

### Change Data Folders

Edit `mbr_workflow.py`:

```python
run_mbr_workflow(
    graph_folder="your_data/Graph",
    ppt_folder="your_data/presentations",
    template_path="your_data/template.pptx"
)
```

---

## Troubleshooting

### No goal slides found?

- Check that slides have "Goals Achieved" in the title
- Verify the month name matches (case-insensitive)
- Ensure PPT files are in `sample_data/ppt-data/`

### Graph data errors?

- Files must be named `LT-SimplifiedMetrics-{project}.xlsx`
- Check that row 2 has headers, row 3 has data
- Required columns must be present

### Want more details?

- See `output/mbr_workflow_log.txt` for execution details
- Read `MBR_WORKFLOW_GUIDE.md` for complete documentation

---

## Next Steps

1. ✅ Review the generated Excel files
2. ✅ Follow PPT consolidation instructions
3. ✅ Customize for your actual data
4. ✅ Run monthly to automate your MBR process

---

## Full Documentation

- **`MBR_WORKFLOW_GUIDE.md`** - Complete guide
- **`FINAL_SUMMARY.md`** - Implementation details
- **`WHATS_NEW.md`** - Features and examples
- **`README.md`** - Main project documentation

---

**Questions?** Check the logs and documentation files!
