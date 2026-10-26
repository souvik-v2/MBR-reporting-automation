# What's New - MBR Workflow Features

## 🎉 New Features Added

### 1. **MBR Workflow** (Complete Automation)

A new end-to-end workflow that automates the creation of Management Business Review deliverables.

**Run it with:**
```bash
py main.py --mbr-workflow
```

**What it does:**
- Extracts graph metrics from multiple project files
- Consolidates PowerPoint presentations
- Generates summary reports
- Creates a final consolidated PowerPoint with template

---

### 2. **Graph Data Extractor** (`graph_data_extractor.py`)

Automatically reads and extracts data from `LT-SimplifiedMetrics-{project}.xlsx` files.

**Extracts:**
- V2 Carryover
- LT Carryover
- Rework
- Reopened SP (LT)
- Week/Sprint information

**Supports Projects:**
- CanopyNext
- ConsumerWeb
- FLEX
- Phoenix
- SEOD
- WPM

**Output:**
- `graph_data_summary.xlsx` - Consolidated metrics from all projects

---

### 3. **PowerPoint Consolidator** (`pptx_consolidator.py`)

Automatically extracts "Goals Achieved" slides from team presentations and inserts them into an MBR template.

**Features:**
- Finds slides matching "Goals Achieved in {Month}"
- Copies slides with all formatting, shapes, and content
- Inserts into template at specified position
- Maintains original slide design

**Output:**
- `consolidated_mbr.pptx` - Final presentation with goals slides

---

### 4. **Complete Workflow Orchestrator** (`mbr_workflow.py`)

Coordinates all MBR tasks in a single execution.

**Process:**
1. Extract graph data from all project files
2. Create consolidated PowerPoint
3. Generate summary reports
4. Save detailed logs

**Outputs:**
- `graph_data_summary.xlsx`
- `consolidated_mbr.pptx`
- `mbr_workflow_summary.txt`
- `mbr_workflow_log.txt`

---

## How It Works

### Input Structure

```
sample_data/
├── Graph/                           # Graph metrics files
│   ├── LT-SimplifiedMetrics-CanopyNext.xlsx
│   ├── LT-SimplifiedMetrics-consumerweb.xlsx
│   ├── LT-SimplifiedMetrics-flex.xlsx
│   ├── LT-SimplifiedMetrics-phoenix.xlsx
│   ├── LT-SimplifiedMetrics-seod.xlsx
│   └── LT-SimplifiedMetrics-wpm.xlsx
│
├── ppt-data/                        # Team presentations
│   ├── V2Solutions_SEOD_May_2026_Report.pptx
│   └── [Other team presentations...]
│
└── final-data/                      # MBR template
    └── 06-01-2026.pptx
```

### Output Structure

```
output/
├── graph_data_summary.xlsx          # Extracted metrics
├── consolidated_mbr.pptx             # Final presentation
├── mbr_workflow_summary.txt          # Human-readable summary
└── mbr_workflow_log.txt              # Detailed execution log
```

---

## Usage Examples

### Basic Usage

```bash
# Run complete MBR workflow
py main.py --mbr-workflow
```

### Custom Month

```python
from mbr_workflow import run_mbr_workflow

run_mbr_workflow(month="June")
```

### Extract Graph Data Only

```python
from graph_data_extractor import extract_all_graph_data

data = extract_all_graph_data("sample_data/Graph")
print(data)
```

### Consolidate PowerPoint Only

```python
from pptx_consolidator import create_consolidated_pptx

create_consolidated_pptx(
    template_path="template.pptx",
    ppt_data_folder="presentations",
    output_path="output.pptx",
    month="May"
)
```

---

## Configuration

### Change Data Folders

Edit `mbr_workflow.py`:

```python
run_mbr_workflow(
    graph_folder="your_path/Graph",
    ppt_folder="your_path/ppt-data",
    template_path="your_path/template.pptx",
    output_folder="output",
    month="May"
)
```

### Change PPT Insertion Point

Edit `pptx_consolidator.py`:

```python
create_consolidated_pptx(
    ...,
    insert_after_slide=3  # Insert after slide 3 instead of 5
)
```

---

## Integration with Existing Features

The MBR Workflow **complements** the existing Sprint MBR Reporting features:

| Feature | What It Does | How to Run |
|---------|-------------|------------|
| **Original MBR Report** | Sprint KPI calculation, charts, metrics | `py main.py` (default) |
| **MBR Workflow** | Graph data extraction + PPT consolidation | `py main.py --mbr-workflow` |

Both can be run independently or together!

---

## Troubleshooting

### Common Issues

1. **No goal slides found**
   - Check slide titles contain "Goals Achieved"
   - Verify month name matches

2. **Graph data errors**
   - Ensure Excel files follow naming: `LT-SimplifiedMetrics-{project}.xlsx`
   - Check row 2 has headers, row 3 has data

3. **PowerPoint errors**
   - Close PowerPoint files before running
   - Verify template file exists
   - Check write permissions

### Debug Logs

Check `output/mbr_workflow_log.txt` for detailed error messages and execution traces.

---

## Documentation

- **`MBR_WORKFLOW_GUIDE.md`** - Complete guide for the MBR workflow
- **`README.md`** - Main project documentation
- **`GETTING_STARTED.md`** - Quick start guide

---

## Next Steps

1. ✅ Run `py main.py --mbr-workflow` to test with sample data
2. ✅ Review generated files in `output/` folder
3. ✅ Customize for your data structure if needed
4. ✅ Read `MBR_WORKFLOW_GUIDE.md` for advanced usage

---

**Questions?** Check the documentation files or review the logs for detailed information.
