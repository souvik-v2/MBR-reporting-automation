# MBR Workflow Guide

## Overview

The MBR (Management Business Review) Workflow is a specialized automation that:

1. **Extracts Graph Data** from multiple project-specific Excel files
2. **Consolidates PowerPoint Presentations** by extracting "Goals Achieved" slides
3. **Generates Summary Reports** in both Excel and text formats

## What It Does

### 📊 Graph Data Extraction

Automatically reads data from `LT-SimplifiedMetrics-{project}.xlsx` files and extracts:

- **V2 Carryover**: Carryover story points from V2
- **LT Carryover**: Carryover story points from LendingTree
- **Rework**: Rework story points
- **Reopened SP (LT)**: Reopened story points for LendingTree
- **Week**: The week/sprint period

Supports multiple projects:

- CanopyNext
- ConsumerWeb
- FLEX
- Phoenix
- SEOD
- WPM

### 🎤 PowerPoint Consolidation

Automatically:

1. Scans team presentation files for slides with "Goals Achieved in {Month}"
2. Extracts matching slides
3. Inserts them into a template PowerPoint at the specified position
4. Creates a final consolidated presentation

### 📄 Output Files

The workflow generates:

1. **`graph_data_summary.xlsx`**
   Excel file with extracted metrics from all projects (V2/LT Carryover, Rework, Reopened SP)

2. **`goals_achieved_report.xlsx`**
   Excel file with all content extracted from "Goals Achieved" slides (titles, text, tables)

3. **`ppt_consolidation_instructions.txt`**
   Step-by-step guide for manually consolidating PowerPoint slides
   _(Due to python-pptx limitations, slides must be copied manually to preserve formatting)_

4. **`mbr_workflow_summary.txt`**
   Human-readable summary of extracted data and process

5. **`mbr_workflow_log.txt`**
   Detailed execution log with timestamps

## How to Use

### Quick Start

```bash
# Run the MBR workflow
py main.py --mbr-workflow
```

### Custom Parameters

You can customize the workflow by editing `mbr_workflow.py`:

```python
run_mbr_workflow(
    graph_folder="your_graph_data_folder",
    ppt_folder="your_ppt_folder",
    template_path="your_template.pptx",
    output_folder="output",
    month="June"  # Change the month to extract
)
```

## File Structure

### Input Files Expected

```
sample_data/
├── Graph/
│   ├── LT-SimplifiedMetrics-CanopyNext.xlsx
│   ├── LT-SimplifiedMetrics-consumerweb.xlsx
│   ├── LT-SimplifiedMetrics-flex.xlsx
│   ├── LT-SimplifiedMetrics-phoenix.xlsx
│   ├── LT-SimplifiedMetrics-seod.xlsx
│   └── LT-SimplifiedMetrics-wpm.xlsx
├── ppt-data/
│   ├── Team_A_May_Report.pptx
│   ├── Team_B_May_Report.pptx
│   └── ...
└── final-data/
    └── 06-01-2026.pptx  # Template PowerPoint
```

### Graph Data File Format

Each `LT-SimplifiedMetrics-{project}.xlsx` should have:

- Row 2 (header): Column names
- Row 3 (data): Latest week's data

Required columns:

- Week Starting On
- V2 Carryover
- LT Carryover
- Rework
- Reopened SP (LT)

### PowerPoint Files

Team presentation files should:

- Be in `.pptx` format
- Contain slides with titles like:
  - "Goals Achieved in May"
  - "Goals Achieved in June"
  - "May Goals"
  - "May Achievements"

## Advanced Usage

### Change the Month

To extract goals for a different month:

```python
from mbr_workflow import run_mbr_workflow

run_mbr_workflow(month="June")
```

### Change Insertion Point

To insert goals slides at a different position in the template:

```python
from pptx_consolidator import create_consolidated_pptx

create_consolidated_pptx(
    template_path="template.pptx",
    ppt_data_folder="presentations",
    output_path="output.pptx",
    month="May",
    insert_after_slide=3  # Insert after slide 3
)
```

### Extract Graph Data Only

```python
from graph_data_extractor import extract_all_graph_data, create_graph_data_dataframe

graph_data = extract_all_graph_data("sample_data/Graph")
df = create_graph_data_dataframe(graph_data)
df.to_excel("graph_summary.xlsx", index=False)
```

## Troubleshooting

### No goal slides found

- Check that PPT files contain slides with "Goals Achieved" in the title
- Verify the month name matches (case-insensitive)
- Check PPT file names don't start with `~$` (temporary files are skipped)

### Graph data extraction errors

- Verify Excel files are named `LT-SimplifiedMetrics-{project}.xlsx`
- Check that row 2 contains column headers
- Ensure row 3 contains the latest week's data
- Required columns must be present

### PowerPoint consolidation issues

- Template file must exist and be accessible
- Ensure PowerPoint files are not open in another application
- Check that you have write permissions in the output folder

## Integration with Main Workflow

The MBR workflow runs independently and can be executed:

1. **Standalone**: `py main.py --mbr-workflow`
2. **As part of main process**: The regular `py main.py` still works and generates the MBR report
3. **Programmatically**: Import and call `run_mbr_workflow()` from your own scripts

## Output Example

```
output/
├── graph_data_summary.xlsx                 # ✓ Graph metrics for all projects
├── goals_achieved_report.xlsx              # ✓ Extracted goals content (text, tables, etc.)
├── ppt_consolidation_instructions.txt      # ✓ Manual consolidation guide
├── mbr_workflow_summary.txt                # ✓ Human-readable summary
└── mbr_workflow_log.txt                    # ✓ Detailed log
```

**Note on PowerPoint Consolidation:**

Due to limitations in the `python-pptx` library, automated slide copying cannot preserve complex formatting, images, and charts reliably. The workflow instead:

1. **Identifies** which slides contain goals (e.g., "Goals Achieved in May")
2. **Extracts** all text content, tables, and metadata to an Excel report
3. **Generates** step-by-step instructions for manual consolidation

**To create the final consolidated PowerPoint:**

- Follow the instructions in `ppt_consolidation_instructions.txt`
- This takes ~2 minutes and preserves all formatting perfectly

## Next Steps

After running the workflow:

1. **Review** the generated `mbr_workflow_summary.txt`
2. **Open** `consolidated_mbr.pptx` to verify goals slides were inserted correctly
3. **Check** `graph_data_summary.xlsx` for accuracy
4. **Customize** the template or parameters as needed for your team

---

**For questions or issues**: Check the logs in `mbr_workflow_log.txt` for detailed error messages.
