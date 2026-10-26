# 🎉 Automated PowerPoint Consolidation - Now Available!

## What's New

The MBR workflow now includes **fully automated PowerPoint consolidation**! No more manual slide copying.

---

## ✅ Before (Manual)

```bash
py main.py --mbr-workflow
# Creates instructions file
# You manually copy slides (2 minutes)
# Save consolidated_mbr_manual.pptx
```

---

## 🚀 After (Automated)

```bash
py main.py --mbr-workflow
# Creates EVERYTHING automatically
# Including consolidated_mbr.pptx - ready to use!
```

---

## How It Works

### Technology: Windows COM Automation

The system uses `win32com.client` to control Microsoft PowerPoint directly:

1. Opens the template PowerPoint
2. Opens each team presentation
3. Finds "Goals Achieved" slides
4. **Copies and pastes** slides programmatically
5. Saves the consolidated result

### Why COM Instead of python-pptx?

| Feature | python-pptx | COM Automation |
|---------|-------------|----------------|
| Formatting | ❌ Lost | ✅ Perfect |
| Images | ❌ Lost | ✅ Perfect |
| Charts | ❌ Lost | ✅ Perfect |
| Tables | ❌ Lost | ✅ Perfect |
| Layouts | ❌ Lost | ✅ Perfect |
| Animations | ❌ Lost | ✅ Perfect |
| Speed | Fast | Medium |
| Platform | Any | Windows only |

**Result:** COM automation = 100% fidelity, same as manual copying!

---

## Installation

### Required Package

```bash
py -m pip install pywin32
```

This package is now included in `requirements.txt`.

### Requirements

- **Windows OS** (COM automation is Windows-only)
- **Microsoft PowerPoint** installed
- **pywin32** package

---

## Usage

### For May (Default)

```bash
py main.py --mbr-workflow
```

### For Different Month

```bash
py mbr_assistant.py
# Then say: "Run everything for June"
```

Or:

```python
from mbr_workflow import run_mbr_workflow
run_mbr_workflow(month="June")
```

---

## Output Files

After running the workflow, you get:

1. **`graph_data_summary.xlsx`** - Graph metrics for all projects
2. **`goals_achieved_report.xlsx`** - Extracted goals content
3. **`consolidated_mbr.pptx`** - ✨ **AUTOMATED consolidated PowerPoint**
4. **`ppt_consolidation_instructions.txt`** - Backup manual instructions
5. **`mbr_workflow_summary.txt`** - Summary report

---

## Example Output

```
[STEP 1] Extracting graph data from LT-SimplifiedMetrics files...
  ✓ Extracted data for 6 projects

[STEP 2] Extracting 'May' goals from presentations...
  ✓ Found 1 goal slides

[STEP 3] Creating consolidated PowerPoint automatically...
  ✓ Consolidated PowerPoint created: output\consolidated_mbr.pptx

[STEP 4] Creating summary report...
  ✓ Summary report saved

✅ SUCCESS!
📊 Summary:
   Original template slides: 14
   Goal slides added: 1
   Final presentation slides: 15
```

---

## Fallback to Manual

If automated consolidation fails (e.g., PowerPoint not installed), the system will:

1. Still create the instruction file
2. Print a message explaining the issue
3. Allow you to consolidate manually (2 minutes)

---

## Technical Details

### New Module: `pptx_auto_consolidator.py`

Contains:
- `find_goal_slides_com()` - Find slides using COM
- `consolidate_pptx_com()` - Automated consolidation

### Updated Module: `mbr_workflow.py`

Now includes:
- Step 3: Automated PowerPoint consolidation
- Fallback to manual if automation fails
- Updated summary report

### Dependencies

Added to `requirements.txt`:
```
pywin32>=312
```

---

## Benefits

✅ **Zero manual work** - fully automated
✅ **Perfect fidelity** - preserves all formatting
✅ **Fast** - completes in seconds
✅ **Reliable** - same result every time
✅ **Fallback** - manual option if needed

---

## Migration Guide

If you were using the manual workflow:

### Old Way
```bash
py main.py --mbr-workflow
# Open output/ppt_consolidation_instructions.txt
# Manually copy slides
# Save as consolidated_mbr_manual.pptx
```

### New Way
```bash
py main.py --mbr-workflow
# Done! Open output/consolidated_mbr.pptx
```

**That's it!** 🎉

---

## Troubleshooting

### Error: "pywin32 not installed"

```bash
py -m pip install pywin32
```

### Error: "PowerPoint not found"

- Make sure Microsoft PowerPoint is installed
- Restart your computer after installing PowerPoint

### Error: "Permission denied"

- Close any open PowerPoint windows
- Run the command again

### Automation Fails

- Check `output/ppt_consolidation_instructions.txt`
- Follow manual consolidation steps (2 minutes)

---

**Enjoy fully automated MBR reporting!** 🚀
