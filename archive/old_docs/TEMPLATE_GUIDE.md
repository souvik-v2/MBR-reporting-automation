# Template PowerPoint Guide

## Understanding the Template

### Current Setup

The template file `sample_data/final-data/06-01-2026.pptx` is used as a **reference/base** for the consolidated MBR presentation.

### What Happens With Template

When you run:
```bash
py main.py --mbr-workflow
```

The system:
1. ✅ Opens the template PowerPoint
2. ✅ Finds "Goals Achieved" slides from team presentations
3. ✅ **Copies goal slides** into the template
4. ✅ Saves as `output/consolidated_mbr.pptx`

**The template is NOT modified** - it's used as a starting point.

---

## ✅ Can the System Work Without Template?

**YES!** The system has been updated to work with or without a template.

### Scenario 1: Template Exists (Recommended)

```bash
# Template: sample_data/final-data/06-01-2026.pptx exists
py main.py --mbr-workflow
```

**Result:**
- Opens template (14 slides)
- Adds goal slides from teams
- Final presentation = Template + Goal slides

**Example:**
```
Template slides (14):
  1. Monthly Business Review
  2. Delivery Snapshot
  ...
  14. Thank You

Goal slides found (1):
  - SEOD: Goals Achieved in May

Final consolidated (15 slides):
  1-5. [Template slides]
  6. Goals Achieved in May  ← NEW from SEOD
  7-15. [Rest of template]
```

---

### Scenario 2: Template Does NOT Exist

```bash
# Template file deleted or missing
py main.py --mbr-workflow
```

**Result:**
- ⚠️ Detects missing template
- ℹ️ Creates **new blank presentation**
- Adds goal slides from teams
- Final presentation = Only goal slides (no template)

**Example:**
```
Template: NOT FOUND

Goal slides found (1):
  - SEOD: Goals Achieved in May

Final consolidated (1 slide):
  1. Goals Achieved in May  ← Only this slide
```

---

## Use Cases

### Use Case 1: Full MBR Report (With Template)

**When:** You want a complete MBR presentation with:
- Cover slides
- Summary slides
- Project metrics
- Goal slides from teams
- Thank you slide

**Setup:**
```
✅ Keep template: sample_data/final-data/06-01-2026.pptx
✅ Run: py main.py --mbr-workflow
```

**Output:** Complete 15-slide presentation

---

### Use Case 2: Goals Only (Without Template)

**When:** You only want goal slides from teams, nothing else

**Setup:**
```
❌ Delete or move template
✅ Run: py main.py --mbr-workflow
```

**Output:** Presentation with only goal slides (1+ slides)

---

## Configuration

### Option 1: Change Template Path

Edit `mbr_workflow.py` line 28:

```python
template_path: str = "path/to/your/template.pptx"
```

### Option 2: Use No Template

Call workflow with no template:

```python
from mbr_workflow import run_mbr_workflow

run_mbr_workflow(
    template_path=None,  # No template
    use_template=False,
    month="May"
)
```

### Option 3: Use Chatbot

```bash
py mbr_assistant.py
```

Then configure the template path or disable template usage.

---

## Recommendations

### For Production MBR

✅ **Use a template** for consistent, professional reports:
- Company branding
- Standard headers/footers
- Consistent formatting
- Overview slides
- Project metrics

### For Quick Goal Reviews

❌ **Skip the template** if you only need:
- Quick goal slide compilation
- Testing
- Temporary reports

---

## Template Structure

### Recommended Template Content

Your template should include:

1. **Title/Cover Slide** - MBR title, date, team
2. **Overview Slides** - Summary, highlights
3. **Placeholder for Goals** - Space where goal slides will be inserted
4. **Project Metrics** - Standard project slides (optional)
5. **Closing Slide** - Thank you, next steps

### Where Goal Slides Are Inserted

By default: **After slide 5**

You can change this in the configuration:

```python
consolidate_pptx_com(
    template_path="...",
    insert_after_slide=3,  # Insert after slide 3 instead
    ...
)
```

---

## Summary

| Question | Answer |
|----------|--------|
| Is template required? | **No**, but recommended |
| Can I delete the template? | **Yes**, system will create new presentation |
| What if I want only goal slides? | Delete template or set `use_template=False` |
| Can I use my own template? | **Yes**, change `template_path` in config |
| Will template be modified? | **No**, original template is never changed |
| Where is output saved? | `output/consolidated_mbr.pptx` |

---

## Example Commands

### With Template (Default)
```bash
py main.py --mbr-workflow
# Uses template + adds goal slides
```

### Without Template (Goals Only)
```bash
# Delete or move template first
move "sample_data\final-data\06-01-2026.pptx" "template_backup.pptx"

# Run workflow
py main.py --mbr-workflow
# Creates new presentation with only goal slides
```

---

**The system is flexible!** Use it with or without a template based on your needs. 🚀
