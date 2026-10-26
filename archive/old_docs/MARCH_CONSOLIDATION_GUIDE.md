# March 2026 MBR Consolidated Report Guide

## What Was Found

The system scanned all team presentations and found **7 slides** to consolidate:

### Goals Achieved in March (5 slides):
1. **LT_CanopyNext** - Slide 3: Goals Achieved in March
2. **LT_Core_SVC** - Slide 3: Goals Achieved in March  
3. **LT_FLEX** - Slide 3: Goals Achieved in March
4. **LT_UALS** - Slide 3: Goals Achieved
5. **SEOD** - Slide 3: Goals Achieved in March

### In Progress Ticket (2 slides):
6. **LT_Core_SVC** - Slide 5: In Progress Ticket
7. **SEOD** - Slide 7: In Progress Ticket

---

## How to Create the Consolidated PowerPoint

### Step 1: Close ALL PowerPoint Files

**IMPORTANT:** Close any open PowerPoint presentations, including:
- `V2Solutions_SEOD_March_2026_Report.pptx`
- Any other team presentations
- Any previously created consolidated files

### Step 2: Run the Consolidation Script

```bash
py create_consolidated_mbr.py March
```

### Step 3: Check the Output

The consolidated presentation will be saved as:
```
output/consolidated_mbr.pptx
```

---

## What Will Be in the Consolidated Report

### Slide Structure:

```
Slide 1: Goals Achieved in March (CanopyNext)
Slide 2: Goals Achieved in March (Core SVC)
Slide 3: In Progress Ticket (Core SVC)
Slide 4: Goals Achieved in March (FLEX)
Slide 5: Goals Achieved (UALS)
Slide 6: Goals Achieved in March (SEOD)
Slide 7: In Progress Ticket (SEOD)
```

**Total: 7 slides**

---

## SEOD Content

Based on what you shared, the SEOD slides will include:

### Slide 6: Goals Achieved in March (SEOD)

**Topics:**
- Typography & UI Enhancements
- Mobile Rendering & Responsiveness Fixes
- Shortcode Development & Fixes
- Accessibility (A11y) Improvements
- Table UX & Interaction Improvements
- SEO & Indexing Improvements
- Bug Fixes & Stability Improvements

### Slide 7: In Progress Ticket (SEOD)

**Content:**
- VP - Home insurance calculator (SEOD-1922)
- [Any other in-progress items from the slide]

---

## Key Features

✅ **No Template Required** - Creates presentation from scratch
✅ **Extracts Goals** - "Goals Achieved in March" slides
✅ **Extracts In-Progress** - "In Progress Ticket" slides
✅ **Preserves Formatting** - Uses COM automation for perfect fidelity
✅ **All Teams Included** - Automatically scans all presentations

---

## Troubleshooting

### Error: "Access Denied" or "Exception occurred"

**Cause:** PowerPoint files are open or locked

**Solution:**
1. Close ALL PowerPoint windows
2. Check Task Manager for any PowerPoint processes and end them
3. Run the script again

### Error: "No slides found to consolidate"

**Cause:** Month doesn't match slide titles

**Solution:**
- Check that presentations have "Goals Achieved in March" slides
- Verify the month parameter: `py create_consolidated_mbr.py March`

### Files Not Being Copied

**Cause:** File permissions or PowerPoint is running

**Solution:**
1. Run PowerPoint as administrator once to reset permissions
2. Close PowerPoint completely
3. Restart your computer if needed
4. Run the script again

---

##For Different Months

To create consolidated reports for other months:

```bash
# For April
py create_consolidated_mbr.py April

# For May
py create_consolidated_mbr.py May

# For June
py create_consolidated_mbr.py June
```

The script will search for:
- "Goals Achieved in {Month}"
- "In Progress Ticket"

---

## Next Steps After Creation

1. **Open the file:**
   ```
   output/consolidated_mbr.pptx
   ```

2. **Verify content:**
   - Check that all 7 slides are present
   - Verify SEOD slides have the correct content
   - Confirm In-Progress tickets are included

3. **Optional - Add title slide:**
   - Create a cover slide manually
   - Add "March 2026 MBR Report"
   - Insert at the beginning

4. **Optional - Add closing slide:**
   - Add "Thank You" or "Next Steps" slide
   - Insert at the end

---

## Summary

| Item | Details |
|------|---------|
| **Command** | `py create_consolidated_mbr.py March` |
| **Input Folder** | `sample_data/ppt-data/` |
| **Output File** | `output/consolidated_mbr.pptx` |
| **Total Slides** | 7 (5 Goals + 2 In-Progress) |
| **Teams Included** | CanopyNext, Core SVC, FLEX, UALS, SEOD |
| **Template** | None (creates from scratch) |

---

**The consolidated report is ready to use! No manual copying needed.** 🚀
