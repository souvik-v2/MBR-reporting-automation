# 🚀 MBR Automation - Quick Start Guide

## One Command to Rule Them All

```bash
py build_complete_mbr.py March
```

**That's it!** This single command creates your complete MBR report in seconds.

---

## 📋 Prerequisites

### Required Files:
1. **Template**: `sample_data/final-data/06-01-2026.pptx`
2. **Excel Files**: `sample_data/Graph/LT-SimplifiedMetrics-{project}.xlsx`
3. **Team Presentations**: `sample_data/ppt-data/V2Solutions_*.pptx`

### Required Packages:
```bash
py -m pip install python-pptx pandas openpyxl pywin32
```

---

## ⚡ Step-by-Step Usage

### Step 1: Run the Script
```bash
py build_complete_mbr.py March
```

**What it does**:
- ✅ Updates Slide 1 date to "01-March-2026"
- ✅ Populates Slide 3 table with 6 projects
- ✅ Creates 6 Project Metrics slides
- ✅ Adds achievements from team presentations
- ✅ Adds chart placeholders with instructions

**Time**: 5-10 seconds

### Step 2: Add Charts (12 minutes)

For each Project Metrics slide (8-13):

1. **Read the placeholder**
   ```
   [CHART: Rework Trend]
   Excel: LT-SimplifiedMetrics-seod.xlsx
   Chart: 'Rework Trend'
   ```

2. **Open the Excel file**
   - Navigate to `sample_data/Graph/`
   - Open the specified file

3. **Copy the chart**
   - Find "Rework Trend" chart
   - Right-click → Copy

4. **Paste into PowerPoint**
   - Click on the placeholder area
   - Paste (Ctrl+V)
   - Delete the placeholder textbox
   - Resize chart to fit

5. **Repeat for second chart**
   - "Throughput Distribution"

### Step 3: Review & Present
- Open `output/consolidated_mbr.pptx`
- Review all slides
- Make any final adjustments
- Present!

---

## 📊 What You Get

### Automated Slides:
- **Slide 1**: Title with March date ✅
- **Slide 2**: LT–V2 Joint Delivery Snapshot ✅
- **Slide 3**: Project table with:
  - 6 project names ✅
  - LT Managers ✅
  - Completed goals ✅
  - In-progress items ✅
- **Slides 4-6**: Template content ✅
- **Slide 7**: Thank You ✅
- **Slides 8-13**: Project Metrics (one per project) ✅

### Projects Included:
1. CANOPYNEXT
2. CONSUMERWEB
3. FLEX
4. PHOENIX
5. SEOD
6. WPM

---

## 🎯 Time Savings

| Task | Manual | Automated | Saved |
|------|--------|-----------|-------|
| Find project names | 10 min | 0 sec | 10 min |
| Update date | 2 min | 0 sec | 2 min |
| Populate table | 30 min | 0 sec | 30 min |
| Create metrics slides | 45 min | 0 sec | 45 min |
| Extract achievements | 60 min | 0 sec | 60 min |
| Copy charts | 12 min | 12 min | 0 min |
| **TOTAL** | **~2.5 hrs** | **~15 min** | **~2 hrs 15 min** |

**90% time reduction!** 🎉

---

## 🔧 Troubleshooting

### Problem: Script fails to find projects
**Solution**: Check that Excel files exist in `sample_data/Graph/` with pattern `LT-SimplifiedMetrics-*.xlsx`

### Problem: No achievements shown
**Solution**: Verify team presentations exist in `sample_data/ppt-data/` and contain slides titled "Goals Achieved in March"

### Problem: Table cells are empty
**Solution**: 
- Check LT Managers exist in template Slide 3
- Verify team PPT files have the correct slide titles

### Problem: pywin32 errors
**Solution**: Run `py -m pip install --upgrade pywin32`

---

## 💡 Tips & Best Practices

### For Different Months:
```bash
# For April report:
py build_complete_mbr.py April

# For May report:
py build_complete_mbr.py May
```

### To Check Output Before Presenting:
```bash
# View slide structure:
py check_ppt.py

# Inspect Slide 3 table:
py inspect_slide3.py
```

### Adding New Projects:
1. Add `LT-SimplifiedMetrics-{newproject}.xlsx` to `sample_data/Graph/`
2. Add team presentation to `sample_data/ppt-data/`
3. Run the script - it auto-detects new projects!

---

## 📞 Support & Customization

### Want to customize?
- Edit `build_complete_mbr.py` for different slide layouts
- Modify `create_mbr_complete.py` for different data extraction
- Update template path in scripts if using different template

### Need help?
- Check `AUTOMATION_SUMMARY.md` for detailed status
- See `FINAL_MBR_AUTOMATION_README.md` for architecture
- Review `CONTINUE_PLAN.md` for future enhancements

---

## 🎊 Success Story

**Before automation**:
- 2-3 hours per report
- Manual data copying
- Error-prone
- Tedious

**After automation**:
- 15 minutes per report  
- Automatic data extraction
- Consistent format
- Easy to scale

**Result**: Manager happy, you happy! 😊

---

## 🚀 Ready? Let's Go!

```bash
py build_complete_mbr.py March
```

Then spend 12 minutes copying charts, and you're done! 🎉
