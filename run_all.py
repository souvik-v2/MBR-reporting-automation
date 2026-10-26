"""
run_all.py -- generate every output (PowerPoint, Excel, dashboard) in one
command, fully auto-detected from sample_data/.

Re-run this exact command any time files in sample_data/ change (new
team deck, updated Graph workbook, new month's template) -- nothing
needs to be edited first.
"""
import subprocess
import sys

print("=" * 78)
print("STEP 1/2: Excel master workbook + HTML dashboard + MBR Excel report")
print("=" * 78)
r1 = subprocess.run([sys.executable, "main.py"])

print()
print("=" * 78)
print("STEP 2/2: MBR PowerPoint deck")
print("=" * 78)
r2 = subprocess.run([sys.executable, "build_mbr.py"])

print()
print("=" * 78)
if r1.returncode == 0 and r2.returncode == 0:
    print("Done. All outputs are in the output/ folder:")
    print("  - master_file.xlsx        (merged data from every file in sample_data/)")
    print("  - dashboard.html          (interactive view of the same data)")
    print("  - mbr_report.xlsx         (Sprint KPIs, trends, goals)")
    print("  - MBR_<Month>_<Year>.pptx (the finished MBR slide deck)")
    print("  - MBR_<Month>_<Year>.pdf  (PDF copy of the same deck)")
    print("  - MBR_..._REVIEW_CHECKLIST.md (anything worth a quick human look)")
    print()
    print("To ask questions about this data instead, run: python chatbot_app.py")
else:
    print("One or both steps failed -- see the output above for details.")
    sys.exit(1)
