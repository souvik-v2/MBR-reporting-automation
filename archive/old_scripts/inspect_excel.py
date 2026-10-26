"""Inspect Excel files to understand structure"""
import pandas as pd
from pathlib import Path

excel_file = "sample_data/Graph/LT-SimplifiedMetrics-seod.xlsx"

print(f"Inspecting: {excel_file}\n")
print("="*80)

# Read Excel to see sheets
xl = pd.ExcelFile(excel_file)
print(f"\nSheets in file: {xl.sheet_names}\n")

# Check first sheet
for sheet_name in xl.sheet_names[:2]:  # First 2 sheets
    print(f"\n{'='*80}")
    print(f"Sheet: {sheet_name}")
    print('='*80)
    
    df = pd.read_excel(excel_file, sheet_name=sheet_name, nrows=5)
    
    print(f"\nColumns: {df.columns.tolist()}\n")
    print("First 3 rows:")
    print(df.head(3).to_string())
    print()

# Also check if there are charts
from openpyxl import load_workbook

wb = load_workbook(excel_file)
print("\n" + "="*80)
print("Checking for charts...")
print("="*80 + "\n")

for sheet_name in wb.sheetnames:
    sheet = wb[sheet_name]
    if hasattr(sheet, '_charts') and sheet._charts:
        print(f"Sheet '{sheet_name}' has {len(sheet._charts)} chart(s)")
        for idx, chart in enumerate(sheet._charts):
            chart_title = chart.title.text if hasattr(chart, 'title') and chart.title and hasattr(chart.title, 'text') else "No title"
            print(f"  Chart {idx+1}: {chart_title}")
