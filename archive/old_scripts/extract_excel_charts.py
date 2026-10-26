"""
Extract charts from Excel files as images.

Uses WIN32COM to export Excel charts as PNG images.
"""
import sys
import os
from pathlib import Path
import time

try:
    import win32com.client
    from win32com.client import constants
except ImportError:
    print("ERROR: pywin32 required. Install with: py -m pip install pywin32")
    sys.exit(1)


def extract_charts_from_excel_com(excel_path, output_folder, project_name):
    """
    Extract Rework Trend and Throughput Distribution charts from Excel.
    
    Returns:
        dict: {'rework_trend': 'path/to/image.png', 'throughput': 'path/to/image.png'}
    """
    print(f"   📊 Extracting charts for {project_name}...")
    
    excel_path = str(Path(excel_path).absolute())
    output_folder = Path(output_folder)
    output_folder.mkdir(parents=True, exist_ok=True)
    
    charts_found = {}
    
    excel = None
    wb = None

    try:
        # Start Excel - don't set any properties that might fail
        excel = win32com.client.Dispatch("Excel.Application")

        # Open workbook (read-only to avoid issues)
        wb = excel.Workbooks.Open(excel_path, ReadOnly=True)

        print(f"      Opened workbook with {wb.Sheets.Count} sheets")
        
        # Search all sheets for charts
        for sheet in wb.Sheets:
            try:
                chart_objects = sheet.ChartObjects()
                
                if chart_objects.Count > 0:
                    print(f"      Sheet '{sheet.Name}' has {chart_objects.Count} chart(s)")
                    
                    for idx in range(1, chart_objects.Count + 1):
                        chart_obj = chart_objects(idx)
                        chart = chart_obj.Chart
                        
                        # Get chart title
                        chart_title = ""
                        if chart.HasTitle:
                            chart_title = str(chart.ChartTitle.Text).lower()
                        
                        print(f"         Chart {idx}: {chart_title}")
                        
                        # Identify chart type
                        chart_type = None
                        output_filename = None
                        
                        if 'rework' in chart_title or 'trend' in chart_title:
                            chart_type = 'rework_trend'
                            output_filename = f"{project_name}_rework_trend.png"
                        elif 'throughput' in chart_title or 'distribution' in chart_title:
                            chart_type = 'throughput'
                            output_filename = f"{project_name}_throughput.png"
                        
                        # Export chart as image
                        if chart_type and output_filename:
                            output_path = output_folder / output_filename
                            
                            # Export chart
                            chart.Export(str(output_path.absolute()))
                            
                            charts_found[chart_type] = str(output_path)
                            print(f"         ✅ Exported: {output_filename}")
                            
                            time.sleep(0.2)
                            
            except Exception as e:
                print(f"      ⚠️  Error processing sheet '{sheet.Name}': {e}")
                continue
        
        print(f"      ✓ Found {len(charts_found)} chart(s): {list(charts_found.keys())}")

    except Exception as e:
        print(f"      ❌ Error: {e}")
        import traceback
        traceback.print_exc()

    finally:
        # Cleanup
        try:
            if wb:
                wb.Close(SaveChanges=False)
            if excel:
                excel.Quit()
        except:
            pass
    
    return charts_found


def test_extraction():
    """Test chart extraction."""
    print("\n" + "="*80)
    print("🧪 TESTING CHART EXTRACTION")
    print("="*80 + "\n")
    
    excel_file = "sample_data/Graph/LT-SimplifiedMetrics-seod.xlsx"
    output_folder = "output/charts"
    
    if not Path(excel_file).exists():
        print(f"❌ File not found: {excel_file}")
        return
    
    print(f"📂 Excel file: {excel_file}")
    print(f"📁 Output folder: {output_folder}\n")
    
    charts = extract_charts_from_excel_com(excel_file, output_folder, "SEOD")
    
    print("\n" + "="*80)
    print("✅ TEST COMPLETE")
    print("="*80)
    print(f"\nCharts extracted: {len(charts)}")
    for chart_type, path in charts.items():
        print(f"  • {chart_type}: {path}")
    
    if charts:
        print(f"\n📁 Check folder: {output_folder}")
    else:
        print("\n⚠️  No charts found. Check Excel file structure.")


if __name__ == "__main__":
    test_extraction()
