"""
Export charts from Excel files as PNG images.
Works with pywin32 on Windows.
"""
import os
import sys
import time
from pathlib import Path

def export_charts_from_excel(excel_path, output_folder, project_name):
    """
    Export Rework Trend and Throughput Distribution charts as PNG files.
    
    Returns:
        dict: {'rework_trend': 'path.png', 'throughput': 'path.png'}
    """
    try:
        import pythoncom
        import win32com.client as win32
        
        pythoncom.CoInitialize()
        
        excel_path = str(Path(excel_path).absolute())
        output_folder = Path(output_folder)
        output_folder.mkdir(parents=True, exist_ok=True)
        
        print(f"   📊 Exporting charts from {project_name}...")
        
        charts_exported = {}
        excel = None
        wb = None
        
        try:
            # Create Excel instance
            excel = win32.Dispatch('Excel.Application')
            
            # Open workbook
            wb = excel.Workbooks.Open(excel_path, ReadOnly=1)
            
            print(f"      Opened {wb.Sheets.Count} sheets")
            
            # Search all sheets
            for sheet in wb.Sheets:
                try:
                    if sheet.ChartObjects().Count > 0:
                        print(f"      Sheet '{sheet.Name}': {sheet.ChartObjects().Count} chart(s)")
                        
                        for i in range(1, sheet.ChartObjects().Count + 1):
                            chart_obj = sheet.ChartObjects(i)
                            chart = chart_obj.Chart
                            
                            # Get title
                            title = ""
                            try:
                                if chart.HasTitle:
                                    title = chart.ChartTitle.Text.lower()
                            except:
                                pass
                            
                            print(f"         Chart {i}: '{title}'")
                            
                            # Determine chart type
                            output_file = None
                            chart_key = None
                            
                            if 'rework' in title or 'trend' in title:
                                chart_key = 'rework_trend'
                                output_file = output_folder / f"{project_name}_rework_trend.png"
                            elif 'throughput' in title or 'distribution' in title:
                                chart_key = 'throughput'
                                output_file = output_folder / f"{project_name}_throughput.png"
                            
                            # Export
                            if output_file and chart_key:
                                try:
                                    chart.Export(str(output_file.absolute()), "PNG")
                                    charts_exported[chart_key] = str(output_file)
                                    print(f"         ✅ Exported to {output_file.name}")
                                    time.sleep(0.1)
                                except Exception as e:
                                    print(f"         ⚠️ Export failed: {e}")
                                    
                except Exception as e:
                    print(f"      ⚠️ Error in sheet: {e}")
                    continue
            
            print(f"      ✓ Exported {len(charts_exported)} chart(s)")
            
        finally:
            # Cleanup
            if wb:
                wb.Close(SaveChanges=False)
            if excel:
                excel.Quit()
            
            pythoncom.CoUninitialize()
        
        return charts_exported
        
    except ImportError:
        print("      ❌ pywin32 not installed. Install with: py -m pip install pywin32")
        return {}
    except Exception as e:
        print(f"      ❌ Error: {e}")
        import traceback
        traceback.print_exc()
        return {}


def test_export():
    """Test chart export."""
    print("\n" + "="*80)
    print("🧪 TESTING CHART EXPORT")
    print("="*80 + "\n")
    
    excel_file = "sample_data/Graph/LT-SimplifiedMetrics-seod.xlsx"
    output_folder = "output/charts"
    
    if not Path(excel_file).exists():
        print(f"❌ File not found: {excel_file}")
        return
    
    charts = export_charts_from_excel(excel_file, output_folder, "SEOD")
    
    print("\n" + "="*80)
    print("✅ TEST COMPLETE")
    print("="*80)
    print(f"\nCharts exported: {len(charts)}")
    for key, path in charts.items():
        print(f"  • {key}: {path}")
        if Path(path).exists():
            size = Path(path).stat().st_size
            print(f"    File size: {size:,} bytes")


if __name__ == "__main__":
    test_export()
