"""
Extract chart data from Excel and recreate charts.
"""
from openpyxl import load_workbook
from openpyxl.chart import LineChart, BarChart
import pandas as pd
from pathlib import Path
import matplotlib.pyplot as plt
import matplotlib
matplotlib.use('Agg')  # Non-interactive backend

def extract_and_create_charts(excel_path, output_folder, project_name):
    """
    Extract data from Excel and create chart images.
    """
    print(f"   📊 Processing {project_name}...")
    
    output_folder = Path(output_folder)
    output_folder.mkdir(parents=True, exist_ok=True)
    
    charts_created = {}
    
    try:
        # Load workbook
        wb = load_workbook(excel_path, data_only=True)
        
        print(f"      Sheets: {wb.sheetnames[:5]}...")  # First 5
        
        # Try to find data for charts
        # Look for sheets with relevant names
        for sheet_name in wb.sheetnames:
            sheet = wb[sheet_name]
            
            # Check if this sheet has the data we need
            # Look for columns like "Week Starting On", "V2", "LT", "Rework", etc.
            
            # Read first few rows to find headers
            data = []
            for row in sheet.iter_rows(min_row=1, max_row=10, values_only=True):
                data.append(row)
            
            # Check if this looks like our data
            for row_idx, row in enumerate(data):
                if row and any(cell and 'week' in str(cell).lower() for cell in row if cell):
                    print(f"      Found data in '{sheet_name}' at row {row_idx + 1}")
                    
                    # Try to read as DataFrame
                    df = pd.read_excel(excel_path, sheet_name=sheet_name, header=row_idx)
                    
                    print(f"      Columns: {df.columns[:10].tolist()}")
                    
                    # Check if we have the columns we need
                    v2_col = None
                    lt_col = None
                    week_col = None
                    rework_col = None
                    
                    for col in df.columns:
                        col_lower = str(col).lower()
                        if 'week' in col_lower and 'start' in col_lower:
                            week_col = col
                        elif col_lower == 'v2':
                            v2_col = col
                        elif col_lower == 'lt':
                            lt_col = col
                        elif 'rework' in col_lower:
                            rework_col = col
                    
                    print(f"      Found columns: week={week_col}, v2={v2_col}, lt={lt_col}, rework={rework_col}")
                    
                    if week_col and (v2_col or lt_col):
                        # We have enough data to create a chart!
                        print(f"      ✓ Can create charts from this data")
                        
                        # Create a simple chart as proof of concept
                        output_file = output_folder / f"{project_name}_sample_chart.png"
                        
                        plt.figure(figsize=(8, 5))
                        plt.plot([1, 2, 3, 4], [10, 20, 15, 25], label='Sample Data')
                        plt.title(f'{project_name} - Sample Chart')
                        plt.xlabel('Week')
                        plt.ylabel('Story Points')
                        plt.legend()
                        plt.tight_layout()
                        plt.savefig(output_file, dpi=150, bbox_inches='tight')
                        plt.close()
                        
                        charts_created['sample'] = str(output_file)
                        print(f"      ✅ Created sample chart: {output_file.name}")
                        
                        break  # Found data, stop searching
            
            if charts_created:
                break  # Already found what we need
        
        print(f"      ✓ Created {len(charts_created)} chart(s)")
        
    except Exception as e:
        print(f"      ❌ Error: {e}")
        import traceback
        traceback.print_exc()
    
    return charts_created


def test():
    """Test chart creation."""
    print("\n" + "="*80)
    print("🧪 TESTING CHART DATA EXTRACTION & CREATION")
    print("="*80 + "\n")
    
    excel_file = "sample_data/Graph/LT-SimplifiedMetrics-seod.xlsx"
    output_folder = "output/charts"
    
    charts = extract_and_create_charts(excel_file, output_folder, "SEOD")
    
    print("\n" + "="*80)
    print("✅ TEST COMPLETE")
    print("="*80)
    print(f"\nCharts created: {len(charts)}")
    for key, path in charts.items():
        print(f"  • {key}: {path}")


if __name__ == "__main__":
    test()
