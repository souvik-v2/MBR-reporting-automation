"""
Extract graph data from LT-SimplifiedMetrics Excel files.
Extracts: V2 Carryover, LT Carryover, Rework, Reopened SP (LT)
"""
from __future__ import annotations
from typing import List, Dict, Any
import pandas as pd
from pathlib import Path
import re


def extract_project_name(filename: str) -> str:
    """Extract project name from filename like 'LT-SimplifiedMetrics-seod.xlsx'"""
    match = re.search(r'LT-SimplifiedMetrics-(.+)\.xlsx', filename)
    return match.group(1).upper() if match else filename


def read_graph_data(file_path: Path) -> Dict[str, Any]:
    """
    Read graph data from LT-SimplifiedMetrics Excel file.
    
    Returns dict with:
    - project: project name
    - week: week starting date
    - v2_carryover: V2 Carryover value
    - lt_carryover: LT Carryover value
    - rework: Rework value
    - reopened_sp_lt: Reopened SP (LT) value
    """
    # Read the file with row 1 as header (index 1, 0-based)
    df = pd.read_excel(file_path, header=1)
    
    # The first row after header contains column names
    if len(df) > 0:
        # Get column mappings from first row
        col_mapping = df.iloc[0].to_dict()
        
        # Find the actual data row (row 2, which is index 1)
        if len(df) > 1:
            data_row = df.iloc[1]
            
            # Map column names to values
            result = {
                'project': extract_project_name(file_path.name),
                'source_file': file_path.name,
            }
            
            # Find indices for columns we need
            for idx, col_name in col_mapping.items():
                if col_name == 'Week Starting On':
                    result['week'] = data_row[idx]
                elif col_name == 'V2 Carryover':
                    result['v2_carryover'] = data_row[idx]
                elif col_name == 'LT Carryover':
                    result['lt_carryover'] = data_row[idx]
                elif col_name == 'Rework':
                    result['rework'] = data_row[idx]
                elif col_name == 'Reopened SP (LT)':
                    result['reopened_sp_lt'] = data_row[idx]
            
            return result
    
    return {
        'project': extract_project_name(file_path.name),
        'source_file': file_path.name,
        'error': 'Could not parse data'
    }


def extract_all_graph_data(graph_folder: str = "sample_data/Graph") -> List[Dict[str, Any]]:
    """
    Extract graph data from all LT-SimplifiedMetrics files in the folder.
    
    Returns list of dicts, one per project.
    """
    graph_path = Path(graph_folder)
    results = []
    
    for file_path in graph_path.glob("LT-SimplifiedMetrics-*.xlsx"):
        # Skip temporary files
        if file_path.name.startswith('~$'):
            continue
        
        try:
            data = read_graph_data(file_path)
            results.append(data)
        except Exception as e:
            results.append({
                'project': extract_project_name(file_path.name),
                'source_file': file_path.name,
                'error': str(e)
            })
    
    return results


def create_graph_data_dataframe(graph_data: List[Dict[str, Any]]) -> pd.DataFrame:
    """Convert extracted graph data to a pandas DataFrame."""
    return pd.DataFrame(graph_data)


if __name__ == "__main__":
    # Test the extraction
    data = extract_all_graph_data()
    df = create_graph_data_dataframe(data)
    
    print("Extracted Graph Data:")
    print("=" * 80)
    print(df.to_string())
    
    # Show summary
    print("\n" + "=" * 80)
    print("Summary:")
    print(f"Total projects: {len(df)}")
    print(f"Columns extracted: {df.columns.tolist()}")
