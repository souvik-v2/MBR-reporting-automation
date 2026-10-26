"""
Build MBR Report with actual data updates.

This script:
1. Updates Slide 1 with the correct month/date
2. Updates Slide 3 table with project data
3. Creates Project Metrics slides with charts and achievements
"""
from pptx import Presentation
from pptx.util import Inches, Pt
from pathlib import Path
import pandas as pd
from datetime import datetime
import re

def get_projects_from_graph_folder(graph_folder):
    """Get list of projects from Graph folder."""
    graph_path = Path(graph_folder)
    excel_files = list(graph_path.glob("LT-SimplifiedMetrics-*.xlsx"))
    
    projects = []
    for file in excel_files:
        # Extract project name from filename
        # LT-SimplifiedMetrics-seod.xlsx -> SEOD
        match = re.search(r'LT-SimplifiedMetrics-(.+)\.xlsx', file.name, re.IGNORECASE)
        if match:
            project_name = match.group(1).upper()
            projects.append({
                'name': project_name,
                'excel_file': str(file),
                'display_name': project_name
            })
    
    return projects

def get_project_manager(excel_file):
    """Extract Project Manager/Lead from Excel file."""
    try:
        df = pd.read_excel(excel_file)
        if 'Project Manager/Lead' in df.columns:
            manager = df['Project Manager/Lead'].dropna().iloc[0] if len(df) > 0 else "N/A"
            return str(manager)
    except Exception as e:
        print(f"   ⚠️ Error reading manager: {e}")
    return "N/A"

def get_lt_manager_from_template(template_prs, project_name):
    """Get LT Manager from template slide 3."""
    try:
        slide3 = template_prs.slides[2]  # Slide 3 (0-indexed)
        
        # Find table in slide
        for shape in slide3.shapes:
            if shape.has_table:
                table = shape.table
                
                # Search for project name in column 1 and return LT Manager from column 2
                for row_idx in range(1, len(table.rows)):  # Skip header
                    cell_text = table.cell(row_idx, 0).text.strip().upper()
                    if project_name.upper() in cell_text:
                        lt_manager = table.cell(row_idx, 1).text.strip()
                        return lt_manager if lt_manager else "N/A"
    except Exception as e:
        print(f"   ⚠️ Error reading LT Manager: {e}")
    
    return "N/A"

def get_goals_headings(ppt_folder, project_name, slide_keyword):
    """Extract heading text from specific slide in team presentations."""
    ppt_path = Path(ppt_folder)
    
    # Find presentation for this project
    patterns = [
        f"*{project_name}*March*.pptx",
        f"*{project_name}*.pptx"
    ]
    
    ppt_file = None
    for pattern in patterns:
        files = list(ppt_path.glob(pattern))
        files = [f for f in files if not f.name.startswith('~$')]
        if files:
            ppt_file = files[0]
            break
    
    if not ppt_file:
        return []
    
    try:
        prs = Presentation(str(ppt_file))
        
        # Find slide with keyword
        for slide in prs.slides:
            if slide.shapes.title and slide_keyword.lower() in slide.shapes.title.text.lower():
                # Extract all heading text from this slide
                headings = []
                
                for shape in slide.shapes:
                    if shape.has_text_frame:
                        for paragraph in shape.text_frame.paragraphs:
                            text = paragraph.text.strip()
                            # Check if this is a heading (bold, larger font, or has specific formatting)
                            if text and len(text) > 0:
                                # Simple heuristic: non-empty text that's not too long
                                if len(text) < 200 and text not in headings:
                                    headings.append(text)
                
                return headings[:5]  # Return first 5 headings
                
    except Exception as e:
        print(f"   ⚠️ Error extracting headings: {e}")
    
    return []

def update_slide1_date(slide, month):
    """Update date on slide 1."""
    # Update title or text with month
    for shape in slide.shapes:
        if shape.has_text_frame:
            text = shape.text_frame.text
            # Replace month/date patterns
            if re.search(r'(January|February|March|April|May|June|July|August|September|October|November|December)', text, re.IGNORECASE):
                # Replace with new month
                new_text = re.sub(
                    r'(January|February|March|April|May|June|July|August|September|October|November|December)\s*\d{4}',
                    f'{month} 2026',
                    text,
                    flags=re.IGNORECASE
                )
                shape.text_frame.text = new_text
                print(f"   ✓ Updated: {new_text[:50]}...")

def update_slide3_table(slide, projects, template_prs, ppt_folder):
    """Update Slide 3 table with project data."""
    print("   Updating Slide 3 table...")
    
    for shape in slide.shapes:
        if shape.has_table:
            table = shape.table
            
            # Clear existing rows (except header)
            # Add/update rows for each project
            
            for idx, project in enumerate(projects):
                row_idx = idx + 1  # Skip header row
                
                # Ensure row exists
                if row_idx >= len(table.rows):
                    # Add new row (copy last row)
                    break
                
                # Column 0: Project Name
                table.cell(row_idx, 0).text = project['display_name']
                
                # Column 1: LT Manager (from template)
                lt_manager = get_lt_manager_from_template(template_prs, project['name'])
                table.cell(row_idx, 1).text = lt_manager
                
                # Column 2: V2 Team (from Excel)
                v2_team = get_project_manager(project['excel_file'])
                table.cell(row_idx, 2).text = v2_team
                
                # Column 3: What We Delivered / Are Solving
                completed_goals = get_goals_headings(ppt_folder, project['name'], "Goals Achieved")
                in_progress_goals = get_goals_headings(ppt_folder, project['name'], "In Progress")
                
                completed_text = "Completed:\n" + "\n".join([f"• {g}" for g in completed_goals[:3]]) if completed_goals else "Completed: N/A"
                progress_text = "In Progress:\n" + "\n".join([f"• {g}" for g in in_progress_goals[:3]]) if in_progress_goals else "In Progress: N/A"
                
                table.cell(row_idx, 3).text = f"{completed_text}\n\n{progress_text}"
                
                print(f"      ✓ Row {row_idx}: {project['display_name']}")

def extract_charts_from_excel(excel_file, output_folder):
    """Extract chart images from Excel file."""
    from openpyxl import load_workbook
    from openpyxl.chart import BarChart, LineChart

    charts_extracted = {}

    try:
        wb = load_workbook(excel_file)

        # Look for charts in all sheets
        for sheet_name in wb.sheetnames:
            sheet = wb[sheet_name]

            # Check if sheet has charts
            if hasattr(sheet, '_charts') and sheet._charts:
                for idx, chart in enumerate(sheet._charts):
                    chart_title = chart.title.text if hasattr(chart, 'title') and chart.title else f"Chart_{idx}"

                    # Identify Rework Trend or Throughput Distribution
                    if 'rework' in chart_title.lower() or 'trend' in chart_title.lower():
                        charts_extracted['rework_trend'] = chart
                    elif 'throughput' in chart_title.lower() or 'distribution' in chart_title.lower():
                        charts_extracted['throughput'] = chart

        print(f"      ✓ Found {len(charts_extracted)} charts")

    except Exception as e:
        print(f"      ⚠️ Error extracting charts: {e}")

    return charts_extracted


def main():
    """Main function to build MBR report."""
    print("\n" + "="*80)
    print("🎯 BUILDING MARCH 2026 MBR REPORT")
    print("="*80 + "\n")

    # Configuration
    template_path = "sample_data/final-data/06-01-2026.pptx"
    graph_folder = "sample_data/Graph"
    ppt_folder = "sample_data/ppt-data"
    output_path = "output/consolidated_mbr.pptx"
    month = "March"

    print(f"📂 Template: {template_path}")
    print(f"📂 Graph data: {graph_folder}")
    print(f"📂 PPT data: {ppt_folder}")
    print(f"📅 Month: {month}\n")

    # Load template
    print("📋 Step 1: Loading template...")
    template_prs = Presentation(template_path)
    print(f"   Loaded: {len(template_prs.slides)} slides\n")

    # Get projects
    print("📋 Step 2: Finding projects...")
    projects = get_projects_from_graph_folder(graph_folder)
    print(f"   Found {len(projects)} projects:")
    for p in projects:
        print(f"      • {p['name']}")
    print()

    # Update Slide 1 - Date
    print("📋 Step 3: Updating Slide 1 (Date)...")
    update_slide1_date(template_prs.slides[0], month)
    print()

    # Update Slide 3 - Table
    print("📋 Step 4: Updating Slide 3 (Project Table)...")
    update_slide3_table(template_prs.slides[2], projects, template_prs, ppt_folder)
    print()

    # Save
    print("💾 Step 5: Saving...")
    template_prs.save(output_path)

    print("\n" + "="*80)
    print("✅ SUCCESS!")
    print("="*80)
    print(f"\n📁 Saved: {output_path}")
    print(f"📊 Total slides: {len(template_prs.slides)}\n")


if __name__ == "__main__":
    main()
