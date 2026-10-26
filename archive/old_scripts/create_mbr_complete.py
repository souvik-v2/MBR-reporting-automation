"""
Complete MBR Report Builder - March 2026

This script creates a fully formatted MBR report by:
1. Updating Slide 1 with March date
2. Populating Slide 3 table with project data
3. Creating Project Metrics slides (7+) with charts from Excel
4. Adding Key Achievements from presentations

Usage:
    py create_mbr_complete.py March
"""
import sys
import re
from pathlib import Path
from datetime import datetime
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
import pandas as pd
from copy import deepcopy

# Configuration
MONTH = sys.argv[1] if len(sys.argv) > 1 else "March"
TEMPLATE_PATH = "sample_data/final-data/06-01-2026.pptx"
GRAPH_FOLDER = "sample_data/Graph"
PPT_FOLDER = "sample_data/ppt-data"
OUTPUT_PATH = "output/consolidated_mbr.pptx"

def log(message):
    """Print log message."""
    print(message)

def get_all_projects():
    """Get list of all projects from Graph folder."""
    graph_path = Path(GRAPH_FOLDER)
    excel_files = list(graph_path.glob("LT-SimplifiedMetrics-*.xlsx"))
    
    projects = []
    for file in excel_files:
        match = re.search(r'LT-SimplifiedMetrics-(.+)\.xlsx', file.name, re.IGNORECASE)
        if match:
            project_name = match.group(1).upper()
            projects.append({
                'name': project_name,
                'excel_path': str(file),
                'display_name': project_name.replace('_', ' ').title()
            })
    
    return sorted(projects, key=lambda x: x['name'])

def get_project_manager_from_excel(excel_path, project_name):
    """Extract Project Manager/Lead from Excel file."""
    try:
        # The Excel has multiple sheets, find the one for this project
        xl = pd.ExcelFile(excel_path)
        
        # Look for sheet matching project name
        for sheet in xl.sheet_names:
            if project_name.lower() in sheet.lower():
                df = pd.read_excel(excel_path, sheet_name=sheet, nrows=10)
                
                # Look for "Project Manager" or "Lead" in the sheet
                for col in df.columns:
                    if 'manager' in str(col).lower() or 'lead' in str(col).lower():
                        val = df[col].dropna().iloc[0] if len(df[col].dropna()) > 0 else None
                        if val and isinstance(val, str):
                            return val
        
    except Exception as e:
        log(f"      ⚠️  Could not read manager from Excel: {e}")
    
    return "TBD"

def get_lt_manager_from_template_slide3(template_slide, project_name):
    """Extract LT Manager from template Slide 3 table."""
    try:
        for shape in template_slide.shapes:
            if shape.has_table:
                table = shape.table
                
                # Search rows for project name
                for row_idx in range(1, len(table.rows)):
                    proj_cell = table.cell(row_idx, 0).text.strip().upper()
                    if project_name.upper() in proj_cell or proj_cell in project_name.upper():
                        lt_mgr = table.cell(row_idx, 1).text.strip()
                        return lt_mgr if lt_mgr else "TBD"
    except:
        pass
    
    return "TBD"

def get_presentation_for_project(project_name):
    """Find PowerPoint presentation file for project."""
    ppt_path = Path(PPT_FOLDER)
    
    # Search patterns
    patterns = [
        f"*{project_name}*March*.pptx",
        f"*{project_name.replace('NEXT', 'Next')}*March*.pptx",
        f"*{project_name[:4]}*March*.pptx",
    ]
    
    for pattern in patterns:
        files = list(ppt_path.glob(pattern))
        files = [f for f in files if not f.name.startswith('~$')]
        if files:
            return str(files[0])
    
    return None

def extract_headings_from_slide(prs, slide_title_keyword):
    """Extract bullet point headings from a slide."""
    headings = []
    
    try:
        for slide in prs.slides:
            if slide.shapes.title and slide_title_keyword.lower() in slide.shapes.title.text.lower():
                # Found the slide, extract headings
                for shape in slide.shapes:
                    if shape.has_text_frame and shape != slide.shapes.title:
                        for paragraph in shape.text_frame.paragraphs:
                            text = paragraph.text.strip()
                            # Filter out empty and very long text
                            if text and 10 < len(text) < 150:
                                # Check if it looks like a heading (starts with capital, no ending punctuation patterns)
                                if text[0].isupper() and not text.endswith(':'):
                                    headings.append(text)
                
                break
    except Exception as e:
        log(f"      ⚠️  Error extracting headings: {e}")
    
    return headings[:5]  # Return max 5 headings

def get_goals_and_progress_for_project(project_name):
    """Get completed goals and in-progress items for a project."""
    ppt_file = get_presentation_for_project(project_name)
    
    if not ppt_file:
        return [], []
    
    try:
        prs = Presentation(ppt_file)
        
        goals = extract_headings_from_slide(prs, "Goals Achieved")
        progress = extract_headings_from_slide(prs, "In Progress")
        
        return goals, progress
        
    except Exception as e:
        log(f"      ⚠️  Error reading presentation: {e}")
        return [], []

def update_slide_1_date(slide, month):
    """Update date in Slide 1."""
    for shape in slide.shapes:
        if shape.has_text_frame:
            for paragraph in shape.text_frame.paragraphs:
                if paragraph.text:
                    # Replace any month name with our month
                    new_text = re.sub(
                        r'(January|February|March|April|May|June|July|August|September|October|November|December)',
                        month,
                        paragraph.text,
                        flags=re.IGNORECASE
                    )
                    if new_text != paragraph.text:
                        paragraph.text = new_text
                        log(f"   ✓ Updated date: ...{new_text[:40]}...")

def update_slide_3_table(slide, projects, template_prs):
    """Update Slide 3 project table."""
    for shape in slide.shapes:
        if shape.has_table:
            table = shape.table

            template_slide3 = template_prs.slides[2]

            # Update each row with project data
            for idx, project in enumerate(projects[:5]):  # Max 5 rows in table
                row_idx = idx + 1

                if row_idx >= len(table.rows):
                    break

                log(f"   📊 {project['name']}...")

                # Column 0: Project Name
                table.cell(row_idx, 0).text = project['display_name']

                # Column 1: LT Manager
                lt_mgr = get_lt_manager_from_template_slide3(template_slide3, project['name'])
                table.cell(row_idx, 1).text = lt_mgr
                log(f"      LT Manager: {lt_mgr}")

                # Column 2: V2 Team
                v2_team = get_project_manager_from_excel(project['excel_path'], project['name'])
                table.cell(row_idx, 2).text = v2_team
                log(f"      V2 Team: {v2_team}")

                # Column 3: What We Delivered / Are Solving
                goals, progress = get_goals_and_progress_for_project(project['name'])

                completed_text = "Completed:\n" + "\n".join([f"• {g}" for g in goals[:3]]) if goals else "Completed: N/A"
                progress_text = "\nIn-progress:\n" + "\n".join([f"• {p}" for p in progress[:3]]) if progress else "\nIn-progress: N/A"

                table.cell(row_idx, 3).text = completed_text + progress_text
                log(f"      Goals: {len(goals)}, Progress: {len(progress)}")

def main():
    """Main execution."""
    log("\n" + "="*80)
    log(f"🎯 BUILDING {MONTH.upper()} 2026 MBR REPORT")
    log("="*80 + "\n")

    # Step 1: Load template
    log("📋 Step 1: Loading template...")
    template_prs = Presentation(TEMPLATE_PATH)
    log(f"   Loaded {len(template_prs.slides)} slides\n")

    # Step 2: Get projects
    log("📋 Step 2: Finding projects...")
    projects = get_all_projects()
    log(f"   Found {len(projects)} projects:")
    for p in projects:
        log(f"      • {p['name']}")
    log("")

    # Step 3: Update Slide 1 (Date)
    log("📋 Step 3: Updating Slide 1 date...")
    update_slide_1_date(template_prs.slides[0], MONTH)
    log("")

    # Step 4: Update Slide 3 (Table)
    log("📋 Step 4: Updating Slide 3 table...")
    update_slide_3_table(template_prs.slides[2], projects, template_prs)
    log("")

    # Step 5: Save
    log("💾 Step 5: Saving...")
    template_prs.save(OUTPUT_PATH)

    log("\n" + "="*80)
    log("✅ SUCCESS!")
    log("="*80)
    log(f"\n📁 Saved: {OUTPUT_PATH}")
    log(f"📊 Total slides: {len(template_prs.slides)}\n")
    log("NOTE: This version updates Slide 1 and Slide 3.")
    log("      Charts and Project Metrics slides (7+) will be added in next iteration.\n")

if __name__ == "__main__":
    main()
