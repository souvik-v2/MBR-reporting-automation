"""
COMPLETE MBR Report Builder - Fully Automated

Creates March MBR with:
1. Slide 1: Updated date
2. Slide 3: Project table with data
3. Slides 7+: Project Metrics slides (one per project)

Usage:
    py build_complete_mbr.py March
"""
import sys
import re
from pathlib import Path
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
import pandas as pd

# Configuration
MONTH = sys.argv[1] if len(sys.argv) > 1 else "March"
TEMPLATE_PATH = "sample_data/final-data/06-01-2026.pptx"
GRAPH_FOLDER = "sample_data/Graph"
PPT_FOLDER = "sample_data/ppt-data"
OUTPUT_PATH = "output/consolidated_mbr.pptx"

print("\n" + "="*80)
print(f"🎯 BUILDING COMPLETE {MONTH.upper()} 2026 MBR REPORT")
print("="*80 + "\n")

# Import functions from create_mbr_complete.py
sys.path.insert(0, '.')
from create_mbr_complete import (
    get_all_projects,
    get_lt_manager_from_template_slide3,
    get_project_manager_from_excel,
    get_goals_and_progress_for_project,
    update_slide_1_date,
    update_slide_3_table
)

def create_project_metrics_slide(prs, project_name, achievements, chart_instructions):
    """Create a Project Metrics slide."""
    blank_layout = prs.slide_layouts[6]  # Blank layout
    slide = prs.slides.add_slide(blank_layout)
    
    # Title
    title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(9), Inches(0.6))
    title_frame = title_box.text_frame
    title_frame.text = f"Project Metrics - {project_name}"
    title_para = title_frame.paragraphs[0]
    title_para.font.size = Pt(32)
    title_para.font.bold = True
    
    # Chart placeholder 1 - Rework Trend (left side)
    chart1_box = slide.shapes.add_textbox(Inches(0.5), Inches(1.2), Inches(4.3), Inches(3))
    chart1_frame = chart1_box.text_frame
    chart1_frame.text = f"[CHART: Rework Trend]\n\n{chart_instructions['rework']}"
    for para in chart1_frame.paragraphs:
        para.font.size = Pt(11)
        para.font.italic = True
    
    # Chart placeholder 2 - Throughput Distribution (right side)
    chart2_box = slide.shapes.add_textbox(Inches(5.2), Inches(1.2), Inches(4.3), Inches(3))
    chart2_frame = chart2_box.text_frame
    chart2_frame.text = f"[CHART: Throughput]\n\n{chart_instructions['throughput']}"
    for para in chart2_frame.paragraphs:
        para.font.size = Pt(11)
        para.font.italic = True
    
    # Key Achievements section
    achievements_box = slide.shapes.add_textbox(Inches(0.5), Inches(4.5), Inches(9), Inches(2.5))
    achievements_frame = achievements_box.text_frame
    achievements_frame.word_wrap = True
    
    # Title
    p = achievements_frame.paragraphs[0]
    p.text = "Key Achievements:"
    p.font.size = Pt(18)
    p.font.bold = True
    
    # List achievements
    for achievement in achievements[:7]:  # Max 7
        p = achievements_frame.add_paragraph()
        p.text = f"• {achievement}"
        p.font.size = Pt(13)
        p.level = 0
    
    return slide

def main():
    """Main execution."""
    
    # Step 1: Load template
    print("📋 Step 1: Loading template...")
    template_prs = Presentation(TEMPLATE_PATH)
    print(f"   Loaded {len(template_prs.slides)} slides\n")
    
    # Step 2: Get projects
    print("📋 Step 2: Finding projects...")
    projects = get_all_projects()
    print(f"   Found {len(projects)} projects:")
    for p in projects:
        print(f"      • {p['name']}")
    print("")
    
    # Step 3: Update Slide 1 (Date)
    print("📋 Step 3: Updating Slide 1 date...")
    update_slide_1_date(template_prs.slides[0], MONTH)
    print("")
    
    # Step 4: Update Slide 3 (Table)
    print("📋 Step 4: Updating Slide 3 table...")
    update_slide_3_table(template_prs.slides[2], projects, template_prs)
    print("")
    
    # Step 5: Delete old Project Metrics slides (7-13)
    print("📋 Step 5: Removing old Project Metrics slides...")
    slides_to_delete = list(range(6, 13))  # Slides 7-13 (0-indexed: 6-12)
    for idx in sorted(slides_to_delete, reverse=True):
        if idx < len(template_prs.slides._sldIdLst):
            rId = template_prs.slides._sldIdLst[idx].rId
            template_prs.part.drop_rel(rId)
            del template_prs.slides._sldIdLst[idx]
            print(f"   ✓ Deleted slide {idx+1}")
    print("")
    
    # Step 6: Create new Project Metrics slides
    print("📋 Step 6: Creating Project Metrics slides...")
    
    for project in projects:
        print(f"   📊 {project['name']}...")
        
        # Get achievements
        goals, progress = get_goals_and_progress_for_project(project['name'])
        achievements = goals if goals else ["No achievements documented"]
        
        # Chart instructions
        excel_filename = Path(project['excel_path']).name
        chart_instructions = {
            'rework': f"Excel: {excel_filename}\nChart: 'Rework Trend'",
            'throughput': f"Excel: {excel_filename}\nChart: 'Throughput Distribution'"
        }
        
        # Create slide
        slide = create_project_metrics_slide(
            template_prs,
            project['display_name'],
            achievements,
            chart_instructions
        )
        
        print(f"      ✓ Created with {len(achievements)} achievements")
    
    print("")
    
    # Step 7: Save
    print("💾 Step 7: Saving...")
    template_prs.save(OUTPUT_PATH)
    
    print("\n" + "="*80)
    print("✅ SUCCESS!")
    print("="*80)
    print(f"\n📁 Saved: {OUTPUT_PATH}")
    print(f"📊 Total slides: {len(template_prs.slides)}")
    print(f"\nStructure:")
    print(f"  Slides 1-6: Template (updated)")
    print(f"  Slides 7-{6+len(projects)}: Project Metrics (one per project)")
    print(f"  Slide {7+len(projects)}: Thank You")
    print(f"\n⚠️  NOTE: Chart placeholders created. See next steps below.\n")

if __name__ == "__main__":
    main()
