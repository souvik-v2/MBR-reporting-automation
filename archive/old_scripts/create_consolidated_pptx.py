"""
Create consolidated PowerPoint using python-pptx library.
This approach directly copies slide XML to preserve formatting.
"""
from pptx import Presentation
from pathlib import Path
from copy import deepcopy
import sys

def find_matching_slides(ppt_path, keywords):
    """Find slides matching keywords."""
    matching = []
    try:
        prs = Presentation(ppt_path)
        for i, slide in enumerate(prs.slides, 1):
            if slide.shapes.title:
                title = slide.shapes.title.text
                for keyword in keywords:
                    if keyword.lower() in title.lower():
                        matching.append((i, title, slide))
                        break
    except Exception as e:
        print(f"   ⚠️ Error reading {Path(ppt_path).name}: {e}")
    return matching

def copy_slide(source_prs, slide_index, dest_prs):
    """Copy a slide from source to destination presentation."""
    try:
        source_slide = source_prs.slides[slide_index - 1]
        
        # Get the slide layout from source
        blank_slide_layout = dest_prs.slide_layouts[6]  # Blank layout
        
        # Add new slide
        dest_slide = dest_prs.slides.add_slide(blank_slide_layout)
        
        # Copy all shapes from source to destination
        for shape in source_slide.shapes:
            el = shape.element
            newel = deepcopy(el)
            dest_slide.shapes._spTree.insert_element_before(newel, 'p:extLst')
        
        return True
    except Exception as e:
        print(f"      ⚠️ Copy error: {e}")
        return False

def create_consolidated(ppt_folder, output_path, month="March"):
    """Create consolidated presentation."""
    print("\n" + "="*80)
    print("🎯 CREATING CONSOLIDATED MBR PRESENTATION")
    print("="*80)
    print(f"\nMonth: {month}")
    print(f"Source: {ppt_folder}")
    print(f"Output: {output_path}\n")
    
    # Find all presentations
    ppt_path = Path(ppt_folder)
    ppt_files = list(ppt_path.glob("*.pptx"))
    ppt_files = [p for p in ppt_files if not p.name.startswith('~$')]
    
    print(f"📂 Found {len(ppt_files)} presentations\n")
    
    # Keywords to search
    goal_keywords = [f"Goals Achieved in {month}", "Goals Achieved", f"{month} Goals"]
    progress_keywords = ["In Progress", "In-Progress", "Progress Ticket", "In Progress Ticket"]
    
    # Collect all slides to copy
    slides_info = []
    
    for ppt_file in sorted(ppt_files):
        team_name = ppt_file.stem.replace("V2Solutions_", "").replace("LT_", "").replace("_March_2026_Report", "").replace("_March2026_Report", "").replace("_March_Report", "").replace("_Mar_2026_Report 1", "")
        print(f"📄 {team_name}")
        
        # Find goals
        goals = find_matching_slides(str(ppt_file), goal_keywords)
        for idx, title, slide in goals:
            slides_info.append((str(ppt_file), idx, title, team_name, "Goals", slide))
            print(f"   ✓ Slide {idx}: {title}")
        
        # Find in-progress
        progress = find_matching_slides(str(ppt_file), progress_keywords)
        for idx, title, slide in progress:
            slides_info.append((str(ppt_file), idx, title, team_name, "In-Progress", slide))
            print(f"   ✓ Slide {idx}: {title}")
        
        if not goals and not progress:
            print(f"   ℹ️  No matching slides")
    
    print(f"\n📊 Total: {len(slides_info)} slides\n")
    
    if len(slides_info) == 0:
        print("⚠️  No slides to consolidate!")
        return
    
    # Create consolidated presentation
    print("🚀 Creating consolidated presentation...\n")
    
    # Create new presentation
    dest_prs = Presentation()
    
    # Remove default slide
    if len(dest_prs.slides) > 0:
        rId = dest_prs.slides._sldIdLst[0].rId
        dest_prs.part.drop_rel(rId)
        del dest_prs.slides._sldIdLst[0]
    
    # Copy each slide
    for ppt_path, slide_idx, title, team, slide_type, _ in slides_info:
        print(f"   📋 {team} - {title}")
        
        # Open source presentation
        source_prs = Presentation(ppt_path)
        
        # Copy the slide
        if copy_slide(source_prs, slide_idx, dest_prs):
            print(f"      ✅ Copied successfully")
        else:
            print(f"      ❌ Failed to copy")
    
    # Save
    print(f"\n💾 Saving...")
    dest_prs.save(output_path)
    
    print("\n" + "="*80)
    print("✅ SUCCESS!")
    print("="*80)
    print(f"\n📁 Saved: {output_path}")
    print(f"📊 Slides: {len(dest_prs.slides)}\n")

if __name__ == "__main__":
    month = sys.argv[1] if len(sys.argv) > 1 else "March"
    
    create_consolidated(
        ppt_folder="sample_data/ppt-data",
        output_path="output/consolidated_mbr.pptx",
        month=month
    )
