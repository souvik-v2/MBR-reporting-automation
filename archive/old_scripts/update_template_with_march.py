"""
Update template PowerPoint with March data from team presentations.

This script:
1. Copies the template as base
2. Inserts "Goals Achieved in March" slides after slide 5
3. Inserts "In Progress Ticket" slides
4. Keeps template formatting
"""
from pptx import Presentation
from pathlib import Path
from copy import deepcopy
import shutil

def copy_slide_to_presentation(source_prs, source_slide_idx, dest_prs, insert_position):
    """Copy a slide from source to destination at specific position."""
    try:
        source_slide = source_prs.slides[source_slide_idx - 1]
        
        # Use blank layout
        blank_layout = dest_prs.slide_layouts[6]
        
        # Insert slide at position
        dest_slide = dest_prs.slides.add_slide(blank_layout)
        
        # Move to correct position
        dest_slides = list(dest_prs.slides._sldIdLst)
        dest_slides.insert(insert_position, dest_slides.pop())
        
        # Copy all shapes
        for shape in source_slide.shapes:
            el = shape.element
            newel = deepcopy(el)
            dest_slide.shapes._spTree.insert_element_before(newel, 'p:extLst')
        
        return dest_slide
    except Exception as e:
        print(f"      ⚠️ Error: {e}")
        return None

def find_team_presentation(team_name, ppt_folder):
    """Find presentation file for a team."""
    ppt_path = Path(ppt_folder)
    
    # Mapping of template team names to file patterns
    patterns = {
        "SEOD": "*SEOD*March*.pptx",
        "Consumer Web": "*Core*SVC*March*.pptx",
        "CET": "*Core*SVC*March*.pptx",  # Same as Consumer Web
        "WPM": "*WPM*March*.pptx",
        "Phoenix": "*Phoenix*March*.pptx",
        "SLE": "*CanopyNext*March*.pptx",
        "FLEX": "*FLEX*March*.pptx",
    }
    
    pattern = patterns.get(team_name, f"*{team_name}*March*.pptx")
    files = list(ppt_path.glob(pattern))
    files = [f for f in files if not f.name.startswith('~$')]
    
    return str(files[0]) if files else None

def find_slide_by_title(prs, keyword):
    """Find slide index by title keyword."""
    for i, slide in enumerate(prs.slides, 1):
        if slide.shapes.title:
            title = slide.shapes.title.text
            if keyword.lower() in title.lower():
                return i
    return None

def main():
    print("\n" + "="*80)
    print("🎯 UPDATING TEMPLATE WITH MARCH DATA")
    print("="*80 + "\n")
    
    template_path = "sample_data/final-data/06-01-2026.pptx"
    output_path = "output/consolidated_mbr.pptx"
    ppt_data_folder = "sample_data/ppt-data"
    
    print(f"📂 Template: {template_path}")
    print(f"📂 Data folder: {ppt_data_folder}")
    print(f"📁 Output: {output_path}\n")
    
    # Load template
    print("📋 Loading template...")
    template_prs = Presentation(template_path)
    print(f"   Template has {len(template_prs.slides)} slides\n")
    
    # Find all team presentations
    ppt_path = Path(ppt_data_folder)
    team_pptx = list(ppt_path.glob("*March*.pptx"))
    team_pptx = [p for p in team_pptx if not p.name.startswith('~$')]
    
    print(f"📊 Found {len(team_pptx)} March presentations\n")
    
    # Collect "Goals Achieved" and "In Progress" slides
    print("🔍 Finding slides to insert...\n")
    
    slides_to_insert = []
    
    for ppt_file in sorted(team_pptx):
        team_name = ppt_file.stem.replace("V2Solutions_", "").replace("LT_", "").replace("_March_2026_Report", "").replace("_March2026_Report", "").replace("_March_Report", "").replace("_Mar_2026_Report 1", "")
        
        print(f"📄 {team_name}")
        
        try:
            source_prs = Presentation(str(ppt_file))
            
            # Find "Goals Achieved in March"
            goals_idx = find_slide_by_title(source_prs, "Goals Achieved")
            if goals_idx:
                slides_to_insert.append((str(ppt_file), goals_idx, f"{team_name} - Goals", source_prs))
                print(f"   ✓ Goals Achieved (Slide {goals_idx})")
            
            # Find "In Progress"
            progress_idx = find_slide_by_title(source_prs, "In Progress")
            if progress_idx:
                slides_to_insert.append((str(ppt_file), progress_idx, f"{team_name} - In Progress", source_prs))
                print(f"   ✓ In Progress (Slide {progress_idx})")
            
            if not goals_idx and not progress_idx:
                print(f"   ℹ️  No matching slides")
                
        except Exception as e:
            print(f"   ⚠️ Error: {e}")
    
    print(f"\n📊 Total slides to insert: {len(slides_to_insert)}\n")
    
    # Insert slides after slide 5 (Suggestions | Recommendations)
    print("🚀 Inserting slides into template...\n")
    
    insert_position = 5  # After "Suggestions | Recommendations"
    
    for ppt_path, slide_idx, description, source_prs in slides_to_insert:
        print(f"   📋 Inserting: {description}")
        
        result = copy_slide_to_presentation(source_prs, slide_idx, template_prs, insert_position)
        
        if result:
            print(f"      ✅ Inserted at position {insert_position + 1}")
            insert_position += 1
        else:
            print(f"      ❌ Failed")
    
    # Save
    print(f"\n💾 Saving consolidated presentation...")
    template_prs.save(output_path)
    
    print("\n" + "="*80)
    print("✅ SUCCESS!")
    print("="*80)
    print(f"\n📁 Saved: {output_path}")
    print(f"📊 Total slides: {len(template_prs.slides)}")
    print(f"\nStructure:")
    print(f"  Slides 1-5: Template overview")
    print(f"  Slides 6-{5+len(slides_to_insert)}: Goals & In-Progress (March)")
    print(f"  Remaining: Project Metrics + Thank You")
    print()

if __name__ == "__main__":
    main()
