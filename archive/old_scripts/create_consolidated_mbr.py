"""
Create Consolidated MBR PowerPoint from Team Presentations.

This script:
1. Scans all team presentations in ppt-data folder
2. Extracts "Goals Achieved" slides
3. Extracts "In Progress Ticket" slides
4. Creates a consolidated presentation with all slides
"""
from __future__ import annotations
import sys
from pathlib import Path
from typing import List, Tuple, Dict
import time

try:
    import win32com.client
except ImportError:
    print("ERROR: pywin32 package required")
    print("Install with: py -m pip install pywin32")
    sys.exit(1)


def find_matching_slides(ppt_path: str, keywords: List[str], powerpoint_app=None) -> List[Tuple[int, str]]:
    """Find slides matching any of the given keywords."""
    matching_slides = []
    close_app = False

    try:
        if powerpoint_app is None:
            powerpoint = win32com.client.Dispatch("PowerPoint.Application")
            powerpoint.Visible = 1
            close_app = True
        else:
            powerpoint = powerpoint_app

        presentation = powerpoint.Presentations.Open(str(Path(ppt_path).absolute()), WithWindow=False)

        for i, slide in enumerate(presentation.Slides, 1):
            try:
                if slide.Shapes.HasTitle:
                    title = slide.Shapes.Title.TextFrame.TextRange.Text
                    for keyword in keywords:
                        if keyword.lower() in title.lower():
                            matching_slides.append((i, title))
                            break
            except:
                pass

        presentation.Close()

        if close_app:
            powerpoint.Quit()

    except Exception as e:
        print(f"   ⚠️  Error: {e}")

    return matching_slides


def create_consolidated_presentation(
    ppt_data_folder: str,
    output_path: str,
    month: str = "March"
):
    """Create consolidated presentation from team presentations."""
    
    print("\n" + "=" * 80)
    print("🎯 CREATING CONSOLIDATED MBR PRESENTATION")
    print("=" * 80)
    print(f"\nMonth: {month}")
    print(f"Source folder: {ppt_data_folder}")
    print(f"Output: {output_path}\n")
    
    # Keywords to search for
    goal_keywords = [f"Goals Achieved in {month}", "Goals Achieved", f"{month} Goals"]
    progress_keywords = ["In Progress", "In-Progress", "Progress Ticket", "In Progress Ticket"]
    
    # Find all presentations
    ppt_folder = Path(ppt_data_folder)
    team_presentations = list(ppt_folder.glob("*.pptx"))
    team_presentations = [p for p in team_presentations if not p.name.startswith('~$')]
    
    print(f"📂 Found {len(team_presentations)} presentations\n")

    # Initialize PowerPoint once for all scanning
    try:
        powerpoint = win32com.client.Dispatch("PowerPoint.Application")
        powerpoint.Visible = 1
    except:
        print("❌ Could not start PowerPoint")
        return

    # Scan all presentations
    slides_to_copy = []

    for ppt_file in sorted(team_presentations):
        team_name = ppt_file.stem.replace("V2Solutions_", "").replace("_March_2026_Report", "").replace("_March2026_Report", "").replace("_March_Report", "").replace("_Mar_2026_Report 1", "")
        print(f"📄 {team_name}")

        # Find goal slides
        goal_slides = find_matching_slides(str(ppt_file), goal_keywords, powerpoint)
        for slide_num, title in goal_slides:
            slides_to_copy.append((str(ppt_file), slide_num, title, team_name, "Goals"))
            print(f"   ✓ Slide {slide_num}: {title}")

        # Find in-progress slides
        progress_slides = find_matching_slides(str(ppt_file), progress_keywords, powerpoint)
        for slide_num, title in progress_slides:
            slides_to_copy.append((str(ppt_file), slide_num, title, team_name, "In-Progress"))
            print(f"   ✓ Slide {slide_num}: {title}")

        if not goal_slides and not progress_slides:
            print(f"   ℹ️  No matching slides found")
    
    print(f"\n📊 Total slides to consolidate: {len(slides_to_copy)}\n")
    
    if len(slides_to_copy) == 0:
        print("⚠️  No slides found to consolidate!")
        return
    
    # Create consolidated presentation
    print("🚀 Creating consolidated presentation...")

    try:
        # Create new presentation
        consolidated = powerpoint.Presentations.Add()
        
        # Remove default slide
        if consolidated.Slides.Count > 0:
            consolidated.Slides(1).Delete()
        
        # Copy all slides
        slide_position = 0
        for ppt_path, slide_num, title, team_name, slide_type in slides_to_copy:
            print(f"   📋 Copying: {team_name} - {title}")
            
            try:
                source_prs = powerpoint.Presentations.Open(ppt_path, WithWindow=False)
                source_slide = source_prs.Slides(slide_num)
                source_slide.Copy()
                
                slide_position += 1
                consolidated.Slides.Paste(slide_position)
                
                source_prs.Close()
                time.sleep(0.3)
                
            except Exception as e:
                print(f"      ⚠️  Error copying slide: {e}")
        
        # Save consolidated presentation
        print(f"\n💾 Saving consolidated presentation...")
        output_full_path = str(Path(output_path).absolute())
        
        # Remove existing file if it exists
        if Path(output_path).exists():
            try:
                Path(output_path).unlink()
            except:
                print(f"   ⚠️  Creating new file name...")
                output_full_path = output_full_path.replace('.pptx', '_new.pptx')
        
        consolidated.SaveAs(output_full_path)
        consolidated.Close()
        powerpoint.Quit()
        
        print("\n" + "=" * 80)
        print("✅ SUCCESS!")
        print("=" * 80)
        print(f"\n📁 Saved to: {output_path}")
        print(f"📊 Total slides: {len(slides_to_copy)}")
        print()
        
    except Exception as e:
        print(f"\n❌ ERROR: {e}")
        raise


if __name__ == "__main__":
    import sys

    # Get month from command line or use default
    month = sys.argv[1] if len(sys.argv) > 1 else "March"

    print("\n⚠️  NOTE: This will work best if PowerPoint files are closed.")
    print("         Continuing in 2 seconds...")
    time.sleep(2)

    create_consolidated_presentation(
        ppt_data_folder="sample_data/ppt-data",
        output_path="output/consolidated_mbr.pptx",
        month=month
    )
