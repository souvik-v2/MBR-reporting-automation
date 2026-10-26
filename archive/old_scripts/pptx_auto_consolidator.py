"""
Automated PowerPoint Consolidator using Windows COM (PowerPoint Application).

This uses Microsoft PowerPoint's COM interface to copy slides programmatically,
preserving all formatting, images, charts, and animations.

Requires: Microsoft PowerPoint installed on Windows
"""
from __future__ import annotations
import sys
from pathlib import Path
from typing import List, Tuple
import time

try:
    import win32com.client
except ImportError:
    print("ERROR: pywin32 package required for automated PowerPoint consolidation")
    print("Install with: py -m pip install pywin32")
    sys.exit(1)


def find_slides_by_keywords_com(ppt_path: str, keywords: List[str]) -> List[Tuple[int, str]]:
    """Find slides matching any of the given keywords using COM."""
    matching_slides = []

    try:
        powerpoint = win32com.client.Dispatch("PowerPoint.Application")
        powerpoint.Visible = 1  # Make visible for debugging

        presentation = powerpoint.Presentations.Open(str(Path(ppt_path).absolute()), WithWindow=False)

        for i, slide in enumerate(presentation.Slides, 1):
            try:
                if slide.Shapes.HasTitle:
                    title = slide.Shapes.Title.TextFrame.TextRange.Text
                    # Check if any keyword matches
                    for keyword in keywords:
                        if keyword.lower() in title.lower():
                            matching_slides.append((i, title))
                            break
            except:
                pass

        presentation.Close()

    except Exception as e:
        print(f"Error scanning {ppt_path}: {e}")

    return matching_slides


def find_goal_slides_com(ppt_path: str, month: str = "May") -> List[Tuple[int, str]]:
    """Find slides with 'Goals Achieved in {month}' using COM."""
    keywords = [
        f"Goals Achieved in {month}",
        f"Goals Achieved",
        f"{month} Goals",
    ]
    return find_slides_by_keywords_com(ppt_path, keywords)


def consolidate_pptx_com(
    template_path: str,
    ppt_data_folder: str,
    output_path: str,
    month: str = "May",
    insert_after_slide: int = 5
) -> str:
    """
    Consolidate PowerPoint presentations using COM automation.

    This preserves all formatting, images, charts, and animations.

    Args:
        template_path: Path to template PPTX (optional - can be None to create new)
        ppt_data_folder: Folder containing team presentations
        output_path: Where to save consolidated PPTX
        month: Month to search for in goal slides
        insert_after_slide: Insert goal slides after this position (ignored if no template)

    Returns:
        Path to consolidated presentation
    """
    print("\n" + "=" * 80)
    print("🚀 AUTOMATED POWERPOINT CONSOLIDATION (Using COM)")
    print("=" * 80)
    print(f"\nTemplate: {template_path if template_path else 'None (creating new)'}")
    print(f"Data folder: {ppt_data_folder}")
    print(f"Output: {output_path}")
    print(f"Month: {month}")
    if template_path:
        print(f"Insert after slide: {insert_after_slide}")
    print()

    try:
        # Initialize PowerPoint
        powerpoint = win32com.client.Dispatch("PowerPoint.Application")
        powerpoint.Visible = 1  # Show PowerPoint during processing

        # Check if template exists
        use_template = template_path and Path(template_path).exists()

        if use_template:
            # Open template
            print(f"📂 Opening template...")
            template = powerpoint.Presentations.Open(str(Path(template_path).absolute()))
            initial_slide_count = template.Slides.Count
            print(f"   Template has {initial_slide_count} slides")
        else:
            # Create new presentation
            if template_path and not Path(template_path).exists():
                print(f"⚠️  Template not found: {template_path}")
            print(f"📄 Creating new presentation...")
            template = powerpoint.Presentations.Add()
            # Remove default slide
            if template.Slides.Count > 0:
                template.Slides(1).Delete()
            initial_slide_count = 0
            insert_after_slide = 0
            print(f"   Created blank presentation")
        
        # Find all team presentations
        ppt_folder = Path(ppt_data_folder)
        team_presentations = list(ppt_folder.glob("*.pptx"))
        team_presentations = [p for p in team_presentations if not p.name.startswith('~$')]
        
        print(f"\n🔍 Scanning {len(team_presentations)} presentations for '{month}' goals...")
        
        slides_added = 0
        insert_position = insert_after_slide
        
        for ppt_file in team_presentations:
            print(f"\n   📄 {ppt_file.name}")
            goal_slides = find_goal_slides_com(str(ppt_file), month)
            
            if goal_slides:
                # Open source presentation
                source_prs = powerpoint.Presentations.Open(str(ppt_file.absolute()), WithWindow=False)
                
                for slide_num, title in goal_slides:
                    print(f"      ✓ Copying slide #{slide_num}: {title}")
                    
                    # Copy slide from source
                    source_slide = source_prs.Slides(slide_num)
                    source_slide.Copy()
                    
                    # Paste into template at specific position
                    insert_position += 1
                    template.Slides.Paste(insert_position)
                    
                    slides_added += 1
                    time.sleep(0.5)  # Brief pause for stability
                
                source_prs.Close()
        
        # Save consolidated presentation
        print(f"\n💾 Saving consolidated presentation...")
        output_full_path = str(Path(output_path).absolute())

        # Remove existing file if it exists (retry logic for locked files)
        if Path(output_path).exists():
            try:
                # Close the template first to release file lock
                template.Close()
                time.sleep(0.5)
                Path(output_path).unlink()
                time.sleep(0.5)
                # Reopen to save
                if use_template:
                    template = powerpoint.Presentations.Open(str(Path(template_path).absolute()))
                else:
                    # Recreate for new presentation
                    print("  ⚠️  Cannot overwrite existing file while PowerPoint is open")
                    print(f"  ℹ️  Saving as: {output_full_path.replace('.pptx', '_new.pptx')}")
                    output_full_path = output_full_path.replace('.pptx', '_new.pptx')
            except Exception as e:
                print(f"  ⚠️  Could not remove existing file: {e}")
                print(f"  ℹ️  Saving as: {output_full_path.replace('.pptx', '_new.pptx')}")
                output_full_path = output_full_path.replace('.pptx', '_new.pptx')

        template.SaveAs(output_full_path)
        template.Close()

        # Close PowerPoint
        powerpoint.Quit()
        
        print(f"\n" + "=" * 80)
        print("✅ SUCCESS!")
        print("=" * 80)
        print(f"\n📊 Summary:")
        print(f"   Original template slides: {initial_slide_count}")
        print(f"   Goal slides added: {slides_added}")
        print(f"   Final presentation slides: {initial_slide_count + slides_added}")
        print(f"\n📁 Saved to: {output_path}")
        print("\n" + "=" * 80)
        
        return output_path
        
    except Exception as e:
        print(f"\n❌ ERROR: {e}")
        print("\nTroubleshooting:")
        print("1. Make sure Microsoft PowerPoint is installed")
        print("2. Install pywin32: py -m pip install pywin32")
        print("3. Close any open PowerPoint windows and try again")
        raise


if __name__ == "__main__":
    # Test automated consolidation
    try:
        result = consolidate_pptx_com(
            template_path="sample_data/final-data/06-01-2026.pptx",
            ppt_data_folder="sample_data/ppt-data",
            output_path="output/consolidated_mbr.pptx",
            month="March",
            insert_after_slide=5
        )
        print(f"\n🎉 Done! Open the file: {result}")
    except Exception as e:
        print(f"\n❌ Failed: {e}")
