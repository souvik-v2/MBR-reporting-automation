"""
Create March MBR Report
"""
from pptx_auto_consolidator import consolidate_pptx_com

print("\n" + "="*80)
print("Creating March 2026 MBR Report")
print("="*80 + "\n")

print("⚠️  IMPORTANT: Please close any open PowerPoint files before continuing!")
input("Press Enter when ready to continue...")

try:
    result = consolidate_pptx_com(
        template_path="sample_data/final-data/06-01-2026.pptx",
        ppt_data_folder="sample_data/ppt-data",
        output_path="output/consolidated_mbr_march.pptx",
        month="March",
        insert_after_slide=5
    )
    print(f"\n🎉 SUCCESS! Open the file: {result}")
    print("\nThis file contains:")
    print("  • Template slides (14)")
    print("  • Goals Achieved in March slides from ALL teams")
    print("  • Total slides: Template + Goal slides")
    
except Exception as e:
    print(f"\n❌ Failed: {e}")
    print("\nTroubleshooting:")
    print("1. Close ALL PowerPoint windows")
    print("2. Close the SEOD PowerPoint file if it's open")
    print("3. Run this script again")
