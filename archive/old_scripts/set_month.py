"""
Quick script to change the default month for MBR workflow.

Usage:
    py set_month.py March
    py set_month.py June
"""
import sys
from pathlib import Path

def update_month_in_file(file_path: str, old_month: str, new_month: str):
    """Update month in a single file."""
    path = Path(file_path)
    if not path.exists():
        print(f"⚠️  File not found: {file_path}")
        return False
    
    content = path.read_text(encoding='utf-8')
    
    # Replace month value
    patterns = [
        f"month: str = \"{old_month}\"",
        f"'month': '{old_month}'",
        f"\"month\": \"{old_month}\""
    ]
    
    updated = False
    for pattern in patterns:
        if pattern in content:
            new_pattern = pattern.replace(old_month, new_month)
            content = content.replace(pattern, new_pattern)
            updated = True
    
    if updated:
        path.write_text(content, encoding='utf-8')
        print(f"✅ Updated: {file_path}")
        return True
    else:
        print(f"ℹ️  No changes needed: {file_path}")
        return False

def main():
    if len(sys.argv) < 2:
        print("Usage: py set_month.py <month_name>")
        print("\nExamples:")
        print("  py set_month.py March")
        print("  py set_month.py June")
        print("  py set_month.py December")
        return
    
    new_month = sys.argv[1].capitalize()
    
    print(f"\n🔄 Changing default month to: {new_month}\n")
    
    # Try to detect current month
    current_months = []
    for file in ["mbr_workflow.py", "mbr_chatbot.py", "mbr_assistant.py"]:
        if Path(file).exists():
            content = Path(file).read_text(encoding='utf-8')
            for month in ["January", "February", "March", "April", "May", "June", 
                         "July", "August", "September", "October", "November", "December"]:
                if f'"{month}"' in content or f"'{month}'" in content:
                    current_months.append(month)
                    break
    
    old_month = current_months[0] if current_months else "May"
    print(f"Current month detected: {old_month}\n")
    
    # Update all files
    files = [
        "mbr_workflow.py",
        "mbr_chatbot.py",
        "mbr_assistant.py"
    ]
    
    updated_count = 0
    for file in files:
        if update_month_in_file(file, old_month, new_month):
            updated_count += 1
    
    print(f"\n{'=' * 60}")
    print(f"✅ Updated {updated_count} file(s) to use month: {new_month}")
    print(f"{'=' * 60}\n")
    print("Now you can run:")
    print(f"  py main.py --mbr-workflow")
    print("\nOr use the chatbot:")
    print(f"  py mbr_assistant.py")
    print()

if __name__ == "__main__":
    main()
